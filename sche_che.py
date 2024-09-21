from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from PyQt5.QtCore import QObject, pyqtSignal
import argparse
from tqdm import tqdm

import pandas as pd
from db_models import Class, Schedule, Teacher, TeacherSchedule, Cabinet
from database import get_db

def load_data_from_excel(file_path):
    # Пример загрузки данных из Excel файла
    df = pd.read_excel(file_path)
    db = next(get_db())

    for index, row in df.iterrows():
        class_ = Class(name=row['class_name'])
        db.add(class_)
        db.commit()

        schedule = Schedule(class_id=class_.id, day=row['day'], lesson_number=row['lesson_number'], subject=row['subject'], cabinet=row['cabinet'])
        db.add(schedule)
        db.commit()

    db.close()

def export_data_to_excel(file_path):
    # Пример экспорта данных в Excel файл
    db = next(get_db())
    schedules = db.query(Schedule).all()
    df = pd.DataFrame([schedule.__dict__ for schedule in schedules])
    df.to_excel(file_path, index=False)
    db.close()







class FilePreparator(QObject):
    preparation_progress = pyqtSignal(int)


    def row_normalization(self, wb):
        # будем приводить в единый вид все файлы
        # просто делаю каждую строку урока сдвоенной
        # для простоты будем задваивать все строки.

        ws = wb.active
        wb_out = Workbook()
        ws_out = wb_out.active

        rows_with_merged_cells = set()
        for merge_range in ws.merged_cells.ranges:
            if merge_range.min_row == merge_range.max_row - 1 and merge_range.min_col == merge_range.max_col:
                rows_with_merged_cells.add(merge_range.min_row)

        row = 1
        total_rows = ws.max_row - len(rows_with_merged_cells)
        pbar = tqdm(total=total_rows)
        while row < ws.max_row:
            if not row in rows_with_merged_cells:
                ws_out.append([cell.value for cell in ws[row]])
                ws_out.append([None for cell in ws[row]])
                for col in range(1, len(ws[row]) + 1):
                    ws_out.merge_cells(start_row=ws_out.max_row - 1, start_column=col, end_row=ws_out.max_row,
                                       end_column=col)
                # есть непрокрасы. не стал разбираться
                # for col in range(1, len(ws[row]) + 1):
                #     cell_in = ws.cell(row=row, column=col)
                #     cell_out = ws_out.cell(row=ws_out.max_row - 1, column=col)
                #     cell_out.value = cell_in.value
                #     cell_out.font = copy(cell_in.font)
                #     cell_out.fill = copy(cell_in.fill)
                #     cell_out.border = copy(cell_in.border)
            else:
                ws_out.append([cell.value for cell in ws[row]])
                row += 1
                ws_out.append([cell.value for cell in ws[row]])
                for col in range(1, len(ws[row]) + 1):
                    if ws_out.cell(ws_out.max_row, col).value is None:
                        ws_out.merge_cells(start_row=ws_out.max_row - 1, start_column=col, end_row=ws_out.max_row,
                                           end_column=col)
            self.preparation_progress.emit(row)
            row += 1
            pbar.update(1)
        pbar.close()
        # постобработка
        # мешаем эксельному движку опять все превратить в дата/время
        for row in range(1, ws_out.max_row + 1):
            if '#' in str(ws_out.cell(row, 1).value):
                for col in range(1,6):
                    ws_out.merge_cells(start_row=row, start_column=col * 2, end_row=row, end_column=col * 2 + 1)
                ws_out.unmerge_cells(start_row=row - 2, start_column=1, end_row=row - 1 , end_column=1)
                ws_out.unmerge_cells(start_row=row - 2, start_column=2, end_row=row - 1 , end_column=2)
                ws_out.merge_cells(start_row=row - 2, start_column=1, end_row=row - 1 , end_column=2)
            if ':' in str(ws_out.cell(row, 1).value):
                ws_out.cell(row, 1).value = str(ws_out.cell(row, 1).value)[1:-3]

        return wb_out

    def row_normalization_single_line(self, wb):
        # ае! я поборол это дело без дублирования строк!

        # тащемта у нас всего два проблемных случая спаренной по вертикали строки
        # это когда шапка класса и собственно урок в двух кабинетах
        ws = wb.active
        wb_out = Workbook()
        ws_out = wb_out.active

        merged_cells = ws.merged_cells.ranges
        row = 1
        row_out = 1

        pbar = tqdm(total=ws.max_row)
        while row < ws.max_row:
            # делаем шапку класса
            if ws.cell(row, 1).value and ws.cell(row, 1).value == '#':
                ws_out.append([None for x in range(11)])
                ws_out.append([None for x in range(11)])
                ws_out.merge_cells(start_row=row_out, start_column=1, end_row=row_out + 1, end_column=1)
                ws_out.cell(row_out, 1).value = '№'
                for i in range(5):
                    ws_out.merge_cells(start_row=row_out, start_column=2 + i * 2, end_row=row_out,
                                       end_column=2 + i * 2 + 1)
                    ws_out.cell(row_out, 2 + i * 2).value = ws.cell(row, 2 + i * 2).value
                    ws_out.cell(row_out + 1, 2 + i * 2).value = ws.cell(row + 1, 2 + i * 2).value
                    ws_out.cell(row_out + 1, 2 + i * 2 + 1).value = ws.cell(row + 1, 2 + i * 2 + 1).value
                row += 2
                row_out += 2
                pbar.update(2)
            # случай спаренной строки урока
            elif (any(ws.cell(row, 1).coordinate in range_str for range_str in merged_cells) and ws.cell(row, 1).value
                  and not ws.cell(row, 1).value == '#' and not 'Класс' in str(ws.cell(row, 1).value)):
                this_row = ws[row]
                next_row = ws[row + 1]
                for this_row_cell, next_row_cell in zip(this_row, next_row):
                    if next_row_cell.value is not None:
                        this_row_cell.value = f'{this_row_cell.value}\n{next_row_cell.value}'
                ws_out.append([cell.value for cell in this_row])
                if ':' in str(ws_out.cell(row_out, 1).value):
                    ws_out.cell(row_out, 1).value = str(ws_out.cell(row_out, 1).value)[:5]
                row += 2
                row_out += 1
                pbar.update(2)
            else:
                #все остальные случаи не содержат объединений на вертикали, так что пофиг
                ws_out.append([cell.value for cell in ws[row]])
                if ':' in str(ws_out.cell(row_out, 1).value):
                    ws_out.cell(row_out, 1).value = str(ws_out.cell(row_out, 1).value)[1:5]
                row += 1
                row_out += 1
                pbar.update(1)
        pbar.close()
        return wb_out


class DifferenceEngine(QObject):
    checking_progress = pyqtSignal(int)
    def bold_difference(self, old_wb, new_wb):
        dif_cell_font = Font(bold=True)
        old_ws = old_wb.active
        new_ws = new_wb.active
        if old_ws.max_row != new_ws.max_row:
            print('files length dont match!')
        limit = min(old_ws.max_row, new_ws.max_row)
        for row in range(1, limit + 1):
            for col in range(1, len(new_ws[row]) + 1):
                if new_ws.cell(row, col).value != old_ws.cell(row, col).value:
                    new_ws.cell(row, col).font = dif_cell_font


    def bold_difference_v2(self, old_wb, new_wb):
        dif_cell_font = Font(bold=True)
        old_ws = old_wb.active
        new_ws = new_wb.active
        row = 1
        old_row = 1
        while row < new_ws.max_row:
            if 'Класс' in str(new_ws.cell(row, 1).value):
                print(new_ws.cell(row, 1).value, row)
                flag = False
                for x in range(old_row, old_ws.max_row + 1):
                    if old_ws.cell(x, 1).value == new_ws.cell(row, 1).value:
                        old_row = x
                        flag = True
                        break
                if not flag:
                    print('No matches')
                    raise Exception
                cur_new_row = row + 2
                cur_old_row = old_row + 2
                while not (new_ws.cell(cur_new_row, 1).value is None and new_ws.cell(cur_new_row + 1, 1).value is None):
                    for col in range(1, len(new_ws[cur_new_row]) + 1):
                        if new_ws.cell(cur_new_row, col).value != old_ws.cell(cur_old_row, col).value:
                            if new_ws.cell(cur_new_row, col).value is None:
                                new_ws.cell(cur_new_row, col).value = '-окно-'
                            new_ws.cell(cur_new_row, col).font = dif_cell_font
                    cur_old_row += 1
                    cur_new_row += 1
                row = cur_new_row
            row += 1
        return new_wb


    def day_assemble(self, wb, day):
        res_wb = Workbook()
        res_ws = res_wb.active
        ws = wb.active


class SearchSystem(QObject):
    preparation_progress = pyqtSignal(int)

    def search_teacher_window_by_lesson_n(self, wb, teacher_name, day_n_0, lesson_n):
        # ws
        pass


def console_init():
    parser = argparse.ArgumentParser()
    parser.add_argument('-n', '--new_file', help='input filename')
    parser.add_argument('-o', '--old_file', help='input filename')
    args = parser.parse_args()
    new_wb = load_workbook(args.new_file)
    old_wb = load_workbook(args.old_file)

    print('normalizing previous schedule')
    old_wb = row_normalization(old_wb)
    print('normalizing new schedule')
    new_wb = row_normalization(new_wb)

    print('checking differences')
    bold_difference(old_wb, new_wb)
    print('done!')
    new_wb.save(f'{args.new_file.split(".")[0]}_checked.xlsx')


def create_common_teacher_schedule(school_wb):
    MAX_COL = 111
    wb_out = Workbook()
    ws_out = wb_out.active
    for i in range(8):
        ws_out.append([None for i in range(56)])
    ws_out.merge_cells(start_row=1, start_column=1, end_row=6, end_column=56)
    ws_out.cell(1, 1).value = 'Расписание уроков на 2024-2025'
    ws_out.merge_cells(start_row=7, start_column=1, end_row=8, end_column=1)
    ws_out.cell(7, 1).value = 'Ф.И.О.'
    days = ['понедельник', 'вторник', 'среда', 'четверг', 'пятница']
    for i in range(5):
        ws_out.merge_cells(start_row=7, start_column=2 + i * 11, end_row=7, end_column=2 + (i + 1) * 11 - 1)
        ws_out.cell(7, 2 + i * 11).value = days[i]
        for j in range(11):
            ws_out.cell(8, 2 + i * 11 + j).value = j + 1
    ws = school_wb.active

    row = 8
    pbar = tqdm(total=ws.max_row - row + 1)
    while row <= ws.max_row:
        cur_row = list()
        cur_row.append(str(ws.cell(row, 1).value))
        col = 2
        while col < MAX_COL:
            if ws.cell(row, col).value:
                this_class = str(ws.cell(row, col).value)
                if '_' in this_class:
                    # тут тогда Г Е И получаются в 11б
                    # res = list()
                    # for x in this_class.split(','):
                    #     y = x.split('_')
                    #     print(y, row, col)
                    #     res.append(y[0].strip() + y[1][0].upper())
                    # this_class = ', '.join(res)
                    this_class = this_class.split('_')[0]
                this_room = str(ws.cell(row, col + 1).value)
                if 'С' in this_room:
                    this_room = 'СЗ'
                if "П" in this_room:
                    this_room = 'П'
                cur_row.append('\n'.join((this_class, this_room)))
            else:
                cur_row.append(None)
            col += 2
        ws_out.append(cur_row)
        row += 1
        pbar.update(1)
    pbar.close()


    return wb_out


def create_common_pupils_schedule(normalized_wb):
    pass




if __name__ == '__main__':
    wb_in = load_workbook('учительское раписание.xlsx')
    wb_res = create_common_teacher_schedule(wb_in)
    wb_res.save('teachers_schedule.xlsx')