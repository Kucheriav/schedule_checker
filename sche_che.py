from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from PyQt5.QtCore import QObject, pyqtSignal
from openpyxl.utils import get_column_letter
from tqdm import tqdm
from copy import copy
from os import path, getcwd
import database
import db_models


CUR_SCHEDULES_FOLDER_PATH = path.join(getcwd(), 'current_version_schedules')
RESULT_FOLDER_PATH = path.join(getcwd(), 'changed_schedules')
DAYS = ['понедельник', 'вторник', 'среда', 'четверг', 'пятница']
CABINETS_WITH_EL_SCHOOL = ['101', '102', '103', '107', '201', '202', '203', '204', '205', '206', '207', '208', '209', '301', '302',
            '303', '304', '305', '306', '307', '308', '401', '402', '403', '404', '405', '406', '407', '408', '409',
            '411', '412', 'Акт.зал', 'СЗ', 'СЗ', 'СЗ', 'П']
CABINETS = ['101', '102', '107', '206', '208', '209', '301', '302', '303', '304', '305', '306', '307', '308', '401', '402',
            '403', '404', '405', '406', '407', '408', '409', '411', '412',  'СЗ', 'СЗ', 'СЗ', 'П']
CABINETS_SET = set(CABINETS)

class FuncToolBox(QObject):
    progress_status = pyqtSignal(int)

    def row_normalization(self, wb):
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

    def bold_difference_in_lessons_files(self, old_wb, new_wb):
        dif_cell_font = Font(bold=True)
        old_ws = old_wb.active
        new_ws = new_wb.active
        row = 1
        old_row = 1
        while row < new_ws.max_row:
            if 'Класс' in str(new_ws.cell(row, 1).value):
                # print(new_ws.cell(row, 1).value, row)
                flag = False
                for x in range(old_row, old_ws.max_row + 1):
                    if old_ws.cell(x, 1).value == new_ws.cell(row, 1).value:
                        old_row = x
                        flag = True
                        break
                if not flag:
                    print('No matches')
                    raise Exception
                cur_new_row = row + 3
                cur_old_row = old_row + 3
                while not (new_ws.cell(cur_new_row, 1).value is None):
                    for col in range(1, len(new_ws[cur_new_row]) + 1):
                        if new_ws.cell(cur_new_row, col).value != old_ws.cell(cur_old_row, col).value:
                            if new_ws.cell(cur_new_row, col).value is None:
                                new_ws.cell(cur_new_row, col).value = '-окно-'
                            new_ws.cell(cur_new_row, col).font = dif_cell_font
                            new_ws.cell(cur_new_row, col).fill = PatternFill(start_color='ffff00', end_color='ffff00',
                                                                             fill_type='solid')
                            # print(cur_new_row, col)
                    cur_old_row += 1
                    cur_new_row += 1
                row = cur_new_row
            row += 1
        return new_wb

    def bold_difference_in_school_schedule_teacher_ver(self, old_wb, new_wb):
        dif_cell_font = Font(bold=True)
        old_ws = old_wb.active
        new_ws = new_wb.active
        new_ws_active_row = 9
        old_ws_active_row = 9
        while new_ws_active_row < new_ws.max_row:
            if str(new_ws.cell(new_ws_active_row, 2).value) != str(old_ws.cell(new_ws_active_row, 2).value):
                print(f'Расхождение в строке {new_ws_active_row}')
                raise Exception
            for col in range(1, len(new_ws[new_ws_active_row]) + 1):
                if new_ws.cell(new_ws_active_row, col).value != old_ws.cell(old_ws_active_row, col).value:
                    if new_ws.cell(new_ws_active_row, col).value is None:
                        # print(f'Изменения в {new_ws_active_row, col}: окно')
                        new_ws.cell(new_ws_active_row, col).value = '-окно-'
                    new_ws.cell(new_ws_active_row, col).font = dif_cell_font
                    new_ws.cell(new_ws_active_row, col).fill = PatternFill(start_color='ffff00', end_color='ffff00',
                                                                     fill_type='solid')
                    # print(f'Изменения в {new_ws_active_row, col}: {new_ws.cell(new_ws_active_row, col).value}')
            new_ws_active_row += 1
            old_ws_active_row += 1
        return new_wb

    def bold_difference_in_school_schedule_student_ver(self, old_wb, new_wb):
        old_ws = old_wb.active
        new_ws = new_wb.active
        new_ws_active_row = 7
        old_ws_active_row = 7
        while new_ws_active_row < new_ws.max_row:
            if str(new_ws.cell(new_ws_active_row, 2).value) != str(old_ws.cell(new_ws_active_row, 2).value):
                print(f'Расхождение в строке {new_ws_active_row}')
                raise Exception
            for col in range(3, len(new_ws[new_ws_active_row]) + 1):
                if new_ws.cell(new_ws_active_row, col).value != old_ws.cell(old_ws_active_row, col).value:
                    if new_ws.cell(new_ws_active_row, col).value is None:

                        # print(f'Изменения в {new_ws_active_row, col}: окно')
                        new_ws.cell(new_ws_active_row, col).value = '-окно-'
                    new_ws.cell(new_ws_active_row, col).font = Font(bold=True)
                    new_ws.cell(new_ws_active_row, col).fill = PatternFill(start_color='ffff00', end_color='ffff00',
                                                                     fill_type='solid')
                    # print(f'Изменения в {new_ws_active_row, col}: {new_ws.cell(new_ws_active_row, col).value}')
            new_ws_active_row += 1
            old_ws_active_row += 1
        return new_wb

    def day_assemble(self, wb, day):
        res_wb = Workbook()
        res_ws = res_wb.active
        ws_in = wb.active
        res_ws.append([None])
        res_ws.append([None])
        # frame
        res_ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1)
        res_ws.cell(1, 1).value = '№'
        for i in range(1, 12):
            res_ws.append([str(i)])

        row_in = 1
        copy_counter = 0
        while row_in < ws_in.max_row:
            if not('Класс' in str(ws_in.cell(row_in, 1).value)):
                row_in += 1
                continue
            this_class = ws_in.cell(row_in, 1).value.split(' - ')[1]
            this_class_row = row_in
            row_in += 3
            need_to_copy = False
            while not (ws_in.cell(row_in, 1).value is None):
                if ws_in.cell(row_in, day * 2).font.bold or ws_in.cell(row_in, day * 2 + 1).font.bold:
                    need_to_copy = True
                    break
                row_in += 1

            if need_to_copy:
                start_lesson = 1
                if '6' in this_class or '7' in this_class:
                    start_lesson = 5
                res_ws.cell(1, 1 + copy_counter * 2 + 1).value = this_class
                res_ws.merge_cells(start_row=1, start_column=1 + copy_counter * 2 + 1, end_row=1, end_column=1 + copy_counter * 2 + 2)
                res_ws.cell(2, 1 + copy_counter * 2 + 1).value = 'Предмет'
                res_ws.cell(2, 1 + copy_counter * 2 + 2).value = 'Каб.'
                for row in range(12 - start_lesson):
                    for col in range(2):
                        source_cell = ws_in.cell(this_class_row + 3 + row, day * 2 + col)
                        target_cell = res_ws.cell(2 + start_lesson + row, 1 + copy_counter * 2 + col + 1)
                        target_cell.value = source_cell.value
                        target_cell.font = copy(source_cell.font)
                        target_cell.fill = copy(source_cell.fill)
                copy_counter += 1
        return res_wb

    def search_teacher_window_by_lesson_n(self, wb, teacher_name, day_n_0, lesson_n):

        # ws
        pass

    def create_school_schedule_teacher_ver(self, school_wb):
        ## в файле убирается столбец нумерации и специализации. должно остаться 111 столбцов
        ELEMENTARY_SCHOOL_TEACHERS = {'Балахонова Е. М.', 'Горбачева Е. В.', 'Домашенкина О. В.', 'Киселева Н. И.',
                                      'Стражева Г. Н.', 'Чаркина О. В.', 'Ченцова Е. Н.', 'Даймичева Р. Ф.', 'Тихоненкова А. Н.',
                                      'Смагина М. А.', 'Хретинина А. А.', 'Доронкина Л. В.', 'Мазина О. А.', 'Савватеева Г. А.',
                                      'Соколова Я. А.'}
        MAX_COL_INPUT_FILE = 111
        LESSONS_N = 11
        MAX_COL_OUTPUT_FILE = LESSONS_N * len(DAYS) + 3
        wb_out = Workbook()
        ws_out = wb_out.active

        def create_frame():
            for i in range(8):
                ws_out.append([None for i in range(MAX_COL_OUTPUT_FILE)])
            ws_out.merge_cells(start_row=1, start_column=1, end_row=6, end_column=MAX_COL_OUTPUT_FILE)
            ws_out.cell(1, 1).value = 'Расписание уроков на 2024-2025'
            ws_out.merge_cells(start_row=7, start_column=2, end_row=8, end_column=2)
            ws_out.cell(7, 2).value = 'Ф.И.О.'
            for i in range(5):
                ws_out.merge_cells(start_row=7, start_column=3 + i * 11, end_row=7, end_column=3 + (i + 1) * 11 - 1)
                ws_out.cell(7, 3 + i * 11).value = DAYS[i]
                for j in range(11):
                    ws_out.cell(8, 3 + i * 11 + j).value = j + 1

        create_frame()
        ws = school_wb.active
        input_file_row = 8
        pbar = tqdm(total=ws.max_row - input_file_row + 1)
        teacher_counter = 1
        busy_cabinets_list_of_sets = [set() for i in range(55)]
        while input_file_row <= ws.max_row:
            cur_row = [teacher_counter]
            teacher = str(ws.cell(input_file_row, 1).value)
            if teacher in ELEMENTARY_SCHOOL_TEACHERS or teacher == 'None':
            # if teacher == 'None':
                input_file_row += 1
                pbar.update(1)
                continue
            cur_row.append(teacher)
            input_file_col = 2
            while input_file_col < MAX_COL_INPUT_FILE:
                this_class = this_room = ''
                if ws.cell(input_file_row, input_file_col).value:
                    this_class = str(ws.cell(input_file_row, input_file_col).value)
                    if '_' in this_class:
                        this_class = this_class.split('_')[0]
                    this_room = str(ws.cell(input_file_row, input_file_col + 1).value)
                    if 'С' in this_room:
                        this_room = 'СЗ'
                    if "П" in this_room:
                        this_room = 'П'
                    if input_file_col//2 - 1== 0:
                        print(input_file_row, input_file_col, this_room)
                    busy_cabinets_list_of_sets[input_file_col//2 - 1].add(this_room)
                    cur_row.append('\n'.join((this_class, this_room)))
                else:
                    cur_row.append(None)
                input_file_col += 2
            cur_row.append(teacher_counter)
            ws_out.append(cur_row)
            input_file_row += 1
            teacher_counter += 1
            pbar.update(1)
        pbar.close()

        ws_out.append([None for x in range(MAX_COL_OUTPUT_FILE)])
        for output_file_col in range(3, MAX_COL_OUTPUT_FILE):
            this_set = busy_cabinets_list_of_sets[output_file_col - 3]
            free_cabinets = sorted(list(CABINETS_SET - this_set))
            if not free_cabinets:
                free_cabinets.append('НЕТ')
            if len(free_cabinets) <= 10:
                ws_out.cell(9 + teacher_counter - 1, output_file_col).value = '\n'.join(free_cabinets)
        return wb_out

    def create_school_schedule_student_ver(self, normalized_wb):
        START_ROW_OUT = 7
        START_COL_OUT = 3
        N_CLASS = 24
        MAX_COL = N_CLASS * 2 + 4

        def create_frame():
            for i in range(56):
                ws_out.append([None for x in range(MAX_COL)])
            for i in range(1, 6):
                ws_out.merge_cells(start_row=i, start_column=1, end_row=i, end_column=11)
            for i in range(5):
                ws_out.merge_cells(start_row=7 + i * 10, start_column=1, end_row=7 + (i + 1) * 10 - 2, end_column=1)
                ws_out.cell(7 + i * 10, 1).alignment = Alignment(textRotation=90)
                ws_out.cell(7 + i * 10, 1).value = DAYS[i]
                ws_out.merge_cells(start_row=7 + i * 10, start_column=MAX_COL, end_row=7 + (i + 1) * 10 - 2, end_column=MAX_COL)
                ws_out.cell(7 + i * 10, 1).alignment = Alignment(textRotation=90)
                ws_out.cell(7 + i * 10, MAX_COL).value = DAYS[i]
                for j in range(9):
                    ws_out.cell(7 + i * 10 + j, 2).value = j + 1
                    ws_out.cell(7 + i * 10 + j, MAX_COL - 1).value = j + 1
            ws_out.cell(6, 1).alignment = Alignment(textRotation=90)
            ws_out.cell(6, 1).value = 'День'
            ws_out.cell(6, 2).alignment = Alignment(textRotation=90)
            ws_out.cell(6, 2).value = 'Урок'
            ws_out.cell(6, MAX_COL).alignment = Alignment(textRotation=90)
            ws_out.cell(6, MAX_COL).value = 'День'
            ws_out.cell(6, MAX_COL - 1).alignment = Alignment(textRotation=90)
            ws_out.cell(6, MAX_COL - 1).value = 'Урок'

        def minimize_lesson_cabinet(lesson: str, cabinet: str):
            new_lesson = ''
            new_cabinet = ''
            if 'С' in cabinet:
                new_cabinet = 'СЗ'
            elif "П" in cabinet:
                new_cabinet = 'П'
            elif "А" in cabinet:
                new_cabinet = 'АЗ'
            else:
                new_cabinet = cabinet

            if 'английский язык' in lesson:
                new_lesson = lesson.replace('английский язык', 'англ. яз.')
            elif 'обществознание' in lesson:
                new_lesson = lesson.replace('обществознание', 'общ-знание')
            elif 'профориентация' in lesson:
                new_lesson = lesson.replace('профориентация', 'профориент')
            else:
                new_lesson = lesson

            return new_lesson, new_cabinet

        def copying_middle_school():
            nonlocal class_counter, row_in
            ws_out.cell(6, 3 + class_counter * 2).value = this_class
            row_in += 3
            pbar.update(3)
            lesson_counter = 0
            while ws_in.cell(row_in, 1).value:

                for day_counter in range(5):
                    lesson = ws_in.cell(row_in, 2 + day_counter * 2).value
                    cabinet = str(ws_in.cell(row_in, 2 + day_counter * 2 + 1).value)
                    this_row = START_ROW_OUT + day_counter * 10 + lesson_counter
                    this_col = START_COL_OUT + class_counter * 2
                    if lesson:
                        lesson, cabinet = minimize_lesson_cabinet(lesson, cabinet)
                        ws_out.cell(this_row, this_col).value = lesson
                        ws_out.cell(this_row, this_col + 1).value = cabinet
                lesson_counter += 1
                pbar.update(1)
                row_in += 1
                continue
            class_counter += 1

        def merging_high_school():
            nonlocal class_counter, row_in
            # выход через return если во входящем файле след.класс  другой
            while True:
                this_class = ws_in.cell(row_in, 1).value.split(' - ')[1]
                ws_out.cell(6, 3 + class_counter * 2).value = this_class.split('_')[0]
                row_in += 3
                pbar.update(3)
                lesson_counter = 0
                while ws_in.cell(row_in, 1).value:
                    for day_counter in range(5):
                        lesson = ws_in.cell(row_in, 2 + day_counter * 2).value
                        cabinet = str(ws_in.cell(row_in, 2 + day_counter * 2 + 1).value)
                        if lesson:
                            lesson, cabinet = minimize_lesson_cabinet(lesson, cabinet)
                            this_row = START_ROW_OUT + day_counter * 10 + lesson_counter
                            this_col = START_COL_OUT + class_counter * 2
                            # учитываем, что если у групп общий предмет - его не надо дублировать и подписывать группы
                            if not ws_out.cell(this_row, this_col).value:
                                ws_out.cell(this_row, this_col).value = f'{this_class.split("_")[1]}-{lesson}'
                                ws_out.cell(this_row, this_col + 1).value = cabinet
                            else:
                                if ws_out.cell(this_row, this_col + 1).value != cabinet:
                                    ws_out.cell(this_row, this_col).value += f'\n{this_class.split("_")[1]}-{lesson}'
                                    ws_out.cell(this_row, this_col + 1).value += f'\n{cabinet}'
                                else:
                                    ws_out.cell(this_row, this_col).value = ws_out.cell(this_row, this_col).value.split('-')[-1]

                    lesson_counter += 1
                    pbar.update(1)
                    row_in += 1
                    continue
                # проверка на выход. если класс тот же - class counter не рогаем. это столбцы в выходном файле.
                if not str(ws_in.cell(row_in + 1, 1).value).split(' - ')[-1].split('_')[0] == this_class.split('_')[0]:
                    class_counter += 1
                    return
                else:
                    row_in += 1


        wb_out = Workbook()
        ws_out = wb_out.active
        ws_in = normalized_wb.active
        create_frame()
        class_counter = 0
        pbar = tqdm(total=ws_in.max_row)
        row_in = 1
        while row_in < ws_in.max_row:
            if not ws_in.cell(row_in, 1).value or 'Класс' not in ws_in.cell(row_in, 1).value:
                pbar.update(1)
                row_in += 1
                continue
            this_class = ws_in.cell(row_in, 1).value.split(' - ')[1]
            #  в 10/11 классах надо сливать профили
            if '1' in this_class:
                merging_high_school()
            else:
                copying_middle_school()

        # впихиваемся в лист
        for col in range(3, MAX_COL):
            if col % 2 == 0:
                ws_out.column_dimensions[get_column_letter(col)].width = 38 * 0.138
            else:
                ws_out.column_dimensions[get_column_letter(col)].width = 125 * 0.138
        return wb_out

    def find_teacher_changes_in_student_schedule(self, old_wb, new_wb):
        # По учителям
        # не срабатывает если менялись уроки, а кабинет тот же. в словарь вносится запись на нечетном столбце
        # кабинета, а он не менялся. надо как-то выносить изпод условия несовпадения ячеек
        # не корреткно отрабатывает при замене урков разных учителей местами
        # 'Козлова В. Е.': {'lesson': 2, 'before': ('химия', 308), 'after': ('математика', 302)}
        # 'Грибова А. С.': {'lesson': 4, 'before': ('математика', 302), 'after': ('химия', 308)}
        # как это пофиксить без наличия в базе расписания?
        db = database.Database()
        db.init_db()
        cur_class = ''
        cur_class_id = None
        # {teacher : [{lesson: n, before: (subject, cabinet), after: (subject, cabinet)}, {} ....]}
        teachers_changes_dict = dict()
        old_ws = old_wb.active
        new_ws = new_wb.active
        row = 1
        old_row = 1
        while row < new_ws.max_row:
            if 'Класс' in str(new_ws.cell(row, 1).value):
                cur_class = new_ws.cell(row, 1).value.split(' - ')[1]
                # print(new_ws.cell(row, 1).value, row)
                flag = False
                for x in range(old_row, old_ws.max_row + 1):
                    if old_ws.cell(x, 1).value == new_ws.cell(row, 1).value:
                        old_row = x
                        flag = True
                        break
                if not flag:
                    print('No matches')
                    raise Exception
                cur_new_row = row + 3
                cur_old_row = old_row + 3
                lesson_number = 1
                while not (new_ws.cell(cur_new_row, 1).value is None):
                    for col in range(1, len(new_ws[cur_new_row]) + 1):
                        if new_ws.cell(cur_new_row, col).value != old_ws.cell(cur_old_row, col).value:
                            # Делаем 1 раз для каждого класса. в конце действий по условию найденного слова Класс - сбрасываем
                            if not cur_class_id:
                                cur_class_id = \
                                    db.session.query(db_models.Class.id).filter(
                                        db_models.Class.name == cur_class).one()[0]
                            # Они точно не совпадают. Может окно в новом
                            if new_ws.cell(cur_new_row, col).value is None:
                                # Цикл идет по всем столбцам. В нечетных - кабинет, в четных предмет
                                if col % 2 == 0:
                                    subject_after = 'окно'
                                    subject_before = old_ws.cell(cur_old_row, col).value
                                    subject_before_id = db.session.query(db_models.Subject.id).filter(
                                        db_models.Subject.name == subject_before).one()[0]
                                else:
                                    cabinet_after = 'окно'
                                    cabinet_before = old_ws.cell(cur_old_row, col).value
                            # Может в старом
                            elif old_ws.cell(cur_old_row, col).value is None:
                                # Цикл идет по всем столбцам. В нечетных - кабинет, в четных предмет
                                if col % 2 == 0:
                                    subject_before = 'окно'
                                    subject_after = new_ws.cell(cur_new_row, col).value
                                    subject_after_id = db.session.query(db_models.Subject.id).filter(
                                        db_models.Subject.name == subject_after).one()[0]
                                else:
                                    cabinet_before = 'окно'
                                    cabinet_after = new_ws.cell(cur_new_row, col).value
                            # Может и там и там урок
                            else:
                                # Цикл идет по всем столбцам. В нечетных - кабинет, в четных предмет
                                if col % 2 == 0:
                                    subject_before = subject_before = old_ws.cell(cur_old_row, col).value
                                    subject_before_id = db.session.query(db_models.Subject.id).filter(
                                        db_models.Subject.name == subject_before).one()[0]
                                    subject_after = new_ws.cell(cur_new_row, col).value
                                    subject_after_id = db.session.query(db_models.Subject.id).filter(
                                        db_models.Subject.name == subject_after).one()[0]
                                else:
                                    cabinet_before = old_ws.cell(cur_old_row, col).value
                                    print('---->', new_ws.cell(cur_new_row, col).value, '<-----')
                                    cabinet_after = new_ws.cell(cur_new_row, col).value
                            # не ясно по айди какого предмета (новго или старого) лезть в базу. выбираем не окно
                            if subject_after != 'окно':
                                key_subject_id = subject_after_id
                            else:
                                key_subject_id = subject_before_id
                            print(cur_class, new_ws.cell(cur_new_row, col).value,
                                  old_ws.cell(cur_old_row, col).value)
                            print(cur_class_id, key_subject_id)
                            teacher_obj = db.session.query(db_models.Teacher).join(
                                db_models.ClassTeacherSubjectConnection,
                                db_models.ClassTeacherSubjectConnection.teacher_id == db_models.Teacher.id).filter(
                                db_models.ClassTeacherSubjectConnection.class_id == cur_class_id,
                                db_models.ClassTeacherSubjectConnection.subject_id == key_subject_id).one()
                            print(col)
                            if col % 2 != 0:
                                teacher_name_in_dict = f'{teacher_obj.surname} {" ".join([x[0] + "." for x in teacher_obj.name_last_name.split()])}'
                                print(teacher_name_in_dict)
                                if teacher_name_in_dict in teachers_changes_dict:
                                    teachers_changes_dict[teacher_name_in_dict].append({'lesson': lesson_number,
                                                                                        'before': (subject_before,
                                                                                                   cabinet_before),
                                                                                        'after': (subject_after,
                                                                                                  cabinet_after)})
                                else:
                                    print('кря кря')
                                    teachers_changes_dict[teacher_name_in_dict] = [{'lesson': lesson_number,
                                                                                    'before': (
                                                                                    subject_before, cabinet_before),
                                                                                    'after': (
                                                                                    subject_after, cabinet_after)}]
                            print(cur_new_row, col)
                    cur_old_row += 1
                    cur_new_row += 1
                    lesson_number += 1
                row = cur_new_row
            row += 1
            cur_class_id = ''
        print(teachers_changes_dict)
        return teachers_changes_dict


def normalization_scenario(file):
    wb_in = load_workbook(file)
    toolbox = FuncToolBox()
    wb_out = toolbox.row_normalization(wb_in)
    wb_out.save(f'{file.split(".")[0]}_NORM.xlsx')
    return wb_out


def checking_class_differences_scenario(changes_file, base_file='расписание учеников_NORM.xlsx', day=-1):
    base_wb = load_workbook(path.join(CUR_SCHEDULES_FOLDER_PATH,base_file))
    changes_wb = load_workbook(changes_file)
    toolbox = FuncToolBox()
    if 'NORM' not in base_file:
        base_wb = toolbox.row_normalization(base_wb)
        base_wb.save(f'{base_file.split(".")[0]}_NORM.xlsx')
    if 'NORM' not in changes_file:
        changes_wb = toolbox.row_normalization(changes_wb)
        changes_wb.save(f'{changes_file.split(".")[0]}_NORM.xlsx')

    wb_differs = toolbox.bold_difference_in_lessons_files(base_wb, changes_wb)
    if day == -1:
        filename = f'{changes_file.split(".")[0]}_DIFFERS.xlsx'
        wb_differs.save(path.join(RESULT_FOLDER_PATH, filename))
        print('done!')
    else:
        wb_differs_day = toolbox.day_assemble(wb_differs, day)
        filename = f'{changes_file.split(".")[0]}_DIFFERS_DAY_{day}.xlsx'
        wb_differs_day.save(path.join(RESULT_FOLDER_PATH, filename))
        print('done!')


def checking_school_schedule_teacher_ver_differences_scenario(file1, file2):
    wb_in1 = load_workbook(file1)
    wb_in2 = load_workbook(file2)
    toolbox = FuncToolBox()
    wb_out = toolbox.bold_difference_in_school_schedule_teacher_ver(wb_in1, wb_in2)
    wb_out.save(f'{file2.split(".")[0]}_DIFFERS.xlsx')
    print('done!')

def checking_school_schedule_pupils_ver_differences_scenario(base_filename, new_filename):
    old_pupils_schedule = load_workbook(base_filename)
    new_pupils_schedule = load_workbook(new_filename)
    tool = FuncToolBox()
    diff_schedule = tool.bold_difference_in_school_schedule_student_ver(old_pupils_schedule, new_pupils_schedule)
    res_filename = f'{new_filename.split(".")[0]}_DIFF.xlsx'
    diff_schedule.save(path.join(RESULT_FOLDER_PATH, res_filename))
    print('done!')

def printing_school_schedule_teacher_ver_scenario(file):
    toolbox = FuncToolBox()
    wb = load_workbook(file)
    res = toolbox.create_school_schedule_teacher_ver(wb)
    res.save(f'{file.split(".")[0]}_PRINT.xlsx')
    print('done!')

def printing_pupils_schedule_scenario(file, normalized=False, save_normalized=True):
    toolbox = FuncToolBox()
    wb = load_workbook(file)
    if not (normalized or 'NORM' in file):
        wb = toolbox.row_normalization(wb)
    if save_normalized and not 'NORM' in file:
        wb.save(f'{file.split(".")[0]}_NORM.xlsx')
    res = toolbox.create_school_schedule_student_ver(wb)
    res.save(f'{file.split(".")[0]}_PRINT.xlsx')
    print('done!')

if __name__ == '__main__':
    checking_class_differences_scenario('fr 20 12.xlsx', day=5)
