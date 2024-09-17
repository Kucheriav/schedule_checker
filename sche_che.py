from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
import argparse
from tqdm import tqdm


def row_normalization(wb):
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
            row += 1
        else:
            ws_out.append([cell.value for cell in ws[row]])
            row += 1
            ws_out.append([cell.value for cell in ws[row]])
            for col in range(1, len(ws[row]) + 1):
                if ws_out.cell(ws_out.max_row, col).value is None:
                    ws_out.merge_cells(start_row=ws_out.max_row - 1, start_column=col, end_row=ws_out.max_row,
                                       end_column=col)
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


def bold_difference(old_wb, new_wb):
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


def bold_difference_v2(old_wb, new_wb):
    dif_cell_font = Font(bold=True)
    old_ws = old_wb.active
    new_ws = new_wb.active
    row = 1
    while row < new_ws.max_row:
        if 'Класс' in new_ws.cell(row, 1).value:
            old_row = ''
            for x in range(1,old_ws.max_row + 1):
                if old_ws.cell(x, 1).value == new_ws.cell(row, 1).value:
                    old_row = x
                    break
            if not old_row:
                print('No matches')
                raise Exception
            cur_new_row = row
            cur_old_row = old_row
            offset = 0

        row += 1


#
if __name__ == '__main__':
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
