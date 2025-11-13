import dbconf
import load_aquanet
import load_rgw
import load_wr
import load_santi
import os
import xlrd
import xlwt
from xlutils.copy import copy

def clear_sheet_content(sheet, max_rows=10000, max_cols=9):
    for row in range(1, max_rows):
        for col in range(max_cols):
            sheet.write(row, col, "")  # перезаписываем ячейку пустым значением

def save_data_to_xls(filepath, sheet_name, data_list):
    # Заголовки — ключи из первого словаря
    headers = list(data_list[0].keys()) if data_list else []

    if os.path.exists(filepath):
        book_rd = xlrd.open_workbook(filepath, formatting_info=True)
        book = copy(book_rd)

        try:
            sheet_index = book_rd.sheet_names().index(sheet_name)
            sheet = book.get_sheet(sheet_index)
            nrows = book_rd.sheet_by_name(sheet_name).nrows
            clear_sheet_content(sheet, nrows)
        except ValueError:
            sheet = book.add_sheet(sheet_name)
    else:
        book = xlwt.Workbook()
        sheet = book.add_sheet(sheet_name)

    # Запись данных
    for row_idx, item in enumerate(data_list, start=1):
        for col_idx, header in enumerate(headers):
            sheet.write(row_idx, col_idx, item.get(header, ''))

    book.save(filepath)

models = dbconf.get_site_models()

#проверяем, есть ли обновления прайс-листа и запускаем обработку при необходимости
if load_wr.ready():
    data = load_wr.load_wr()
    save_data_to_xls('test.xls', 'WW', data)

if load_santi.ready():
    data = load_santi.load_santi()
    save_data_to_xls('test.xls', 'SantiLine', data)

if load_aquanet.ready():
    data = load_aquanet.load_aquanet_v()
    data.extend(load_aquanet.load_aquanet_m())
    save_data_to_xls('test.xls', 'Aquanet', data)

if load_rgw.ready():
    data = load_rgw.load_rgw()
    save_data_to_xls('test.xls', 'RGW', data)


