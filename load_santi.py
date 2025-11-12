import xlrd
import re
import json
import os
from datetime import datetime

def ready():
    if os.path.isfile('./prices/santi/ready_flag'):
        os.remove('./prices/santi/ready_flag')
        return True
    else:
        return False

def extract_remains(filepath):
    result = {}

    workbook = xlrd.open_workbook(filepath)
    sheet = workbook.sheet_by_index(0)  # первая вкладка

    for row_idx in range(sheet.nrows):
        row = sheet.row_values(row_idx)

        if len(row) < 3:
            continue  # пропускаем строки с недостаточным количеством ячеек
        try:
            product_id = row[0].replace(" ", "")
        except ValueError:
            continue  # если не задан остаток — пропускаем

        try:
            amount = int(float(row[2]))
        except ValueError:
            amount = row[2].strip()

        if product_id == "":
            continue

        result[product_id] = amount

    return result

remains = extract_remains('./prices/santi/remains.xls')

def process_xls_file(filepath):
    result = []

    # Открываем файл
    workbook = xlrd.open_workbook(filepath)

    # Проходим по всем листам
    for sheet in workbook.sheets():
        category = sheet.name

        for row_idx in range(4, sheet.nrows):
            row = sheet.row_values(row_idx)

            # Проверка: пустая строка (все ячейки пустые)
            if all(cell == '' for cell in row):
                break  # переходим к следующей вкладке

            if (len(row) < 8):
                continue

            try:
                product_id = row[0].replace(" ", "")
                price = float(row[7])
            except ValueError:
                continue  # если ID некорректный или не задана цена — пропускаем

            mult = (100 - 13) / 100 # оптовая скидка от РРЦ

            price_opt = round(price * mult)

            result.append({
                "id": product_id,
                "model": product_id,
                "category": category,
                "brand": 'SantiLine',
                "name": row[2].strip(),
                "pr_opt": round(price_opt),
                "pr_ret": round(price),
                "quantity": remains[product_id] if product_id in remains else 0,
                "promo": ''
            })
            continue

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    l = len(result)
    print(f"{now} Обновлен прайс santiline, получено {l} строк")
    return result


def load_santi():
    return process_xls_file('./prices/santi/prices.xls')
