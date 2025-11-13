import xlrd
import re
import json
import os
from datetime import datetime

def ready():
    if os.path.isfile('./prices/aquanet-v/ready_flag') and os.path.isfile('./prices/aquanet-m/ready_flag'):
        os.remove('./prices/aquanet-v/ready_flag')
        os.remove('./prices/aquanet-m/ready_flag')
        return True
    else:
        return False

def extract_remains(filepath):
    result = {}

    workbook = xlrd.open_workbook(filepath)

    for sheet in workbook.sheets():
        for row_idx in range(1, sheet.nrows):
            row = sheet.row_values(row_idx)

            try:
                product_id = int(row[5])
            except ValueError:
                continue  # если не задан остаток — пропускаем

            try:
                amount = int(row[1])
            except ValueError:
                amount = row[1].strip()

            result[product_id] = amount

    return result

def load_dict(filepath):
    # загрузка словаря с правилами оптовых цен
    with open(filepath, 'r', encoding='utf-8') as f:
       return json.load(f)

def extract_prices(filepath):
    result = {}

    workbook = xlrd.open_workbook(filepath)

    for sheet in workbook.sheets():
        for row_idx in range(1, sheet.nrows):
            row = sheet.row_values(row_idx)

            try:
                product_id = int(row[1])
            except ValueError:
                continue  # если не задан код товара - пропускаем

            try:
                amount = float(row[4])
            except ValueError:
                amount = row[4].strip()

            result[product_id] = amount

    return result

def process_xls_file_v(filepath):
    result = []

    remains = extract_remains('./prices/aquanet-v/remains.xls')
    price_rules = load_dict('./prices/aquanet-v/rules.json')

    # Открываем файл
    workbook = xlrd.open_workbook(filepath)

    # Проходим по всем листам
    for sheet in workbook.sheets():

        if not sheet.name.strip().upper() in price_rules:
            continue

        rules = price_rules[sheet.name.strip().upper()]

        # Пропускаем первые две строки на вкладке
        for row_idx in range(2, sheet.nrows):
            row = sheet.row_values(row_idx)

            try:
                product_id = int(row[rules['id']])
            except ValueError:
                continue  # если ID некорректный - пропускаем, это не товар

            try:
                price = float(row[rules['price_ret']])
            except ValueError:
                price = 0  # если не задана цена - пропускаем

            mult = (100 + rules['percent']) / 100

            price_opt = round(price * mult)

            result.append({
                "id"       : str(product_id),
                "model"    : str(product_id).rjust(8, '0'),
                "category" : rules['category'],
                "brand"    : 'Aquanet',
                "name"     : str(row[rules['name']]).strip(),
                "pr_opt"   : round(price_opt),
                "pr_ret"   : round(price),
                "quantity" : remains[product_id] if product_id in remains else '-',
            })

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    l = len(result)
    print(f"{now} Обработан прайс aquanet (ванны), получено {l} строк")
    return result


def process_xls_file_m(filepath):
    result = []

    prices = extract_prices('./prices/aquanet-m/prices.xls')
    price_rules = load_dict('./prices/aquanet-m/rules.json')

    # Открываем файл
    workbook = xlrd.open_workbook(filepath)
    c = {}
    # Проходим по всем листам
    for sheet in workbook.sheets():


        # Пропускаем первые две строки на вкладке
        for row_idx in range(1, sheet.nrows):
            row = sheet.row_values(row_idx)

            try:
                product_id = int(row[5])
            except ValueError:
                continue  # если ID некорректный - пропускаем, это не товар

            try:
                amount = int(row[1])
            except ValueError:
                amount = row[1].strip()

            category = row[3].strip()
            price_ret = prices[product_id] if product_id in prices else 0
            mult = (100 + (price_rules[category] if category in price_rules else price_rules['default'])) / 100
            price_opt = round(price_ret * mult)

            result.append({
                "id"       : str(product_id),
                "model"    : row[4].strip(),
                "category" : category,
                "brand"    : row[9].strip(),
                "name"     : row[0].strip(),
                "pr_opt"   : round(price_opt),
                "pr_ret"   : round(price_ret),
                "quantity" : amount
            })

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    l = len(result)
    print(f"{now} Обработан прайс aquanet (мебель), получено {l} строк")
    return result

def load_aquanet_v():
    return process_xls_file_v('./prices/aquanet-v/prices.xls')

def load_aquanet_m():
    return process_xls_file_m('./prices/aquanet-m/remains.xls')
