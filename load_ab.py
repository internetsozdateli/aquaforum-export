import re
import xlrd
import json
import os
from datetime import datetime

def ready():
    if os.path.isfile('./prices/ab/ready_flag'):
        os.remove('./prices/ab/ready_flag')
        return True
    else:
        return False

def extract_prices(filepath):
    result = {}
    pattern_col = r'цвет\w*'
    pattern_meb = r'мебель\w*'

    workbook = xlrd.open_workbook(filepath)

    for sheet in workbook.sheets():
        if bool(re.search(pattern_col, sheet.name, re.IGNORECASE)):
            continue

        # У мебели цена находится в колонке G, у остальных - H
        if bool(re.search(pattern_meb, sheet.name, re.IGNORECASE)):
            price_row = 6
            id_row = 1
        else:
            price_row = 7
            id_row = 2

        for row_idx in range(5, sheet.nrows):
            row = sheet.row_values(row_idx)

            try:
                product_id = int(row[id_row])
            except ValueError:
                continue # код товара не найден, пропускаем

            try:
                price = float(row[price_row])
            except ValueError:
                price = row[price_row].strip()

            result[product_id] = price

    return result


def load_dict(filepath):
    # загрузка словаря с правилами оптовых цен
    with open(filepath, 'r', encoding='utf-8') as f:
       return json.load(f)

def process_xls_file(filepath):
    result = []

    prices = extract_prices('./prices/ab/price_ob.xls') | extract_prices('./prices/ab/price_sm.xls')
    price_rules = load_dict('./prices/ab/rules.json')

    # Открываем файл
    workbook = xlrd.open_workbook(filepath)
    category = ""

    sheet = workbook.sheet_by_index(0)  # первая вкладка

    for row_idx in range(5, sheet.nrows):
        row = sheet.row_values(row_idx)

        c_tmp, product_id = str(row[0]).strip(), str(row[3]).strip()
        if c_tmp != "" and product_id == "":
            category = c_tmp
            continue

        try:
            amount = int(row[7])
        except ValueError:
            amount = row[7].strip()

        try:
            product_id = int(product_id)
        except ValueError:
            continue # продукт не имеет корректного номера

        if not product_id in prices:
            continue

        try:
            price_ret = float(prices[product_id])

            # очищаем название категории и получаем оптовую скидку
            percent = price_rules[category] if category in price_rules else 0
            mult = (100 + percent) / 100

            #рассчитываем цены
            price_opt = round(price_ret * mult)
            price_ret = round(price_ret)
        except (ValueError, TypeError):
            price_ret = prices[product_id]
            price_opt = prices[product_id]

        result.append({
            "id"       : 'AB'+str(product_id),
            "model"    : str(row[4]).strip(),
            "category" : category,
            "brand"    : row[6].strip(),
            "name"     : row[0].strip(),
            "pr_opt"   : price_opt,
            "pr_ret"   : price_ret,
            "quantity" : amount
        })


    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    l = len(result)
    print(f"{now} Обработан прайс allen brau, получено {l} строк")
    return result

def load_allenbrau():
    return process_xls_file('./prices/ab/remains.xls')

