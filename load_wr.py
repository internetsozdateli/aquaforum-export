import xlrd
import re
import json
import os
from datetime import datetime

def ready():
    if os.path.isfile('./prices/wr/ready_flag'):
        os.remove('./prices/wr/ready_flag')
        return True
    else:
        return False

def load_dict(filepath):
    # загрузка словаря с правилами оптовых цен
    with open(filepath, 'r', encoding='utf-8') as f:
       return json.load(f)

def extract_remains(filepath):
    result = {}

    # Открываем файл
    workbook = xlrd.open_workbook(filepath)
    sheet = workbook.sheet_by_index(0)  # первая вкладка

    for row_idx in range(sheet.nrows):
        row = sheet.row_values(row_idx)

        if len(row) < 4:
            continue  # пропускаем строки с недостаточным количеством ячеек
        try:
            product_id = int(row[0])
        except ValueError:
            continue  # если ID некорректный - пропускаем

        try:
            amount = int(row[3])
        except ValueError:
            amount = row[3].strip()

        result[f"WR{product_id}"] = amount

    return result

remains = extract_remains('./prices/wr/remains.xls')
price_rules = {
    'RIVER'      : load_dict('./prices/wr/river.json'),
    'WELTWASSER' : load_dict('./prices/wr/weltwasser.json'),
    'МОНОМАХ'    : load_dict('./prices/wr/monomax.json'),
    'WEMOR'      : load_dict('./prices/wr/wemor.json'),
}

def contains_cyrillic(text):
    return bool(re.match(r'^[А-Яа-яЁё]', text))

def process_xls_file(filepath):
    category = ""
    c_list = {}
    brand = ""
    result = []

    # Открываем файл
    workbook = xlrd.open_workbook(filepath)

    # Проходим по всем листам
    for sheet in workbook.sheets():

        # Пропускаем первые две строки на вкладке
        for row_idx in range(2, sheet.nrows):
            row = sheet.row_values(row_idx)

            # Проверка: пустая строка (все ячейки пустые)
            if all(cell == '' for cell in row):
                break  # переходим к следующей вкладке

            first_cell = str(row[0]).strip()
            second_cell = str(row[1]).strip() if len(row) > 1 else ''

            if first_cell:  # если первая ячейка НЕ пуста - товар
                if (len(row) < 6):
                    continue

                try:
                    product_id = int(float(first_cell))
                    price = float(row[4])
                except ValueError:
                    continue  # если ID некорректный или не задана цена - пропускаем

                p_id = f"WR{product_id}"

                if brand in price_rules :
                     percent = price_rules[brand][category] if category in price_rules[brand] else price_rules[brand]['default']
                     mult = (100 + percent) / 100
                else:
                     mult = 1

                price_opt = round(price * mult)

                result.append({
                    "id"       : p_id,
                    "model"    : str(row[2]).strip(),
                    "category" : category,
                    "brand"    : brand,
                    "name"     : second_cell,
                    "pr_opt"   : round(price_opt),
                    "pr_ret"   : round(price),
                    "quantity" : remains[p_id] if p_id in remains else '-',
                    "promo"    : str(row[5]).strip() if len(row) > 5 else ''
                })
                continue

            if contains_cyrillic(second_cell):
                category = second_cell
                c_list[second_cell] = 0
            elif re.match(r'^\d+\.\s+\S+', second_cell):
                # Извлекаем бренд после числа
                match = re.match(r'^\d+\.\d*\s+(\S+)', second_cell)
                if match:
                    brand = match.group(1)

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    l = len(result)
    print(f"{now} Обновлен прайс wr, получено {l} строк")
    return result

def load_wr():
    return process_xls_file('./prices/wr/prices.xls')

