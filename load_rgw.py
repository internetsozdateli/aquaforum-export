import xlrd
import json
import os
from datetime import datetime

def ready():
    if os.path.isfile('./prices/rgw/ready_flag'):
        os.remove('./prices/rgw/ready_flag')
        return True
    else:
        return False

def extract_prices(filepath):
    result = {}

    workbook = xlrd.open_workbook(filepath)
    fp_flag = True # первая вкладка с оглавлением, пропускаем ее

    for sheet in workbook.sheets():
        if fp_flag:
            fp_flag = False
            continue
        price_row = -1  # колонка с ценой

        for row_idx in range(10, sheet.nrows):
            row = sheet.row_values(row_idx)
            if price_row < 0:
                if row[0].upper() == 'ФОТО':
                    try:
                        price_row = [col.lower() for col in row].index('ррц')
                    except ValueError:
                        price_row = -1
                    continue
                else: continue
            else:
                product_id = row[2].strip()

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

    prices = extract_prices('./prices/rgw/prices.xls')
    price_rules = load_dict('./prices/rgw/rules.json')

    # Открываем файл
    workbook = xlrd.open_workbook(filepath)
    # Проходим по всем листам
    fp_flag = True  # первая вкладка с оглавлением, пропускаем ее
    category = ""
    name = ""
    c = {}

    for sheet in workbook.sheets():
        if fp_flag:
            fp_flag = False
            continue

        # Пропускаем первые две строки на вкладке
        for row_idx in range(10, sheet.nrows):
            row = sheet.row_values(row_idx)

            c_tmp, product_id = row[0].strip(), row[2].strip()
            if c_tmp != "" and product_id == "":
                category = c_tmp
                c[category] = "1"
                continue
            elif (c_tmp != "" and product_id != "") or (c_tmp == "" and product_id == ""):
                continue

            try:
                amount = int(row[9])
            except ValueError:
                amount = row[9].strip()

            name = row[1].strip() if row[1].strip() else name

            try:
                price_ret = float(prices[product_id]) if product_id in prices else 0

                # очищаем название категории и получаем оптовую скидку
                c_clean = category.replace("  "," ")
                percent = price_rules[c_clean] if c_clean in price_rules else price_rules["default"]
                mult = (100 + percent) / 100

                #рассчитываем цены
                price_opt = round(price_ret * mult)
                price_ret = round(price_ret)
            except (ValueError, TypeError):
                price_ret = prices[product_id]
                price_opt = prices[product_id]

            result.append({
                "id"       : product_id,
                "model"    : row[3].strip(),
                "category" : category,
                "brand"    : 'RGW',
                "name"     : name,
                "pr_opt"   : price_opt,
                "pr_ret"   : price_ret,
                "quantity" : amount
            })

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    l = len(result)
    print(f"{now} Обработан прайс rgw, получено {l} строк")
    return result

def load_rgw():
    return process_xls_file('./prices/rgw/remains.xls')

