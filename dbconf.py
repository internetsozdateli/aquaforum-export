import re
import pymysql

def parse_opencart_config(filepath):
    config = {}
    pattern = re.compile(r"define\('(\w+)',\s*'([^']+)'\);")

    with open(filepath, 'r', encoding='utf-8') as f:
        for line in f:
            match = pattern.search(line)
            if match:
                key, value = match.groups()
                config[key] = value

    return {
        'host': config.get('DB_HOSTNAME'),
        'user': config.get('DB_USERNAME'),
        'password': config.get('DB_PASSWORD'),
        'database': config.get('DB_DATABASE'),
        'port': int(config.get('DB_PORT', '3306')),
        'prefix': config.get('DB_PREFIX')
    }

def get_site_models():
    db_config = parse_opencart_config('../config.php')
    models = {}

    conn = pymysql.connect(
        host=db_config['host'],
        user=db_config['user'],
        password=db_config['password'],
        database=db_config['database'],
        port=db_config['port']
    )

    try:
        with conn.cursor() as cursor:
            # SQL-запрос
            sql = "SELECT product_id, model FROM oc_product WHERE price > %s"
            cursor.execute(sql, (0,))  # передача параметра

            # Получение всех результатов
            results = cursor.fetchall()

            # Обработка данных
            for row in results:
                models[row[1]] = row[0]

    finally:
        conn.close()
    return models