import time

from pyodbc import connect
from datetime import datetime

def create_sql_time(str_time: str) -> str:
    print("create_sql_time go")
    formats = [
        '%d-%m-%Y %H:%M:%S',  # Формат с часами, минутами и секундами
        '%d-%m-%Y %H:%M',  # Формат с часами и минутами
        '%d-%m-%Y %H',  # Формат только с часами
        '%d-%m-%Y',  # Формат только с датой
    ]

    if len(str_time) <= 11:
        str_time = str_time.replace(" ", "")

    for fmt in formats:
        try:
            dt = datetime.strptime(str_time, fmt)
            break
        except ValueError:
            continue

    try:
        formatted_time = dt.strftime('%Y-%m-%d %H:%M:%S')
        sql_time = f"#{formatted_time}#"
    except:
        raise SyntaxError
    return sql_time



def sql_main(region: int, forms: list, start_time: str, end_time: str, pwd: str) -> tuple:
    "Создает запрос в БД"
    connection = connect(r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};" + f" DBQ={pwd};")

    start_time_test = time.time()
    # Точка A: начало замера времени

    sql_request = "SELECT Oblast, Rayon, Forma22, SVovlech, Area_ga FROM dkr_table WHERE"
    sql_request += f" Left(SOATO, 1) = '{region}' "

    full_form = "'01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24'"

    forms_sql = ', '.join(f"'{str(form).zfill(2)}'" for form in forms)

    if not forms_sql == full_form:
        if forms_sql == "'01', '02', '03'":
            sql_request += "AND (Ball_PlPoch IS NOT NULL AND Ball_PlPoch <> 0) AND (Forma22 = '01' OR Forma22 = '02' OR Forma22 ='03')"
        elif len(forms_sql) < 72:
            sql_request += f"AND Forma22 IN ({forms_sql}) "
        else:
            forms_sql += ','
            str_form_list = forms_sql.split(' ')
            for i in str_form_list:
                full_form = full_form.replace(i, "")

            sql_request += f"AND Forma22 NOT IN ({full_form}) "

    if start_time or end_time:
        if start_time and end_time:
            start_date_sql = create_sql_time(start_time)
            end_date_sql = create_sql_time(end_time)
            sql_request += f" AND Data_Vvoda BETWEEN {start_date_sql} AND {end_date_sql}"
        elif start_time:
            start_date_sql = create_sql_time(start_time)
            sql_request += f" AND Data_Vvoda > {start_date_sql}"
        else:
            end_date_sql = create_sql_time(end_time)
            sql_request += f" AND Data_Vvoda < {end_date_sql}"

    print(sql_request)

    cursor = connection.cursor()
    cursor.execute(sql_request)
    rows = cursor.fetchall()
    cursor.close()

    end_time_test = time.time()
    execution_time = end_time_test - start_time_test
    print(execution_time)

    return rows


