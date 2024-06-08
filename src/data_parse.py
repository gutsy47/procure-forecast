from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd

from os import listdir


def get_cashflow(path: str, is_header: bool = True) -> list:
    """Возвращает данные из таблиц оборотных ведомостей в виде csv-списка. Ниже приведены заголовки.

    - Quarter: Число, получаемое как 10*Y + Q. Например, 20241 = 2024, 1 квартал
    - Category: Инвентарный/номенклатурный номер категории актива
    - Code: Код в справочнике
    - Name: Наименование нефинансового актива
    - Measure: Единица измерения
    - Start Amount: Количество на первое число (01.01, 01.04, 01.07, 01.10)
    - Start Cost: Стоимость товара на первое число
    - Flow In Amount: Оборот дебет, количество
    - Flow In Cost: Оборот дебет, стоимость
    - Flow Out Amount: Оборот кредит, количество
    - Flow Out Cost: Оборот кредит, стоимость
    - End Amount: Количество на последнее число (31.03, 30.06, 30.09, 30.12)
    - End Cost: Стоимость на последнее число


    :param path: str - Путь к файлу вида "../data/raw/Обороты по счету/*.xlsx"
    :param  is_header: bool - Записывать ли заголовки? При массовом получении для объединения заголовки не нужны
    :return: list - Список вида [[header], [row1], [row2], ...]
    """

    # Сразу отсеиваем таблицы, которые не умеем читать, чтобы не тратить время на загрузку
    if " 105 " not in path and " 21 " not in path and " 101 " not in path:
        raise AttributeError("Таблица неизвестного счёта: " + path)

    # Открываем единственный лист в книге
    workbook: Workbook = load_workbook(filename=path, read_only=True)
    sheet: Worksheet = workbook.active

    # Заголовки и переменная будущей таблицы
    header = [
        "Quarter", "Category", "Code", "Name", "Measure", "Start Amount", "Start Cost", "Flow In Amount",
        "Flow In Cost", "Flow Out Amount", "Flow Out Cost", "End Amount", "End Cost"
    ]
    result = [header] if is_header else []

    # Таблицы имеют две различные структуры в зависимости от счета
    if " 105 " in path:
        quarter_cell = sheet[1][7].value.split()  # Текущий квартал из ячейки (1,8) в виде 10*Y + Q
        quarter = int(quarter_cell[2]) + int(quarter_cell[-2]) * 10

        current_category = None
        for row in sheet.iter_rows(min_row=4, max_col=13, values_only=True):
            if row[0]:
                # Если есть ID, то это актив
                result.append([quarter, current_category] + [cell for cell in row[2:]])
            elif row[1]:
                # Если есть категория, то это категория (elif из-за строки с итогами)
                current_category = row[1].split('-')[0].strip()
    else:
        quarter_cell = sheet[2][0].value.split()  # Текущий квартал из ячейки (2,1) в виде 10*Y + Q
        quarter = int(quarter_cell[6]) + int(quarter_cell[-2]) * 10

        current_category = None
        i = 0  # Индекс изменяется нелинейно, так как данные об одном активе разбиты на четыре строки и есть категории
        m = 2 if " 21 за 1" in path else 0  # Магическое число. Таблица за 1 квартал имеет два пустых столбца
        rows = list(sheet.iter_rows(min_row=15, values_only=True))  # iter_rows намного быстрее обращения sheet[x][y]
        while i < len(rows):
            if rows[i][0] and (rows[i][0].startswith("21.") or rows[i][0].startswith("101.")):
                # Если строка начинается с "21." | "101.", то это категория
                current_category = rows[i][0].strip()
            elif rows[i][0] and i < len(rows) - 2 and rows[i + 2][4]:
                # Если есть имя, но нет номера "продукта для которого закупка", то это начало актива
                result.append([
                    quarter,  # Квартал
                    current_category,  # Категория
                    rows[i + 2][4],  # Код в справочнике
                    rows[i][0].strip(),  # Название
                    "ед",  # Единица измерения
                    (rows[i + 1][10] or 0) - (rows[i + 1][11] or 0),  # Количество на начало
                    (rows[i][10] or 0) - (rows[i][11] or 0),  # Стоимость на начало
                    rows[i + 1][12 + m],  # Количество оборота дебет
                    rows[i][12 + m],  # Стоимость оборота дебет
                    rows[i + 1][13 + m],  # Количество оборота кредит
                    rows[i][13 + m],  # Стоимость оборота кредит
                    (rows[i + 1][14 + m] or 0) - (rows[i + 1][15 + m] or 0),  # Количество на конец
                    (rows[i][14 + m] or 0) - (rows[i][15 + m] or 0)  # Стоимость на конец
                ])
                i += 3
            i += 1

    return result


if __name__ == '__main__':
    # Пример получения данных
    data = []
    is_first = True
    cashflow_paths = ["../data/raw/Обороты по счету/" + x for x in listdir("../data/raw/Обороты по счету")]
    for x in cashflow_paths:
        try:
            data += get_cashflow(x, is_header=is_first)
            is_first = False
        except PermissionError as err:
            print(err)

    df = pd.DataFrame(data[1:], columns=data[0])
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", None)
    print(df)

    df.to_csv("../data/processed/cashflow.csv", index=False)
