from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

import pandas as pd

from os import listdir

from timeit import default_timer as timer


def get_cashflow(path: str, is_header: bool = True) -> list:
    """Возвращает данные из таблиц оборотных ведомостей в виде csv-списка. Ниже приведены заголовки.

    - Quarter: Число, получаемое как 10*Y + Q. Например, 20241 = 2024, 1 квартал
    - Category: Инвентарный/номенклатурный номер категории актива
    - Code: TODO Код в каком-то справочнике. Каком?
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

    # Проверяем, что входящий путь является путём к книге
    if not path.endswith(".xlsx") or path.split('/')[-1].startswith("~$"):
        raise ValueError("Указан неверный путь: " + path)

    header = [
        "Quarter", "Category", "Code", "Name", "Measure", "Start Amount", "Start Cost", "Flow In Amount",
        "Flow In Cost", "Flow Out Amount", "Flow Out Cost", "End Amount", "End Cost"
    ]
    result = [header] if is_header else []

    # Таблицы имеют три различные структуры в зависимости от номера категории (21, 101, 105)
    if " 105 " in path:
        # Открываем единственный лист в книге
        workbook: Workbook = load_workbook(filename=path)
        sheet: Worksheet = workbook.active

        # Получаем текущий квартал из 8 столбца вида 10*Y + Q
        quarter_cell = sheet[1][7].value.split()
        quarter = int(quarter_cell[2]) + int(quarter_cell[-2]) * 10

        current_category = None  # Отслеживает текущую категорию
        for row in sheet.iter_rows(min_row=4, max_col=13):
            if row[0].value:
                # Если есть ID, то это актив
                result.append([quarter, current_category] + [cell.value for cell in row[2:]])
            elif row[1].value:
                # Если есть категория, то это категория (elif из-за строки с итогами)
                current_category = row[1].value.split('-')[0].strip()

    elif " 21 " or " 101" in path:
        # Открываем единственный лист в книге
        workbook: Workbook = load_workbook(filename=path)
        sheet: Worksheet = workbook.active

        quarter_cell = sheet[2][0].value.split()
        quarter = int(quarter_cell[6]) + int(quarter_cell[-2]) * 10

        current_category = None  # Отслеживает текущую категорию
        i = 15  # Индекс изменяется нелинейно, так как данные об одном активе разбиты на четыре строки и есть категории
        while i < sheet.max_row + 1:
            if sheet[i][0].value and sheet[i][0].value[:3] in ["21.", "101"]:
                # Если строка начинается с "21.", то это категория
                current_category = sheet[i][0].value.strip()
            elif sheet[i][0].value and sheet[i + 2][4].value:
                # Если есть имя, но нет номера "продукта для которого закупка", то это начало актива
                m = 2 if " 21 за 1" in path else 0  # Магическое число. Таблица за 1 квартал имеет два пустых столбца
                row = [
                    quarter,  # Квартал
                    current_category,  # Категория
                    sheet[i + 2][4].value,  # Код в справочнике
                    sheet[i][0].value.strip(),  # Название
                    "шт.",  # Единица измерения
                    (sheet[i+1][10].value or 0) - (sheet[i+1][11].value or 0),  # Количество на начало периода
                    (sheet[i][10].value or 0) - (sheet[i][11].value or 0),  # Стоимость на начало периода
                    sheet[i+1][12 + m].value,  # Количество оборота дебет
                    sheet[i][12 + m].value,  # Стоимость оборота дебет
                    sheet[i+1][13 + m].value,  # Количество оборота кредит
                    sheet[i][13 + m].value,  # Стоимость оборота кредит
                    (sheet[i+1][14 + m].value or 0) - (sheet[i+1][15 + m].value or 0),  # Количество на конец периода
                    (sheet[i][14 + m].value or 0) - (sheet[i][15 + m].value or 0)  # Стоимость на конец периода
                ]
                result.append(row)
                i += 3
            i += 1

    else:
        raise AttributeError("Неизвестная таблица: " + path)

    return result


if __name__ == '__main__':
    # Пример получения данных
    data = []
    cashflow_paths = ["../data/raw/Обороты по счету/" + x for x in listdir("../data/raw/Обороты по счету")]
    start = timer()
    for i, x in enumerate(cashflow_paths):
        try:
            start1 = timer()
            data += get_cashflow(x, is_header=True if i == 0 else False)  # Берём заголовки только в первый раз
            print(f"Прочитано {i + 1}/{len(cashflow_paths)} за {round(timer() - start1)} секунд: {x}")
        except ValueError as error:
            print(error)
    print("Данные прочитаны. Время чтения: ", round(timer() - start), "секунд")

    df = pd.DataFrame(data[1:], columns=data[0])
    pd.set_option("display.max_columns", None)
    pd.set_option("display.max_rows", None)
    pd.set_option("display.width", None)
    print(df)

    print("Идёт выгрузка в ../data/processed/cashflow.csv...")
    df.to_csv("../data/processed/cashflow.csv", index=False)
    print("Выгрузка завершена.")
