import math
from typing import Dict, List, Tuple

import pandas as pd

from parsing.models import MonthCategories
from settings import MONTHS, PATH_TO_DATA, CATEGORY_NAME_SHEET, COLORS


def prepare_data() -> List[MonthCategories]:
    """Подготовка данных, парсинг xlsx файла"""
    book = pd.ExcelFile(PATH_TO_DATA)
    finance = []

    for sheet in book.sheet_names:
        if not __is_sheet_finance(sheet):
            continue
        month, year = __get_month_and_year_by_sheet(sheet)
        data = pd.read_excel(PATH_TO_DATA, sheet_name=sheet, header=1)

        month_category = MonthCategories(year=year, month=month,
                                         income_data=__create_month_data(data, "Название", "Цена"),
                                         expenditure_data=__create_month_data(data, "Название.1", "Цена.1"))
        finance.append(month_category)
    return finance


def __is_sheet_finance(sheet: str) -> bool:
    """Функция проверяет является ли этот лист финансовым. Название вида: 'Год. Месяц' """
    name = sheet.split()
    if len(name) == 2:
        year, month = name
        if year[-1] == "." and year[:-1].isdigit() and month in MONTHS:
            return True
    return False


def __get_month_and_year_by_sheet(sheet: str) -> Tuple[int, int]:
    """Парсинг названия листа, вида: 'Год. Месяц' """
    year, month = sheet.split()
    year = year[:-1]
    return MONTHS[month], int(year)


def __create_month_data(data, name_of_category: str, name_of_cost: str) -> Dict[str, int]:
    """Составление словарей из категорий и трат за месяц за эту категорию"""
    result = {}

    for category, summa in zip(data[name_of_category], data[name_of_cost]):
        if isinstance(category, float) and math.isnan(category):
            break
        if category not in result:
            result[category] = 0
        result[category] += summa
    return result


def create_dict_categories() -> Tuple[Dict[str, int], Dict[str, int]]:
    """Нумерация категорий расходов и доходов, используется для выбора цвета графика"""
    def create_dict(add=""):
        result = {}
        len_colors = len(COLORS)

        for index, category in enumerate(data["Название" + add]):
            if isinstance(category, float) and math.isnan(category):
                break
            result[category] = COLORS[index % len_colors]

        return result

    data = pd.read_excel(PATH_TO_DATA, sheet_name=CATEGORY_NAME_SHEET, header=0)
    expenditure_category = create_dict()
    income_category = create_dict(".1")

    return expenditure_category, income_category
