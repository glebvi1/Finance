from string import ascii_uppercase
from typing import Dict, Tuple

from openpyxl import load_workbook
from openpyxl.chart import PieChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import PatternFill

from parsing.models import MonthCategories
from settings import (COLOR_BAD, COLOR_GOOD, INVERT_MONTHS, PATH_TO_DATA,
                      RESULT_NAME_SHEET)


def get_statistics(data):
    workbook = load_workbook(PATH_TO_DATA)
    if RESULT_NAME_SHEET in workbook.sheetnames:
        del workbook[RESULT_NAME_SHEET]
    worksheet = workbook.create_sheet(RESULT_NAME_SHEET)

    worksheet["A1"] = "Баланс:"

    global_expenditure = 0
    global_income = 0
    index_number = 4

    for month_category in data:
        month_category: MonthCategories
        current_expenditure, current_income = __create_month_description(worksheet, month_category, index_number)
        __create_charts(worksheet, index_number, month_category)

        global_expenditure += current_expenditure
        global_income += current_income

        index_number += 10

    worksheet["B1"] = global_income - global_expenditure
    workbook.save(PATH_TO_DATA)


def __create_month_description(worksheet, month_category: MonthCategories, index_number: int) -> Tuple[int, int]:
    """Создание полей: месяц, расход, доход, дельта"""
    month_category: MonthCategories
    worksheet[f"A{index_number}"] = INVERT_MONTHS[month_category.month]
    expenditure, income = month_category.get_expenditure(), month_category.get_income()

    worksheet[f"B{index_number - 1}"] = "Расход"
    worksheet[f"C{index_number - 1}"] = expenditure

    worksheet[f"B{index_number}"] = "Доход"
    worksheet[f"C{index_number}"] = income

    worksheet[f"B{index_number + 1}"] = "Дельта"
    worksheet[f"C{index_number + 1}"] = income - expenditure

    __create_font(worksheet, index_number)

    return expenditure, income


def __create_font(worksheet, index_number: int) -> None:
    """Создание подцветки для описание"""
    worksheet[f"B{index_number - 1}"].fill = worksheet[f"C{index_number - 1}"].fill =\
        PatternFill(start_color=COLOR_BAD, end_color=COLOR_BAD, fill_type="solid")

    worksheet[f"B{index_number}"].fill = worksheet[f"C{index_number}"].fill =\
        PatternFill(start_color=COLOR_GOOD, end_color=COLOR_GOOD, fill_type="solid")

    if worksheet[f"C{index_number + 1}"].value >= 0:
        worksheet[f"C{index_number + 1}"].fill = worksheet[f"B{index_number + 1}"].fill =\
            PatternFill(start_color=COLOR_GOOD, end_color=COLOR_GOOD, fill_type="solid")
    else:
        worksheet[f"C{index_number + 1}"].fill = worksheet[f"B{index_number + 1}"].fill =\
            PatternFill(start_color=COLOR_BAD, end_color=COLOR_BAD, fill_type="solid")


def __create_charts(worksheet, index_number: int, month_category: MonthCategories) -> None:
    """Создание графиков, запись данных для графиков"""
    max_col_income = __write_data_for_chart(worksheet, index_number, month_category.income_data)
    max_col_expenditure = __write_data_for_chart(worksheet, index_number+2, month_category.expenditure_data)

    __draw_pie_chart(worksheet=worksheet, index_number=index_number, max_col=max_col_income,
                     title=f"Доходы {INVERT_MONTHS[month_category.month]}", chart_index=f"E{index_number-1}")
    __draw_pie_chart(worksheet=worksheet, index_number=index_number+2, max_col=max_col_expenditure,
                     title=f"Расходы {INVERT_MONTHS[month_category.month]}", chart_index=f"K{index_number-1}")


def __write_data_for_chart(worksheet, index_number: int, data: Dict[str, int]) -> int:
    """Запись значений для графиков"""
    row_index = 10
    for key, value in data.items():
        worksheet[f"{ascii_uppercase[row_index]}{index_number-1}"] = key
        worksheet[f"{ascii_uppercase[row_index]}{index_number}"] = value
        row_index += 1
    return row_index


def __draw_pie_chart(worksheet, index_number: int, max_col: int, title: str, chart_index: str):
    """Рисование графиков"""
    pie = PieChart()

    labels = Reference(worksheet, min_col=11, max_col=max_col, min_row=index_number-1, max_row=index_number-1)
    data_ref = Reference(worksheet, min_col=11, max_col=max_col, min_row=index_number, max_row=index_number)

    series = Series(data_ref)
    pie.append(series)
    pie.set_categories(labels)

    pie.dataLabels = DataLabelList()
    pie.dataLabels.showVal = True

    pie.title = title
    pie.height = 5
    pie.width = 10

    worksheet.add_chart(pie, chart_index)
