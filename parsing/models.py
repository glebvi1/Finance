class MonthCategories:
    """Модель, отображающая одну таблицу"""
    def __init__(self, year: int, month: int, income_data: dict, expenditure_data: dict):
        self.year = year
        self.month = month
        self.income_data = income_data  # {cat1: sum1, ...}
        self.expenditure_data = expenditure_data

    def get_expenditure(self):
        return sum(self.expenditure_data.values())

    def get_income(self):
        return sum(self.income_data.values())

    def __str__(self):
        return f"{self.month}.{self.year}, доход: {self.income_data}, расход: {self.expenditure_data}"
