from statistics.statistics import get_statistics

from parsing.parse import download_data, upload_to_disk
from parsing.prepare_data import prepare_data, create_dict_categories

if __name__ == '__main__':
    download_data()
    data = prepare_data()
    expenditure_category, income_category = create_dict_categories()
    get_statistics(data, expenditure_category, income_category)
    upload_to_disk()
