from statistics.statistics import get_statistics

from parsing.parse import download_data, upload_to_disk
from parsing.prepare_data import prepare_data

if __name__ == '__main__':
    download_data()
    data = prepare_data()
    get_statistics(data)
    upload_to_disk()
