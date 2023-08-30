import yadisk

from settings import PATH_TO_DATA, TOKEN, YADISK_PATH_TO_DATA

y = yadisk.YaDisk(token=TOKEN)


def download_data():
    """Скачивание файла с финансами с Яндекс Диска"""
    if y.check_token():
        y.download(YADISK_PATH_TO_DATA, PATH_TO_DATA)


def upload_to_disk():
    """Загрузка файла на диск"""
    if y.check_token():
        y.remove(YADISK_PATH_TO_DATA)
        y.upload(PATH_TO_DATA, YADISK_PATH_TO_DATA)
