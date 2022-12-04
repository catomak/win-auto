from datetime import datetime
from time import sleep
from configparser import ConfigParser
import logging
import schedule
import re
import json

ERRORS = {
    "config_not_found": "Can't find configuration file '{config}' in program directory",
    "file_not_found": "{file} not found. Check file location and its full path in configuration file",
    "start_time": "Start time must be in str format (ex.: 09:00)",
    "sleep_time": "Sleep time must be in int seconds format between 0 and 86400 (ex.: 60)",
    "mercury_data_incorrect": "Mercury data not found or is incorrect. The data was not recorded",
    "date_format": "Unrecognized date format. The correct formats ",
    "automation_failed": "Failed to execute programs automatically: {programs}.\nPlease, run the programs manually.",
    "recipients_error": "Telegram recipients for bug report not found",
    "token_error": "token to message sending not found"
}


def get_config(file_name: str) -> dict:
    try:
        configuration = ConfigParser()
        configuration.read(file_name)
        return configuration
    except FileNotFoundError as ex:
        print(ERRORS.get('config_not_found').format(config=file_name))
        raise ex


config = get_config('config.ini')


def get_logger():
    formatting = '%(asctime)s %(levelname)s %(message)s'
    logging.basicConfig(level=logging.INFO, format=formatting, datefmt='%d/%m/%Y %H:%M:%S')
    logger = logging.getLogger('app.log')
    return logger


log = get_logger()


class Helper:

    @staticmethod
    def get_cur_date(date_format: str) -> str:
        now = datetime.now()
        td = now.today()

        datetime_dict = {
            'dd.mm.yyyy': now.strftime("%d.%m.%Y"),
            '_dd_mm': f'_{str(td.day).rjust(2, "0")}_{str(td.month).rjust(2, "0")}',
            'hh:mm': f'{now.strftime("%H:%M:%S")}'
        }

        if date_format in datetime_dict.keys():
            return datetime_dict[date_format]
        else:
            raise SyntaxError(ERRORS.get('date_format'), {str(datetime_dict.keys())})

    @staticmethod
    def convert_str_to_list(string: str) -> list:
        try:
            return json.loads(string)
        except TypeError as ex:
            print('Incorrect sting for converting to list')
            raise ex

    @classmethod
    def get_file_name(cls, file_path: str) -> str:
        return file_path[file_path.rfind('\\')+1:]

    @classmethod
    def get_file_folder_path(cls, path: str, file: str):
        return path[:path.rfind('\\'+file)]

    @classmethod
    def get_file_folder_name(cls, folder_path: str) -> str:
        return folder_path[folder_path.rfind('\\')+1:]

    @classmethod
    def parse_file_path(cls, path: str) -> dict:
        file_name = cls.get_file_name(path)
        folder_path = cls.get_file_folder_path(path, file_name)
        folder_name = cls.get_file_folder_name(folder_path)

        return {
            'file_name': file_name,
            'folder_path': folder_path,
            'folder_name': folder_name
        }


class Scheduler:

    def __init__(self, work_time: str, sleep_time: int):
        self.work_time = self.invalidate_time(work_time)
        self.sleep_time = self.invalidate_time(sleep_time)

    def __new__(cls, *args, **kwargs):
        print(f'Program runs every day at {config["SCHEDULE"]["start_time"]}. '
              f'To change start time, edit parameter in the ../program_directory -> "config.ini".')
        return super().__new__(cls)

    @staticmethod
    def invalidate_time(time):
        if isinstance(time, str) and re.match('^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$', time) is None:
            raise ValueError(ERRORS.get('start_time'))
        elif isinstance(time, int) and (time > 86400 or time < 0):
            raise ValueError(ERRORS.get('sleep_time'))
        return time

    def schedule_work(self, function):
        schedule.every().day.at(self.work_time).do(function)
        while True:
            schedule.run_pending()
            sleep(self.sleep_time)
