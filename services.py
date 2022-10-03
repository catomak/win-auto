from datetime import datetime
from json import load
from time import sleep
import schedule
import re


class Helper:
    __instance = None

    ERRORS = {
        'config_not_found': 'Configuration not found. Program directory must contain config.json',
        'program_not_found': 'not found. Check program location and update configuration file',
        'program_failed': ' not found. Check program and its full path in configuration file',
        'process_not_found': ' not found. Check program and its full path and that this program in the Toolbars',
        'start_time': 'Start time must be in str format (ex.: 09:00)',
        'sleep_time': 'Sleep time must be in int seconds format between 0 and 86400 (ex.: 60)',
        'mercury_data_incorrect': 'Mercury data not found or is incorrect. The data was not recorded',
        'date_format': 'Unrecognized date format. The correct formats: ',
        'automation_failed': 'Не удалось автоматически выполнить программы: ',
        'recipients_error': 'Telegram recipients for bug report not found',
        'token_error': 'token to message sending not found'
    }

    now = datetime.now()
    td = now.today()

    def __new__(cls, *args, **kwargs):
        if cls.__instance is None:
            cls._instance = super.__new__(cls)
        return cls.__instance

    def __del__(self):
        self.__instance = None

    @staticmethod
    def get_cur_date(date_format):
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
            raise SyntaxError(Helper.ERRORS['date_format'], {str(datetime_dict.keys())})

    @classmethod
    def get_config(cls, path):
        configuration = {}
        try:
            with open(path) as f:
                data = load(f)
                for d in data:
                    configuration[d] = data.get(d)
            return configuration
        except FileNotFoundError:
            raise (Helper.ERRORS['config_not_found'])


class Scheduler:

    def __init__(self, work_time: str, sleep_time: int):
        self.work_time = self.invalidate_time(work_time)
        self.sleep_time = self.invalidate_time(sleep_time)

    def __new__(cls, *args, **kwargs):
        print(f'Program runs every day at {config["start_time"]}. '
              f'To change start time, edit parameter in the [program directory] ->  "config.json".')
        return super().__new__(cls)

    @staticmethod
    def invalidate_time(time):
        if type(time) == str:
            if re.match('^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$', time) is None:
                raise ValueError(Helper.ERRORS['start_time'])
        elif type(time) == int:
            if time > 86400 or time < 0:
                raise ValueError(Helper.ERRORS['sleep_time'])
        else:
            raise TypeError(Helper.ERRORS['start_time'] + '\n' + Helper.ERRORS['sleep_time'])
        return time

    def schedule_work(self, function):
        schedule.every().day.at(self.work_time).do(function)
        while True:
            schedule.run_pending()
            sleep(self.sleep_time)


config = Helper.get_config('config.json')