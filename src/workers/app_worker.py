from ..service import config, log, ERRORS, Helper
from pywinauto.application import Application, AppStartError, ProcessNotFoundError
from pywinauto.findbestmatch import MatchError
from time import sleep
import re


class AppWorker:
    """
    Class for the working with windows applications
    """

    def __init__(self):
        self.program_name = self.__class__.__name__.replace('Worker', '')
        self.app_conf = config[self.program_name]
        self.program_obj = self.launch(self.app_conf['program_path'], self.app_conf['launch_type'])
        self.executed = False

    @staticmethod
    def connect(program: str):
        """Method for connecting to an opened program"""

        try:
            return Application(backend='uia').connect(path=program)
        except FileNotFoundError:
            log.exception(f'Fail connection to  {program}')

    def launch(self, program_path: str, launch_type: str = 'normal'):
        """
        Function for launch a program using one of two method
        :param program_path: the full program path
        :param launch_type: if value == 'normal' an application will be started by calling the executable file
                            if value == 'manual' an application will be started by double-click on program folder
        :return: the instance of Application class type
        """

        launch_types = {
            'normal': self.__normal_launch,
            'manual': self.__manual_launch
        }

        try:
            return launch_types[launch_type](program_path)
        except (AppStartError, ProcessNotFoundError, MatchError):
            log.exception(ERRORS.get('file_not_found').format(file=self.program_name))

    @staticmethod
    def __normal_launch(program_path) -> Application:
        """Description in launch() doc"""

        app = Application(backend='uia').start(program_path)
        return app

    def __manual_launch(self, program_path: str) -> Application:
        """Description in launch() doc"""

        file_parse_dict = Helper.parse_file_path(program_path)
        self.launch(f"explorer.exe {file_parse_dict['folder_path']}", "normal")
        explorer = self.connect('explorer.exe')
        dlg = explorer[file_parse_dict.get('folder_name')]
        dlg[file_parse_dict['file_name']].click_input(double=True)
        sleep(5)
        explorer.kill()
        return self.connect(file_parse_dict['file_name'])

    @classmethod
    def _wait_process(cls, app, exit_comp, progress_field) -> bool:
        """
        :param app:
        :param exit_comp:
        :param progress_field:
        :return: True - if finished after the progress bar is full,
                 False - if finished after closing the pop-up window completion
        """

        dlg = app['Dialog']
        while True:
            sleep(5)
            progress_str = re.sub("[^0-9]", "", dlg[progress_field].window_text())
            progress = int(progress_str) if len(progress_str) > 0 else 0
            if progress >= 100:
                return False
            if dlg[exit_comp].exists():
                dlg.OKButton.click_input()
                return True

    def work(self):
        pass