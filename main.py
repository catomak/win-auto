from pywinauto.application import AppStartError, ProcessNotFoundError, findbestmatch
from os import getenv
from notifications import TgSender
from services import Scheduler, Helper
from automations import AppInterface

config = Helper.get_config('config.json')


class BaseApp:
    @classmethod
    def route(cls):
        """Function for program routing"""

        __ROUTE_MAP = {
            1: 'cls.schedule_launch()',
            2: 'cls.test_launch()',
            3: 'cls.debug_launch()',
            4: 'exit()'
        }

        mode = int(input('Choose mod (1- normal, 2 - test launch, 3 - debug launch, 4 - exit): '))

        if mode in __ROUTE_MAP.keys():
            eval(__ROUTE_MAP[mode])
        else:
            print('Incorrect command')

    @classmethod
    def cycle_job(cls):
        """Program cyclic execution"""

        while True:
            cls.route()

    @classmethod
    def programs_automation(cls):
        """Function to call the alternate execution of automation programs"""

        app_list = [i for i in config['automation_list'].keys() if config['automation_list'][i] == 1]
        print(app_list)

        for p in app_list:
            try:
                AppInterface(p).work()
                app_list.remove(p)
            except AppStartError:
                print(p, Helper.ERRORS['program_failed'])
            except ProcessNotFoundError:
                print(p, Helper.ERRORS['process_not_found'])
            except findbestmatch.MatchError:
                print(p, Helper.ERRORS['process_not_found'])

        if len(app_list) > 0:
            tg = TgSender(app_list, getenv('TG_API'))
            tg.send_out_notifications(config['tg_recipients'])

    @classmethod
    def schedule_launch(cls):
        """Calling the schedule program automation function"""

        sc = Scheduler(config['start_time'], config['sleep_time'])
        sc.schedule_work(cls.programs_automation)

    @classmethod
    def test_launch(cls):
        """One time launch full automation"""
        spc = "-" * 20
        print(f'{spc}One-time launch program. Start at {Helper.get_cur_date("hh:mm")} {spc}')
        cls.programs_automation()
        print(f'{spc} Test launch ended {spc}')

    @staticmethod
    def debug_launch():
        """This function is used for experiments"""
        # path = r'C:\Users\Alex\Desktop\konfigurator'
        # Application(backend='uia').start(f'explorer.exe {path}')
        # explorer = Application(backend='uia').connect(path='explorer.exe')
        # dlg = explorer['konfigurator']
        # dlg['Mercury'].click_input(double=True)
        # user = os.getlogin()
        # print(user)
        # print(os.system("WHERE /R C:\ Mercury.exe"))


def main():
    BaseApp.cycle_job()


if __name__ == "__main__":
    main()
