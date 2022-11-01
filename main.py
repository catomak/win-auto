from datetime import datetime
from os import getenv
from notifications import TgSender
from service import Scheduler, Helper, config
from automations import WithAppRunner


class Launcher:

    @classmethod
    def route(cls):
        """Function for program routing"""

        __ROUTE_DICT = {
            '1': cls.schedule_launch,
            '2': cls.test_launch,
            '3': cls.debug_launch,
        }

        if config['SCHEDULE']['auto_launch'] == '1':
            return __ROUTE_DICT['1']()

        common_str = 'Choose mod (1 - schedule, 2 - one-time launch, 3 - debug): '

        while (mod := input(common_str)) not in __ROUTE_DICT.keys():
            print('Incorrect command')

        return __ROUTE_DICT[mod]()

    @classmethod
    def execute_programs(cls):
        """Function to call the alternate execution of automation programs"""

        print(f'Start of automation at {datetime.now().strftime("%H:%M %d.%m")}')

        programs = [i for i in config['AUTOMATIONS'] if config['AUTOMATIONS'][i] == "1"]
        err_list = []

        for p in programs:
            with WithAppRunner(p) as application:
                if not application.work():
                    err_list.append(p)

        if err_list:
            tg = TgSender(getenv('TG_API'))
            tg.send_out_notifications(config['NOTIFICATIONS']['tg_recipients'], programs)

    @classmethod
    def schedule_launch(cls):
        """Calling the schedule program automation function"""

        sc = Scheduler(config['SCHEDULE']['start_time'], int(config['SCHEDULE']['sleep_time']))
        sc.schedule_work(cls.execute_programs)

    @classmethod
    def test_launch(cls):
        """One time launch full automation"""

        spc = "-" * 20
        print(f'{spc}One-time launch program. Start at {Helper.get_cur_date("hh:mm")} {spc}')
        cls.execute_programs()
        print(f'{spc} Test launch ended {spc}')

    @staticmethod
    def debug_launch():
        """This function is used for experiments"""

        pass


def main():
    Launcher.route()


if __name__ == "__main__":
    main()
