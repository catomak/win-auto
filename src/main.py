from os import getenv
from notifications import TgSender
from service import Scheduler, config, log
from automations import WithAppRunner


class Launcher:

    @classmethod
    def route(cls):
        """Function for program routing"""

        __ROUTE = {
            '1': cls.schedule_launch,
            '2': cls.test_launch,
            '3': cls.debug_launch,
        }

        if config['SCHEDULE']['auto_launch'] == '1':
            return __ROUTE['1']()

        common_str = 'Choose mod (1 - schedule, 2 - one-time launch, 3 - debug): '

        while (mod := input(common_str)) not in __ROUTE.keys():
            print('Incorrect command')

        return __ROUTE[mod]()

    @classmethod
    def execute_programs(cls):
        """Function to call the alternate execution of automation programs"""

        log.info(f'Start automation')

        programs = [i for i in config['AUTOMATIONS'] if config['AUTOMATIONS'][i] == "1"]
        err_list = []

        # TODO: redo with a good context manager
        for p in programs:
            with WithAppRunner(p) as application:
                if application:
                    application.work()
                err_list.append(p)
                continue

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

        log.info('Start one-time launch program')
        cls.execute_programs()
        log.info('Test launch ended')

    @staticmethod
    def debug_launch():
        """This function is used for experiments"""

        pass


def main():
    try:
        Launcher.route()
    except Exception as e:
        log.exception(e)


if __name__ == "__main__":
    main()
