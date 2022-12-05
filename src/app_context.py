from workers.app_worker import AppWorker
from workers.mercury_worker import MercuryWorker
from workers.btc_tool_worker import BtcToolsWorker
from service import log


class WithAppRunner:

    def __init__(self, program_name):
        self.program_name = program_name
        self.application = self.invalidate_app(self.program_name)

    def __enter__(self) -> AppWorker | None:
        if not self.application or not self.application.program_obj:
            log.exception(f"{self.program_name}: Can't run application module")
            return None

        if not self.application.program_obj:
            log.exception(f"{self.program_name}: Can't get application instance")
            return None

        try:
            log.info(f"{self.application.program_name} is running...")
            return self.application
        except Exception as ex:
            log.exception(f"{self.program_name}: App start exception: {ex}")
            return None

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.application.program_obj:
            self.application.program_obj.kill()
            log.info(f'{self.application.program_name} completed')

    @staticmethod
    def invalidate_app(program_name) -> AppWorker:

        apps_dict = {
            'btctools': BtcToolsWorker,
            'mercury': MercuryWorker
        }

        if apps_dict.get(program_name):
            return apps_dict[program_name]()
        else:
            raise ModuleNotFoundError(f'{program_name} class not found')
