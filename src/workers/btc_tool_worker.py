from src.service import config, Helper, log
from time import sleep
from .app_worker import AppWorker


class BtcToolsWorker(AppWorker):

    def work(self) -> bool:
        """Retrieve data on all devices in the network and save it in Excel format"""
        try:
            self.__scan_net()
            self.__export_scan()
            self.__save_scan()
            return True
        except Exception as e:
            log.exception(e)
            return False

    def __scan_net(self):
        sleep(5)
        if self.main_dlg.NoButton:
            self.main_dlg.NoButton.click_input()
        self.main_dlg.ScanButton.click_input()
        self._wait_process(self.program_obj, 'Dialog', 'Progress')
        sleep(1)

    def __export_scan(self):
        self.main_dlg.Header5.click_input()
        self.main_dlg.ExportButton.click_input()
        modal = self.program_obj['Dialog']
        sleep(1)
        today = Helper.get_cur_date('_dd_mm')
        modal[str(config['BtcTools']['data_folder'])].click_input(button='left', double=True)
        modal.ComboBox0.type_keys(f'scan{today}')
        modal.SaveButton.click_input()

    def __save_scan(self,):
        save_window = self.program_obj['Dialog']
        if save_window.YesButton:
            save_window.YesButton.click_input()
        elif save_window.OkButton:
            save_window.OkButton.click_input()
