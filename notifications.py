import telebot
from service import config, Helper, CUST_ERRORS


class Sender:

    def __init__(self, token: str):
        self.token = self.invalidate_token(token)
        self.recipients = Helper.convert_str_to_list(config['NOTIFICATIONS']['tg_recipients'])

    @staticmethod
    def invalidate_token(token):
        """check token format"""

        if isinstance(token, str) and token:
            return token
        else:
            raise ValueError(CUST_ERRORS.get('token_error'))

    @staticmethod
    def form_message(programs: list):
        return CUST_ERRORS.get('automation_failed').format(programs=', '.join(programs))


class TgSender(Sender):

    def __init__(self, token: str):
        super().__init__(token)
        self.bot = self.__validata_bot()

    def __validata_bot(self):
        try:
            return telebot.TeleBot(self.token)
        except Exception as ex:
            print(ex)

    def send_message_to_recipient(self, chat_id: str, message: str):
        """send message to one recipient"""
        self.bot.send_message(chat_id=chat_id, text=message)

    def send_out_notifications(self, recipients, programs: list):
        """send message for all recipients"""

        msg = self.form_message(programs)
        recipients = Helper.convert_str_to_list(recipients)

        if recipients:
            for recipient in recipients:
                self.send_message_to_recipient(recipient, msg)
            print('Error notification sent to telegram')
        else:
            print(CUST_ERRORS.get('recipients_error'))
