from telegram import Bot
from services import Helper


class Sender:

    def __init__(self, programs: list, token: str):
        self.programs = programs
        self.token = self.invalidate_token(token)
        self.message = Helper.ERRORS['automation_failed'] + ', '.join(programs)

    @staticmethod
    def invalidate_token(token):
        """check token format"""

        if type(token) == str and len(token) > 0:
            return token
        else:
            raise ValueError(Helper.ERRORS['token_error'])

    def send_message(self, recipient):
        """To be overridden by a child class"""

        pass

    def send_out_notifications(self, recipients):
        """To be overridden by a child class"""

        pass


class TgSender(Sender):

    def send_message_to_recipient(self, chat_id):
        """send message to one recipient"""

        tg = Bot(self.token)
        tg.send_message(text=self.message, chat_id=chat_id)

    def send_out_notifications(self, recipients):
        """send message for all recipients"""

        if len(recipients) > 0:
            for recipient in recipients:
                self.send_message(recipient)
            print('Error report was sent to telegram')
        else:
            print(Helper.ERRORS['recipients_error'])
