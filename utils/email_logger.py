import datetime


class EmailSendingLogger:
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance.log_data = []
        return cls._instance

    def log_email_sent(self, recipient, subject, success, error_code=None, error_message=None):
        log_entry = {
            'time': datetime.datetime.utcnow() + datetime.timedelta(hours=8),
            'recipient': recipient,
            'subject': subject,
            'success': success,
            'error_message': f'error code {error_code}, {error_message}' if not success else ''
        }
        self.log_data.append(log_entry)

    def get_log_data(self):
        return self.log_data
