import inspect
import json
import logging
import logging.config
import os
import time
from pathlib import Path


class Logger(object):
    def __init__(self, default_path=os.path.join(os.path.dirname(__file__), 'logging_config.json'),
                 default_level=logging.DEBUG):
        self.path = default_path
        self.level = default_level
        self.caller_name = None
        self.detect_caller_info()
        with open(self.path, 'r', encoding='UTF-8') as file:
            logging_config = json.load(file)
        log_folder = Path('/', *os.getenv('LOG_FOLDER').split(',')).resolve().absolute()
        log_folder.mkdir(parents=True, exist_ok=True)
        log_file = Path(log_folder, os.getenv('LOG_FILE'))
        logging_config["handlers"]["info_file"]["filename"] = log_file
        self.logger = self.get_logger(f'{self.caller_name}', logging_config)
        return

    def get_logger(self, name, logging_config):
        logging.config.dictConfig(logging_config)
        logging.Formatter.converter = time.localtime
        logger = logging.getLogger(name)
        logger.setLevel(self.level)
        return logger

    def detect_caller_info(self):
        stack = inspect.stack()
        try:
            caller_frame = stack[2]
            caller_class = caller_frame.frame.f_locals['self'].__class__.__name__
            caller_method = caller_frame.frame.f_code.co_name
            self.caller_name = f'{caller_class}.{caller_method}'
        except (AttributeError, IndexError):
            self.caller_name = None

    def debug(self, msg, *args, **kwargs):
        self.logger.debug(msg, *args, **kwargs)

    def info(self, msg, *args, **kwargs):
        self.logger.info(msg, *args, **kwargs)

    def warning(self, msg, *args, **kwargs):
        self.logger.warning(msg, *args, **kwargs)

    def error(self, msg, *args, **kwargs):
        self.logger.error(msg, *args, **kwargs)

    def critical(self, msg, *args, **kwargs):
        self.logger.critical(msg, *args, **kwargs)
