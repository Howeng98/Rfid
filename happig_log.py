import os
import logging
import datetime
import schedule
import threading


class Log(object):
    def __init__(self, save_path):
        self.save_path = save_path

        '''
        log setting
        '''
        self.logger = logging.getLogger()  # create one log obj for each log kind
        self.logger.setLevel(logging.DEBUG)

        datetime_now = datetime.datetime.now()

        if not os.path.exists(os.path.join(self.save_path, str(datetime_now.date()))):
            os.makedirs(os.path.join(self.save_path, str(datetime_now.date())))
        log_filename = os.path.join(save_path, str(datetime_now.date()),
                'tracelog_' + str(datetime_now.hour) + '.txt')
        log_handler = logging.FileHandler(log_filename)
        log_handler.setLevel(logging.DEBUG)
        log_formatter = logging.Formatter(
                '%(asctime)s - %(levelname)s : %(message)s')
        log_handler.setFormatter(log_formatter)
        self.logger.addHandler(log_handler)
        # maintain log :以小時為單位存放

        '''
        schedule setting and calling＊
        '''
        # schedule.every().day.at("23:00").do(self.directory_create)  # 每天建一個log資料夾存放log
        schedule.every(1).hours.do(self.log_initial)

    def directory_maybe_create(self, datetime_now):
        try:
            if not os.path.exists(os.path.join(self.save_path, str(datetime_now.date()))):
                os.makedirs(os.path.join(self.save_path, str(datetime_now.date())))
        except Exception as e:
            print(e)
            logging.error(e)

    def log_initial(self):
        datetime_now = datetime.datetime.now()
        self.directory_maybe_create(datetime_now)

        try:
            self.logger.handlers[0].stream.close()
            self.logger.removeHandler(self.logger.handlers[0])
            log_filename = os.path.join(self.save_path, str(datetime_now.date()),
                                        'tracelog_' + str(datetime_now.hour) + '.txt')
            log_handler = logging.FileHandler(log_filename)
            log_handler.setLevel(logging.DEBUG)
            log_formatter = logging.Formatter(
                '%(asctime)s - %(levelname)s : %(message)s')
            log_handler.setFormatter(log_formatter)
            self.logger.addHandler(log_handler)
        except Exception as e:
            print(e)
            logging.error(e)
