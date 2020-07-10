from functools import partial
import smtplib

import time
import os
import csv
import logging
from selenium_gmail_direct import Google
from datetime import datetime, timedelta
import threading
import yaml
import sched
import time as time_module

module_logger = logging.getLogger(__name__)
console = logging.StreamHandler()
file_handler = logging.FileHandler("log.txt", "w")
logging.basicConfig(handlers=[file_handler, console],
                    format='%(asctime)-15s: %(name)s: %(levelname)s: %(message)s',
                    level=logging.DEBUG,
                    )


class Automation:
    def __init__(self, master=None):
        self.master = master
        self.username = None
        self.password = None
        self.file = None
        self.yaml_file = r'C:\Gmail_Sender\delay_newsletter_send.yaml'
        self.csv_file = r'C:\Gmail_Sender\send_emailer.csv'
        self.message = None
        self.txt = None
        self.input_message = None
        self.subject = None
        self.limit_count = 1
        self.username_get = None
        self.password_get = None
        self.gmail_user = None
        self.gmail_password = None
        self.win = None
        self.sent_from = None
        self.sent_password = None
        self.body = None
        self.email_limit = None
        self.send_btn = None
        self.retry_label = None
        self.passed_label = None
        self.end_label = None
        self.failed_address = None
        self.send_passed_label = None
        self.retry_send_passed_label = None
        self.retry_failed_address = None
        self.failed_email = []
        self.report_win = None
        self.status_report = None
        self.follow_up_datetime = None
        self.send_successful = 0
        self.send_failed = 0


    @staticmethod
    def about_us():
        logging.info('About Us\n\nVersion 6.12.2020\nWelcome to GMAIL Automatic Send message\n\nSending email one at a time.\n'
                                        'Clearing the cache at 50th email.\nAdding timer of start and end time.\n'
                                        'Wait 24 hours to send again with email limit.\n'
                                        'Multi-Threading to not freeze the program.\n'
                                        'Retry functionality for failed emails.\n'
                                        'Replace %NAME% to the recipient name.\n'
                                        'Security page implemented.\n'
                                        'Got It button bypass.\n'
                                        'Kill program fixed.\n'
                                        'Send HTML format.\n'
                                        'Gmail scheduler for follow up newsletter reiteration.')

    def read_csv(self, filename: str):
        to_list = []
        try:
            with open(filename, encoding='utf-8') as csvfile:
                readCSV = csv.reader(csvfile, delimiter=',')
                for row in readCSV:
                    if 'name' not in str(row).lower() and 'address' not in str(row).lower():
                        to_list.append((row[0], row[1]))
            return to_list
        except Exception as e:
            logging.info('ERROR!!', f'Please make sure CSV file is correct.')

    @staticmethod
    def divide_chunks(l, n):
        # looping till length l
        for i in range(0, len(l), n):
            yield l[i:i + n]

    def google_send(self, to_list):
        g_send = Google(self.sent_from, self.sent_password, to_list, self.subject, self.body,
                        logging, self.limit_count)
        results_output = g_send.run_email()
        self.send_successful += int(results_output['to_list_amount'])

        logging.info(f'Successful email addresses amount: {self.send_successful}')

        logging.info(results_output)

        if results_output['error'] == 'exception':
            logging.error('[google_send] Failed error message')
            senders_email = results_output['failed_email']
            logging.info(f'failed number {int(len(list(dict.fromkeys(senders_email))))}')
            self.send_failed += int(len(list(dict.fromkeys(senders_email))))
            logging.error(results_output['error'])
            logging.error(f'Failed address at {senders_email} amount: {self.send_failed}')
            logging.info(f'Failed email address amount: {self.send_failed}')

            [self.failed_email.append(failed_e) for failed_e in results_output['failed_email']]

    @staticmethod
    def next_time():
        time.sleep(86400)

    def read_yaml(self, file_location):
        with open(file_location, encoding='utf-8') as file:
            # The FullLoader parameter handles the conversion from YAML
            # scalar values to Python the dictionary format
            return yaml.load(file, Loader=yaml.FullLoader)

    def run_future_send(self, email_list=None, delay_date=None):
        if os.path.exists(self.yaml_file):
            if os.path.exists("exceed_email.csv"):
                csv_filename = "exceed_email.csv"
            else:
                csv_filename = self.csv_file
            results = self.read_yaml(self.yaml_file)
            logging.info(f'Running delay email send!\n\n{results}')
            follow_subject = results['subject']
            follow_body = results['body']
            if not delay_date:
                delay_date = results['delay_date']
                self.follow_up_datetime = results['delay_date']
                logging.info(self.follow_up_datetime)
            if not email_list:
                email_list = self.read_csv(csv_filename)

            g_send = Google(self.sent_from, self.sent_password, email_list, follow_subject, follow_body,
                            logging, self.limit_count, True, delay_date)
            return g_send.run_email()

    def send_email(self):
        try:
            results = self.read_yaml(r'C:\Gmail_Sender\newsletter.yaml')
            self.subject = results['subject']
            self.body = results['body']
            self.limit_count = results['email_limit']
            to_list = self.read_csv(self.csv_file)

            time_start = datetime.now()
            logging.info(f'START TIME: {time_start}')
            self.failed_email = []

            try:
                self.google_send(to_list)
                self.run_future_send()

                add_date = datetime.strptime(self.follow_up_datetime, '%m/%d/%Y %H:%M')
                while os.path.exists("exceed_email.csv"):
                    self.next_time()
                    exceed_email = self.read_csv("exceed_email.csv")
                    self.google_send(exceed_email)
                    modified_date = add_date + timedelta(days=1)
                    future_time = datetime.strftime(modified_date, '%m/%d/%Y %H:%M')
                    self.run_future_send(exceed_email, future_time)
                    add_date = future_time
                logging.info(f'Failed email are listed below')
                logging.info(list(dict.fromkeys(self.failed_email)))
                self.send_failed = 0
                self.send_successful = 0

                if self.failed_email:
                    self.next_time()
                    logging.info(f'Re-sending failed email address.')
                    exceed_email = self.read_csv("failed_email.csv")
                    self.google_send(exceed_email)

                while os.path.exists("exceed_email.csv"):
                    self.next_time()
                    exceed_email = self.read_csv("exceed_email.csv")
                    self.google_send(exceed_email)

            except Exception as e:
                logging.error('[send_email] Exception error')
                logging.error(e)
        except Exception as e:
            logging.error(e)
        time_end = datetime.now()
        logging.info(f'END TIME: {time_end}')
        self.file = None
        self.send_successful = 0
        self.send_failed = 0


gmail_sender = Automation()
gmail_sender.send_email()
