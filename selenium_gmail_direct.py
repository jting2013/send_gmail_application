from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from bs4 import BeautifulSoup

import time
from random import randint
import sys, os
import getpass
import subprocess
from clear_cache import ClearCache
import csv
import pyperclip
import logging

# module_logger = logging.getLogger(__name__)
# console = logging.StreamHandler()
# file_handler = logging.FileHandler("selenium.txt", "w")
# logging.basicConfig(handlers=[file_handler, console],
#                     format='%(asctime)-15s: %(name)s: %(levelname)s: %(message)s',
#                     level=logging.INFO,
#                     )


class Google:
    def __init__(self, username, password,  to_address, subject, message, logging, email_limit,
                 follow_up=False, follow_up_datetime: str = None):
        self.username = username
        self.password = password
        self.to_address = to_address
        self.subject = subject
        self.message = message
        self.logging = logging
        self.email_limit = email_limit
        self.follow_up = follow_up
        self.follow_up_datetime = follow_up_datetime
        self.driver = None
        self.wait = None
        self.send_list = None
        self.output_dict = {}
        self.successful_email = []
        self.failed_email = []
        self.total_email = 0
        self.gmail_link = "https://accounts.google.com/ServiceLogin?service=mail&passive=true&rm=false&continue=" \
                          "https://mail.google.com/mail/&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1#identifier"

    @staticmethod
    def divide_chunks(l, n):
        # looping till length l
        for i in range(0, len(l), n):
            yield l[i:i + n]

    def deploy_chrome(self):
        subprocess.call("TASKKILL /f  /IM  CHROME.EXE")
        subprocess.call("TASKKILL /f  /IM  CHROMEDRIVER.EXE")
        try:
            self.logging.info("Entering selenium Gmail")
            local_username = getpass.getuser()
            chrome_options = webdriver.ChromeOptions()
            # chrome_options.headless = True
            # chrome_options.add_argument("--disable-gpu")
            # chrome_options.add_argument("--headless")
            path_local_user = f'C:\\Users\\{local_username}\\AppData\\Local\\Google\\Chrome\\User Data'
            self.logging.info(path_local_user)
            chrome_options.add_argument(f'--user-data-dir={path_local_user}')
            chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
            if getattr(sys, 'frozen', False):
                # executed as a bundled exe, the driver is in the extracted folder
                chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
                self.driver = webdriver.Chrome(chromedriver_path, chrome_options=chrome_options)
            else:
                # executed as a simple script, the driver should be in `PATH`
                self.driver = webdriver.Chrome(chrome_options=chrome_options)

            self.driver.get(self.gmail_link)
            self.driver.maximize_window()
            html = self.driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            self.logging.info(str(soup.prettify()))
            self.logging.info(str(html))
            with open("out.txt", "w") as out:
                out.write(str(soup.prettify("utf-8")))

            self.wait = WebDriverWait(self.driver, 10)

            try:
                login_box = self.wait.until(ec.visibility_of_element_located
                                        ((By.CSS_SELECTOR, 'span[class="RveJvd snByac"]')))
                login_box.click()
                time.sleep(2)
                pw_box = self.wait.until(ec.visibility_of_element_located
                                         ((By.XPATH, '//*[@id="password"]/div[1]/div/div[1]/input')))
                pw_box.send_keys(self.password)
                pw_box.send_keys(Keys.ENTER)
                time.sleep(2)
            except:
                pass
            try:
                got_it = self.wait.until(ec.visibility_of_element_located(
                    (By.CSS_SELECTOR, 'div[class="J-J5-Ji T-I T-I-JN Zd"]')))
                got_it.click()
            except:
                self.wait.until(ec.visibility_of_element_located(
                    (By.CSS_SELECTOR, 'div[class="T-I J-J5-Ji T-I-KE L3"]')))
        except:
            self.driver.save_screenshot("screenfailure_login.png")
            time.sleep(5)

    def write_csv(self, email_list, file_name='exceed_email.csv', write_mode='w'):
        data = dict()
        with open(file_name, mode=write_mode, encoding='utf-8', newline='') as csv_file:
            fieldnames = ['Name', 'Email']
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            writer.writeheader()
            self.logging.info(f'List has this {email_list} and the len is {len(email_list)}')
            for row in email_list:
                data['Name'] = row[0]
                data['Email'] = row[1]
                writer.writerow(data)

    def split_csv_email(self):
        if len(self.to_address) > int(self.email_limit):
            self.write_csv(self.to_address[int(self.email_limit):])
        else:
            if os.path.exists("exceed_email.csv"):
                os.remove("exceed_email.csv")

    def error_output(self, list_chunks):
        self.failed_email.append(list_chunks[1])
        self.send_list.remove(list_chunks)
        self.driver.save_screenshot(f"screenexception_{list_chunks[1]}.png")
        self.output_dict['failed_email'] = self.failed_email
        self.output_dict['error'] = 'exception'
        self.output_dict['to_list_amount'] = self.total_email
        self.output_dict['to_list'] = self.successful_email
        self.write_csv([list_chunks], 'failed_email.csv', 'a')
        self.driver.close()

    def search_follow_up_email(self, email_address: str = None):
        search_box = self.wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[name="q"]')))
        search_box.clear()
        search_box.send_keys(email_address)
        search_box.send_keys(Keys.ENTER)

        table_lookup = self.wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'div[class="BltHke nH oy8Mbf"]')))
        for table_name in table_lookup.find_elements_by_tag_name('table'):
            try:
                self.logging.debug(f'Table class name for delay is: {table_name.get_attribute("class")}')
                if table_name.get_attribute('class') == 'F cf zt' and len(table_name.find_elements_by_tag_name('tr')) > 0:
                    table_name.find_elements_by_tag_name('tr')[0].click()
                    try:
                        reply_btn = self.wait.until(
                            ec.visibility_of_element_located(
                                (By.XPATH,
                                 '/html/body/div[43]/div[3]/div/div[2]/div[1]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[3]/div/table/tr/td[1]/div[2]/div[2]/div/div[3]/div[2]/div/div/div/div/div[2]/div/div/div/div[4]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/div[2]/div[1]/div[4]/table/tbody/tr/td[1]/div/div[2]/div[2]')))
                        reply_btn.click()
                    except:
                        pass
                    try:
                        reply_table = self.wait.until(
                            ec.visibility_of_element_located((By.CSS_SELECTOR, 'table[class="cf wS"')))
                        reply_table.find_elements_by_tag_name('span')[0].click()
                    except:
                        pass

                    try:
                        return self.wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'span[class="mz"]')))
                    except:
                        pass
                    try:
                        send_scheduler = self.wait.until(
                            ec.visibility_of_element_located(
                                (By.CSS_SELECTOR, 'div[class="T-I J-J5-Ji hG T-I-atl L3"]')))

                        send_scheduler.click()
                    except:
                        pass
            except:
                pass



    def follow_up_email(self):
        try:
            send_scheduler = self.wait.until(
                                ec.visibility_of_element_located(
                                    (By.CSS_SELECTOR, 'div[class="T-I J-J5-Ji hG T-I-atl L3"]')))
            send_scheduler.click()
        except:
            pass

        scheduler_btn = self.wait.until(
            ec.visibility_of_element_located(
                (By.CSS_SELECTOR, 'div[class="J-N yr"]')))
        scheduler_btn.click()

        time.sleep(1)
        try:
            scheduler_time_date_button = self.wait.until(
                ec.visibility_of_element_located(
                    (By.CSS_SELECTOR, 'div[class="Az AM"]')))
            scheduler_time_date_button.click()
        except:
            scheduler_time_date = self.wait.until(
            ec.visibility_of_element_located(
                (By.XPATH, '/html/body/div[79]/div[2]/div[2]/div[4]/div[2]')))


            scheduler_time_date.click()

        date_textbox = self.wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[class="hu jA"]')))
        date_textbox.clear()
        date_textbox.send_keys(self.follow_up_datetime.split(' ')[0])

        time_textbox = self.wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[class="hu ks"]')))
        time_textbox.clear()
        time_textbox.send_keys(self.follow_up_datetime.split(' ')[1])
        time_textbox.send_keys(Keys.ENTER)

        time.sleep(5)

    def send_email(self):
        try:
            sched_result = False
            for list_chunks in self.send_list[:int(self.email_limit)]:
                if self.follow_up:
                    sched_result = self.search_follow_up_email(list_chunks[1])
                else:
                    compose = self.wait.until(ec.visibility_of_element_located(
                        (By.CSS_SELECTOR, 'div[class="T-I T-I-KE L3"]')))
                    compose.send_keys(Keys.ENTER)

                    time.sleep(1)
                    email_to = self.wait.until(ec.visibility_of_element_located((By.XPATH, '//textarea[1]')))
                    email_to.send_keys(f'{list_chunks[1]}')
                    pyperclip.copy(self.subject.replace('%NAME%', list_chunks[0]))
                    subject_line = self.wait.until(ec.visibility_of_element_located((By.NAME, 'subjectbox')))
                    subject_line.send_keys(Keys.CONTROL, 'v')
                    time.sleep(1)

                if not sched_result:
                    # pyperclip.copy(self.message.replace('%NAME%', list_chunks[0]))
                    j_body = self.message.replace("%NAME%", list_chunks[0])

                    if self.follow_up:
                        body = self.wait.until(ec.visibility_of_element_located(
                            (By.CSS_SELECTOR, 'div[class="Am aO9 Al editable LW-avf tS-tW"]')))
                    else:
                        body = self.wait.until(
                            ec.visibility_of_element_located((By.CSS_SELECTOR, 'div[class="Am Al editable LW-avf tS-tW"]')))

                    if '<html>' in j_body:
                        self.driver.execute_script(
                            """document.getElementsByClassName("Am Al editable LW-avf tS-tW")[0].innerHTML = '%s'""" %
                            j_body.replace('\n', '').replace("'", '"'), body)
                    elif '</' in j_body:
                        self.driver.execute_script(
                            """document.getElementsByClassName("Am Al editable LW-avf tS-tW")[0].innerHTML = '%s'""" %
                            j_body.replace('\n', '<br><br>').replace("'", '&#39;'), body)
                    else:
                        self.driver.execute_script(
                            """document.getElementsByClassName("Am Al editable LW-avf tS-tW")[0].innerHTML = '%s'""" % j_body.replace(
                                '\n', '<br>').replace("'", '&#39;'), body)

                    if self.follow_up:
                        self.follow_up_email()
                    else:
                        time.sleep(1)
                        body.send_keys(Keys.CONTROL, Keys.ENTER)
                        time.sleep(2)
                    try:
                        self.driver.find_element_by_xpath('//*[@class="Ha"]')
                        self.error_output(list_chunks)
                        return
                    except Exception as e:
                        self.logging.debug(f'Checking for window status {e}')
                        pass

                self.total_email += 1
                if self.total_email % 50 == 0:
                    # time.sleep(5)
                    # clear_cache = ClearCache(self.driver, self.logging)
                    # clear_cache.clear_cache()
                    self.deploy_chrome()
                    # time.sleep(1)
                self.logging.info(f'Total email send: {self.total_email} and last email send was: {list_chunks}')
                self.successful_email.append(list_chunks[1])
                self.send_list.remove(list_chunks)
        except Exception as e:
            self.logging.error(f'Exception error: {e}')
            self.error_output(list_chunks)
            self.send_list.remove(list_chunks)
            return
        self.driver.close()
        self.logging.info(f'Self list still have {self.send_list}')
        self.output_dict['to_list_amount'] = self.total_email
        self.output_dict['to_list'] = self.successful_email
        try:
            if not self.output_dict['error']:
                self.output_dict['error'] = ''
        except Exception as e:
            self.output_dict['error'] = ''
        self.logging.info('*' * 80)
        self.logging.info(self.output_dict)
        self.logging.info('*' * 80)

    def run_email(self):
        self.split_csv_email()
        self.send_list = self.to_address[:int(self.email_limit)]
        while self.send_list:
            try:
                time.sleep(2)
                self.deploy_chrome()
                self.send_email()
            except:
                self.logging.error('run_email failed an exception!!')
                pass

        return self.output_dict
