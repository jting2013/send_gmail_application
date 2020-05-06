from tkinter import *
from tkinter import filedialog, messagebox, ttk
from functools import partial
import smtplib
from random import randint
import time
import os
import csv
import logging
from selenium_gmail_direct import Google
import datetime
import threading


module_logger = logging.getLogger(__name__)
console = logging.StreamHandler()
file_handler = logging.FileHandler("log.txt", "w")
logging.basicConfig(handlers=[file_handler, console],
                    format='%(asctime)-15s: %(name)s: %(levelname)s: %(message)s',
                    level=logging.DEBUG,
                    )

class Window (Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)

        self.master = master
        self.username = None
        self.password = None
        self.file = None
        self.text = None
        self.message = None
        self.txt = None
        self.input_message = None
        self.subject = None
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
        self.failed_email = []
        self.report_win = None
        self.status_report = None
        self.send_successful = 0
        self.send_failed = 0
        self.init_window()

    def init_window(self):
        self.master.title("Gmail Email Application 24 hr with close browser alone with retry failed")

        menu = Menu(self.master)
        self.master.config(menu=menu)

        file = Menu(menu, tearoff=0)
        file.add_command(label='Open', command=self.clicked_dir)
        file.add_command(label='Exit', command=self.master.destroy)
        menu.add_cascade(label='File', menu=file)

        edit = Menu(menu, tearoff=0)
        edit.add_command(label='Date_Time', command=self.showTxt)
        menu.add_cascade(label='Time', menu=edit)

        help = Menu(menu, tearoff=0)
        help.add_command(label='Help Index')
        help.add_command(label='About Us', command=self.about_us)
        menu.add_cascade(label='Help', menu=help)

        # username label and text entry box
        username_label = Label(self.master, text="User Name")
        username_label.grid(row=0, column=0, sticky=E)
        self.username = StringVar()
        username_entry = Entry(self.master, textvariable=self.username, bd =5)
        username_entry.grid(row=0, column=1, sticky=W)

        # password label and password entry box
        password_label = Label(self.master, text="Password")
        password_label.grid(row=1, column=0, sticky=E)
        self.password = StringVar()
        password_entry = Entry(self.master, textvariable=self.password, show='*', bd =5)
        password_entry.grid(row=1, column=1, sticky=W)
        validate_login = partial(self.validate_login, self.username, self.password)

        Button(self.master, text="Login", command=validate_login).grid(row=4, column=0)

        btn2 = Button(root, text="Directory_CSV", command=self.clicked_dir)
        btn2.grid(row=5, column=1, sticky=W)

        email_limit = Label(self.master, text="Email Limit")
        email_limit.grid(row=6, column=0, sticky=E)

        self.email_limit = Entry(self.master, bd=5)
        self.email_limit.grid(row=6, column=1, sticky=W)

        subject = Label(self.master, text="Subject")
        subject.grid(row=7, column=0, sticky=E)

        self.subject = Entry(self.master, bd=5, width=95)
        self.subject.grid(row=7, column=1, sticky=W)

        self.txt = Text(self.master, bd=5)
        self.txt.grid(row=8, column=0, columnspan=2)
        gridframe = Frame(root)

        self.send_btn = Button(gridframe, text="send", command=self.threading_email, bd=5)
        self.send_btn.pack(side=LEFT, padx=5)
        gridframe.grid(row=9, column=1, sticky=W)

    def validate_login(self, username, password):
        self.username_get = self.username.get()
        self.password_get = self.password.get()
        logging.info("username entered :" + self.username_get)
        logging.info("password entered :" + self.password_get)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        server.ehlo()
        try:
            server.login(self.username_get, self.password_get)
            server.close()
            messagebox.showinfo('Success', 'Successfully login!!')
        except:
            messagebox.showinfo('Error!!', 'Make sure Username and Password are correct!\n'
                                           'Make sure GMAIL is also set to "less secure" in Account setting.')

    def clicked_dir(self):
        self.file = filedialog.askopenfile(title="Select file", filetypes=(("CSV Files", "*.csv"),)).name
        print(self.file)

    def clear_message(self):
        self.message.destroy()
        self.file.destroy()

    def showTxt(self):
        self.text = Label(self, text=f'Good morning sir!')
        self.text.grid(row=1, column=3)

    @staticmethod
    def about_us():
        messagebox.showinfo('About Us', 'Welcome to GMAIL Automatic Send message\n\nSending email one at a time.\n'
                                        'Clearing the cache at 50th email.\nAdding timer of start and end time.\n'
                                        'Wait 24 hours to send again with email limit.\n'
                                        'Multi-Threading to not freeze the program.\n'
                                        'Retry functionality for failed emails.\n'
                                        'Replace %NAME% to the recipient name.')

    def read_csv(self, filename=None):
        to_list = []
        if not filename:
            filename = self.file
        try:
            with open(filename) as csvfile:
                readCSV = csv.reader(csvfile, delimiter=',')
                for row in readCSV:
                    if 'name' not in str(row).lower() and 'address' not in str(row).lower():
                        to_list.append((row[0], row[1]))
            return to_list
        except Exception as e:
            messagebox.showinfo('ERROR!!', f'Please make sure CSV file is correct.')

    @staticmethod
    def divide_chunks(l, n):
        # looping till length l
        for i in range(0, len(l), n):
            yield l[i:i + n]

    def pop_up_window(self):
        self.win = Toplevel(root, bd=5)
        self.win.geometry('{}x{}'.format(350, 100))
        self.win.title('Progess is RUNNING')
        win_width = self.win.winfo_width()
        win_height = self.win.winfo_height()
        x = (self.win.winfo_screenwidth() // 2) - (win_width // 2)
        y = (self.win.winfo_screenheight() // 2) - (win_height // 2)
        self.win.geometry('+{}+{}'.format(x, y))
        self.win.transient()
        self.win.resizable(width=False, height=False)
        Label(self.win, text='Please wait...').pack()
        progress_bar = ttk.Progressbar(
            self.win, orient=HORIZONTAL, length=300, mode="indeterminate")
        progress_bar.pack()
        count = 10
        Button(self.win, text="Minimize", command=self.master.iconify, bd=5).pack(side=LEFT, padx=5)
        Button(self.win, text="Maximize", command=self.master.deiconify, bd=5).pack(side=LEFT, padx=5)
        try:
            while True:
                count += 5
                time.sleep(1)
                progress_bar['value'] = count
                self.win.update_idletasks()
        except:
            self.send_btn.config(state='active', relief=RAISED)
            self.win.destroy()

    def report_window(self):
        self.report_win = Toplevel(root, bd=5)
        self.report_win.geometry('{}x{}'.format(450, 100))
        self.report_win.title('Status Report')
        win_width = self.report_win.winfo_width()
        win_height = self.report_win.winfo_height()
        x = (self.report_win.winfo_screenwidth() // 2) - (win_width // 2)
        y = (self.report_win.winfo_screenheight() // 2) - (win_height // 2)
        self.report_win.geometry('+{}+{}'.format(x, y))
        # self.report_win.transient()
        # self.report_win.resizable(width=False, height=False)
        self.status_report = Label(self.report_win, text='Status Report')
        self.status_report.pack()
        vertScrollbar = Scrollbar(self.report_win, orient='vertical')
        vertScrollbar.pack(side='right', fill='y')

    def verify_component(self):
        subject_failed = False
        body_failed = False
        to_list = self.read_csv()
        if to_list:
            self.username_get = self.username.get()
            self.password_get = self.password.get()
            if self.username_get and self.password_get:
                self.sent_from = str(self.username_get)
                self.sent_password = str(self.password_get)
            else:
                self.sent_from = str(self.gmail_user)
                self.sent_password = str(self.gmail_password)
            try:
                self.subject = str(self.subject.get())
            except:
                pass
            if not self.subject:
                messagebox.showinfo('ERROR!!', 'Please make sure there is a SUBJECT.')
                subject_failed = True

            self.body = str(self.txt.get(1.0, "end-1c"))
            if not self.body:
                messagebox.showinfo('ERROR!!', 'Please make sure there is a BODY.')
                body_failed = True
            logging.info(self.body)
            return to_list, subject_failed, body_failed

    def google_send(self, to_list, send_message=None, fail_message=None):
        if not send_message:
            send_message = {'Send': 10}
        if not fail_message:
            fail_message = {'Failed': 11}

        if self.email_limit.get():
            limit_count = self.email_limit.get()
        else:
            limit_count = 500
        g_send = Google(self.sent_from, self.sent_password, to_list, self.subject, self.body,
                                           logging, limit_count)
        results_output = g_send.run_email()
        self.send_successful += int(results_output['to_list_amount'])

        for k, v in send_message.items():
            passed_label = Label(self.master, text=f'{k} email addresses '
                                                   f'amount: {self.send_successful}')
            row_number = send_message.values()
            passed_label.grid(row=int(v), column=0, columnspan=2, sticky=W)

        # now_time = datetime.datetime.now()
        # report_passed = Label(self.report_win, text=f'[{now_time}] Send email address amount: '
        #                                             f'{results_output["to_list_amount"]}')
        # report_passed.pack(expand=True)
        logging.info(results_output)

        if results_output['error'] == 'exception':
            logging.error('[google_send] Failed error message')
            senders_email = results_output['failed_email']
            logging.info(f'failed number {int(len(list(dict.fromkeys(senders_email))))}')
            self.send_failed += int(len(list(dict.fromkeys(senders_email))))
            failed_address = Label(self.master, text=f'{fail_message.keys()} email address amount: {self.send_failed}')
            logging.error(results_output['error'])
            logging.error(f'Failed address at {senders_email} amount: {self.send_failed}')
            for k, v in fail_message.items():
                failed_address = Label(self.master,
                                       text=f'{k} email address amount: {self.send_failed}')
                failed_address.grid(row=v, column=0, columnspan=2, sticky=W)
            [self.failed_email.append(failed_e) for failed_e in results_output['failed_email']]
            # failed_report = Label(self.report_win, text=f'[{now_time}] Failed email address amount: '
            #                                             f'{len(list(dict.fromkeys(senders_email)))}')
            # failed_report.pack(expand=True)

    @staticmethod
    def next_time(time_start):
        send_end_time = datetime.datetime.now()
        calculate_time_delay = (send_end_time - time_start).total_seconds()
        time.sleep(86400 - int(calculate_time_delay))
        # time.sleep(300 - int(calculate_time_delay))

        return datetime.datetime.now()

    def send_email(self):
        self.send_btn.config(state='disable', relief=SUNKEN)
        try:
            to_list, subject_failed, body_failed = self.verify_component()
            if not subject_failed and not body_failed:
                time_start = datetime.datetime.now()
                passed_label = Label(self.master, text=f'START TIME: {time_start}')
                passed_label.grid(row=12, column=0, columnspan=2, sticky=W)

                try:
                    headers = [
                        "From: " + self.sent_from,
                        "Subject: " + self.subject,
                        "To: " + "",
                        "MIME-Version: 1.0",
                        "Content-Type: text/html"]
                    headers = "\r\n".join(headers)
                    message = headers + "\r\n\r\n" + self.body
                    logging.info(f'sent_from {self.sent_from} \nto_list: {to_list}\nmessage: {message}')
                    self.google_send(to_list)

                    while os.path.exists("exceed_email.csv"):
                        time_start = self.next_time(time_start)
                        exceed_email = self.read_csv("exceed_email.csv")
                        self.google_send(exceed_email)
                    logging.info(f'Failed email are listed below')
                    logging.info(list(dict.fromkeys(self.failed_email)))
                    self.send_failed = 0
                    self.send_successful = 0
                    send_message = {'RETRY Send': 15}
                    fail_message = {'RETRY Failed': 16}
                    if self.failed_email:
                        time_start = self.next_time(time_start)
                        retry_label = Label(self.master, text=f'Re-sending failed email address.')
                        retry_label.grid(row=14, column=0, columnspan=2, sticky=W)
                        exceed_email = self.read_csv("failed_email.csv")
                        self.google_send(exceed_email, send_message, fail_message)
                    while os.path.exists("exceed_email.csv"):
                        time_start = self.next_time(time_start)
                        exceed_email = self.read_csv("exceed_email.csv")
                        self.google_send(exceed_email, send_message, fail_message)
                except Exception as e:
                    logging.error('[send_email] Exception error')
                    # logging.error(f'{self.sent_from} and {self.username_get}')
                    logging.error(e)
        except Exception as e:
            logging.error(e)
        time_end = datetime.datetime.now()
        passed_label = Label(self.master, text=f'END TIME: {time_end}')
        passed_label.grid(row=13, column=0, columnspan=2, sticky=W)
        self.file = None
        self.win.destroy()
        self.send_successful = 0
        self.send_failed = 0

    def threading_email(self):
        if os.path.exists("failed_email.csv"):
            os.remove("failed_email.csv")
        if os.path.exists("exceed_email.csv"):
            os.remove("exceed_email.csv")
        thread1 = threading.Thread(target=self.pop_up_window, args=())
        thread2 = threading.Thread(target=self.send_email, args=())
        thread1.start()
        thread2.start()


root = Tk()
root_width = root.winfo_width()
root_height = root.winfo_height()
x = int((root.winfo_screenwidth() / 4) - (root_width / 2))
y = int((root.winfo_screenheight() / 4) - (root_height / 2))
root.geometry('+{}+{}'.format(x, y))
root.resizable(width=False, height=False)
app = Window(root)
root.mainloop()
