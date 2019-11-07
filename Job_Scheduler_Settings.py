# Global Module import
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkcalendar import DateEntry
from dateutil import relativedelta
from Global import grabobjs
from Global import CryptHandle
from Global import ShelfHandle
from math import floor

# This is needed when compiling .Exe since compiler has issues including hidden imported modules (Only for TkCalender)
import babel.numbers

import smtplib
import datetime
import os
import copy
import pyperclip
import pathlib as pl
import sys
import portalocker

if getattr(sys, 'frozen', False):
    application_path = sys.executable
    ico_dir = sys._MEIPASS
else:
    application_path = __file__
    ico_dir = os.path.dirname(__file__)

# Global Variable declaration
curr_dir = os.path.dirname(os.path.abspath(application_path))
main_dir = os.path.dirname(curr_dir)
joblogsdir = os.path.join(main_dir, '05_Job_Logs')
joblogsexpdir = os.path.join(main_dir, '06_Job_Logs_Export')
global_objs = grabobjs(main_dir, 'Job_Scheduler')
icon_path = os.path.join(ico_dir, '%s.ico' % os.path.splitext(os.path.basename(application_path))[0])


class SettingsGUI:
    eupass_txtbox = None
    euname_txtbox = None
    eport_txtbox = None
    eserver_txtbox = None
    view_jobs_button = None
    save_settings_button = None
    view_jobs_obj = None

    # Function that is executed upon creation of SettingsGUI class
    def __init__(self):
        self.header_text = "Welcome to Job Scheduler Settings!\nPlease add or modify network/email settings below"

        self.email_upass_obj = global_objs['Settings'].grab_item('Email_Pass')
        self.asql = global_objs['SQL']
        self.main = Tk()
        self.main.iconbitmap(icon_path)

        # GUI Variables
        self.server = StringVar()
        self.database = StringVar()
        self.email_server = StringVar()
        self.email_port = StringVar()
        self.email_user_name = StringVar()
        self.email_user_pass = StringVar()

        self.main.bind('<Destroy>', self.gui_cleanup)

    def gui_cleanup(self, event):
        self.asql.close()

    # Function to build GUI for settings
    def build_gui(self, header=None):
        # Change to custom header title if specified
        if header:
            self.header_text = header

        # Set GUI Geometry and GUI Title
        self.main.geometry('420x205+500+90')
        self.main.title('Job Scheduler Settings')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        network_frame = LabelFrame(self.main, text='Network Settings', width=508, height=70)
        email_frame = LabelFrame(self.main, text='Email Settings', width=508, height=140)
        buttons_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        network_frame.pack(fill="both")
        email_frame.pack(fill="both")
        buttons_frame.pack(fill='both')

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header_text, width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Network Labels & Input boxes to the Network_Frame
        #     SQL Server Input Box
        server_label = Label(self.main, text='Server:')
        server_txtbox = Entry(self.main, textvariable=self.server, width=20)
        server_label.pack(in_=network_frame, side=LEFT, pady=4, padx=12)
        server_txtbox.pack(in_=network_frame, side=LEFT, pady=4, padx=1)
        server_txtbox.bind('<FocusOut>', self.check_network)

        #     Server Database Input Box
        database_label = Label(self.main, text='Database:')
        database_txtbox = Entry(self.main, textvariable=self.database, width=20)
        database_txtbox.pack(in_=network_frame, side=RIGHT, pady=4, padx=12)
        database_label.pack(in_=network_frame, side=RIGHT, pady=4, padx=3)
        database_txtbox.bind('<KeyRelease>', self.check_network)

        # Apply Email Labels & Input boxes to the EConn_Frame
        #     Email Server Input Box
        eserver_label = Label(email_frame, text='Server:')
        self.eserver_txtbox = Entry(email_frame, textvariable=self.email_server, width=20)
        eserver_label.grid(row=0, column=0, padx=4, pady=5, sticky='e')
        self.eserver_txtbox.grid(row=0, column=1, padx=6, pady=5, sticky='e')

        #     Email Port Input Box
        eport_label = Label(email_frame, text='Port:')
        self.eport_txtbox = Entry(email_frame, textvariable=self.email_port, width=20)
        eport_label.grid(row=0, column=2, padx=4, pady=5, sticky='e')
        self.eport_txtbox.grid(row=0, column=3, padx=6, pady=5, sticky='w')

        #     Email User Name Input Box
        euname_label = Label(email_frame, text='User Name:')
        self.euname_txtbox = Entry(email_frame, textvariable=self.email_user_name, width=20)
        euname_label.grid(row=1, column=0, padx=4, pady=5, sticky='e')
        self.euname_txtbox.grid(row=1, column=1, padx=6, pady=5, sticky='e')

        #     Email User Pass Input Box
        eupass_label = Label(email_frame, text='User Pass:')
        self.eupass_txtbox = Entry(email_frame, textvariable=self.email_user_pass, width=20)
        eupass_label.grid(row=1, column=2, padx=4, pady=5, sticky='e')
        self.eupass_txtbox.grid(row=1, column=3, padx=6, pady=5, sticky='w')
        self.eupass_txtbox.bind('<KeyRelease>', self.hide_pass)

        # Apply Buttons to the Buttons Frame
        #     Save Settings
        self.save_settings_button = Button(buttons_frame, text='Save Settings', width=15, command=self.save_settings)
        self.save_settings_button.grid(row=0, column=0, pady=6, padx=15, sticky='w')

        #     View Job Scheduler
        self.view_jobs_button = Button(buttons_frame, text='View Job(s)', width=15, command=self.view_jobs)
        self.view_jobs_button.grid(row=0, column=1, pady=6, padx=10)

        #     Cancel Button
        cancel_button = Button(buttons_frame, text='Cancel', width=15, command=self.cancel)
        cancel_button.grid(row=0, column=2, pady=6, padx=15, sticky='e')

        # Fill GUI
        self.fill_gui()

        # Show dialog
        self.main.mainloop()

    # Function to fill GUI textbox fields
    def fill_gui(self):
        fill_textbox('Settings', self.server, 'Server')
        fill_textbox('Settings', self.database, 'Database')

        if not self.server.get() or not self.database.get() or not self.asql.test_conn('alch'):
            self.eupass_txtbox.configure(state=DISABLED)
            self.euname_txtbox.configure(state=DISABLED)
            self.eport_txtbox.configure(state=DISABLED)
            self.eserver_txtbox.configure(state=DISABLED)
            self.save_settings_button.configure(state=DISABLED)
            self.view_jobs_button.configure(state=DISABLED)
        else:
            fill_textbox('Settings', self.email_server, 'Email_Server')
            fill_textbox('Settings', self.email_port, 'Email_Port')
            fill_textbox('Settings', self.email_user_name, 'Email_User')

            if not self.email_server.get():
                self.email_server.set('imail.granitenet.com')

            if not self.email_port.get():
                self.email_port.set('587')

            if self.email_upass_obj and isinstance(self.email_upass_obj, CryptHandle):
                self.email_user_pass.set('*' * len(self.email_upass_obj.decrypt_text()))

    def hide_pass(self, event):
        if self.email_upass_obj:
            currpass = self.email_upass_obj.decrypt_text()

            if len(self.email_user_pass.get()) > len(currpass):
                i = 0

                for letter in self.email_user_pass.get():
                    if letter != '*':
                        if i > len(currpass) - 1:
                            currpass += letter
                        else:
                            mytext = list(currpass)
                            mytext.insert(i, letter)
                            currpass = ''.join(mytext)
                    i += 1
            elif len(self.email_user_pass.get()) > 0:
                i = 0

                for letter in self.email_user_pass.get():
                    if letter != '*':
                        mytext = list(currpass)
                        mytext[i] = letter
                        currpass = ''.join(mytext)
                    i += 1

                if len(currpass) - i > 0:
                    currpass = currpass[:i]
            else:
                currpass = None

            if currpass:
                self.email_upass_obj.encrypt_text(currpass)
                self.email_user_pass.set('*' * len(self.email_upass_obj.decrypt_text()))
            else:
                self.email_upass_obj = None
                self.email_user_pass.set("")
        else:
            self.email_upass_obj = CryptHandle()
            self.email_upass_obj.encrypt_text(self.email_user_pass.get())
            self.email_user_pass.set('*' * len(self.email_upass_obj.decrypt_text()))

    # Function to check network settings if populated
    def check_network(self, event):
        state = False

        if self.server.get() and self.database.get() and \
                (global_objs['Settings'].grab_item('Server') != self.server.get() or
                 global_objs['Settings'].grab_item('Database') != self.database.get()):
            self.asql.change_config(server=self.server.get(), database=self.database.get())

            if self.asql.test_conn('alch'):
                add_setting('Settings', self.server.get(), 'Server')
                add_setting('Settings', self.database.get(), 'Database')
                state = True

        if state:
            self.eupass_txtbox.configure(state=NORMAL)
            self.euname_txtbox.configure(state=NORMAL)
            self.eport_txtbox.configure(state=NORMAL)
            self.eserver_txtbox.configure(state=NORMAL)
            self.save_settings_button.configure(state=NORMAL)
            self.view_jobs_button.configure(state=NORMAL)
        else:
            self.eupass_txtbox.configure(state=DISABLED)
            self.euname_txtbox.configure(state=DISABLED)
            self.eport_txtbox.configure(state=DISABLED)
            self.eserver_txtbox.configure(state=DISABLED)
            self.save_settings_button.configure(state=DISABLED)
            self.view_jobs_button.configure(state=DISABLED)

    # Function to save settings when the Save Settings button is pressed
    def save_settings(self):
        if not self.email_server.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Email Server',
                                 parent=self.main)
        elif not self.email_port.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Email Port',
                                 parent=self.main)
        elif not self.email_user_name.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Email User Name',
                                 parent=self.main)
        elif not self.email_user_pass.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Email User Pass',
                                 parent=self.main)
        elif not str(self.email_port.get()).isnumeric():
            messagebox.showerror('Field Error!', 'Email Port is a non-numeric port',
                                 parent=self.main)
        else:
            email_err = 0

            try:
                server = smtplib.SMTP(str(self.email_server.get()), int(self.email_port.get()))

                try:
                    server.ehlo()
                    server.starttls()
                    server.ehlo()
                    server.login(self.email_user_name.get(), self.email_upass_obj.decrypt_text())
                except:
                    email_err = 2
                    pass
                finally:
                    server.close()
            except:
                email_err = 1
                pass

            if email_err == 1:
                messagebox.showerror('Invalid Email Settings!',
                                     'Server and/or Port does not exist. Please specify the correct information',
                                     parent=self.main)
            elif email_err == 2:
                messagebox.showerror('Invalid Email Settings!',
                                     'User name and/or user pass is invalid. Please specify the correct information',
                                     parent=self.main)
            else:
                add_setting('Settings', self.email_server.get(), 'Email_Server')
                add_setting('Settings', self.email_port.get(), 'Email_Port')
                add_setting('Settings', self.email_user_name.get(), 'Email_User')
                add_setting('Settings', self.email_upass_obj.decrypt_text(), 'Email_Pass')

                self.main.destroy()

    def view_jobs(self):
        if self.view_jobs_obj:
            self.view_jobs_obj.cancel()

        self.view_jobs_obj = JobListGUI(self.main)
        self.view_jobs_obj.build_gui()

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        self.main.destroy()


class JobGUI:
    js_date_entry = None
    js_list_box = None
    js_task_label = None
    js_fprefix_label = None
    js_fprefix_entry = None
    jst_list_box = None
    js_file_dir = None
    js_attach_chkbox = None
    js_monday_chkbox = None
    js_tuesday_chkbox = None
    js_wednesday_chkbox = None
    js_thursday_chkbox = None
    js_friday_chkbox = None
    js_saturday_chkbox = None
    js_sunday_chkbox = None
    js_shellcomm_entry = None
    js_list_sel = 0
    jst_list_sel = 0

    # Function that is executed upon creation of JobGUI class
    def __init__(self, class_obj, root, job=None):
        self.root = root
        self.main = Toplevel(root)
        self.job = job
        self.class_obj = class_obj
        self.destroy = False
        self.main.iconbitmap(icon_path)

        if job:
            self.title = 'Modify Existing Job'
            self.header = "Please Modify/Delete Job below.\nPress 'Modify Job' or 'Delete Job' when finished"
        else:
            self.title = 'Add New Job'
            self.header = 'Please Add a new Job below.\nPress ''Add Job'' when finished'

        self.freq_list = list()
        self.freq_days_of_week_list = list()
        self.freq_start_dt_list = list()
        self.freq_missed_run_list = list()
        self.freq_prev_run_list = list()
        self.freq_next_run_list = list()
        self.sub_job_list = list()
        self.sub_job_type_list = list()

        # GUI Variables
        self.job_name = StringVar()
        self.timeout_hh = StringVar()
        self.timeout_mm = StringVar()
        self.email_to = StringVar()
        self.email_cc = StringVar()
        self.js_date = StringVar()
        self.js_fprefix = StringVar()
        self.js_task = StringVar()
        self.job_mm = StringVar()
        self.job_hh = StringVar()
        self.js_shell_comm = StringVar()
        self.js_frequency = IntVar()
        self.js_task_type = IntVar()
        self.js_attach = IntVar()
        self.js_missed_run = IntVar()
        self.js_monday = IntVar()
        self.js_tuesday = IntVar()
        self.js_wednesday = IntVar()
        self.js_thursday = IntVar()
        self.js_friday = IntVar()
        self.js_saturday = IntVar()
        self.js_sunday = IntVar()

        self.main.bind('<Destroy>', self.gui_cleanup)

    def gui_cleanup(self, event):
        if self.root and not self.destroy:
            self.destroy = True
            self.class_obj.fill_gui(True)

    # Function to build GUI for Add/Del/Modify a job
    def build_gui(self):
        if self.job:
            cancel_xpos = 15
            button_name = "Modify Job"
        else:
            cancel_xpos = 232
            button_name = "Add Job"

        # Set GUI Geometry and GUI Title
        self.main.geometry('618x510+500+90')
        self.main.title(self.title)
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        job_frame = LabelFrame(self.main, text='Job', width=508, height=240)
        jname_frame = Frame(job_frame)
        edistro_frame = LabelFrame(job_frame, text='Email Distro(s)', width=508, height=70)
        jschedule_frame = LabelFrame(job_frame, text='Job Schedule(s)', width=618, height=200)
        jschedule_lframe = Frame(jschedule_frame)
        jschedule_rframe = Frame(jschedule_frame)
        jschedule_subframe = Frame(jschedule_lframe)
        jschedule_subframe2 = Frame(jschedule_lframe)
        jschedule_subframe3 = Frame(jschedule_rframe)
        jschedule_subframe4 = Frame(jschedule_rframe)
        jschedule_subframe5 = Frame(jschedule_lframe)
        jtask_frame = LabelFrame(job_frame, text='Job Task(s)', width=508, height=70)
        jtask_lframe = Frame(jtask_frame)
        jtask_rframe = Frame(jtask_frame)
        jtask_subframe = Frame(jtask_lframe)
        jtask_subframe2 = Frame(jtask_lframe)
        jtask_subframe3 = Frame(jtask_rframe)
        jtask_subframe4 = Frame(jtask_rframe)
        buttons_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        job_frame.pack(fill='both')
        jname_frame.grid(row=0, column=0, columnspan=2, sticky=E + W)
        edistro_frame.grid(row=1, column=0, columnspan=2, sticky=E + W)
        jschedule_frame.grid(row=2, column=0, sticky=N + S + W + E)
        jschedule_lframe.grid(row=0, column=0, sticky=W)
        jschedule_rframe.grid(row=0, column=1, sticky=E + W + N + S)
        jschedule_subframe.grid(row=0, column=0)
        jschedule_subframe5.grid(row=1, column=0)
        jschedule_subframe2.grid(row=2, column=0)
        jschedule_subframe3.grid(row=0, column=0, sticky=E + W)
        jschedule_subframe4.grid(row=1, column=0, sticky=E + W)
        jtask_frame.grid(row=2, column=1, sticky=N + S + W + E)
        jtask_lframe.grid(row=0, column=0, sticky=W)
        jtask_rframe.grid(row=0, column=1, sticky=W + N + S)
        jtask_subframe.grid(row=0, column=0)
        jtask_subframe2.grid(row=1, column=0)
        jtask_subframe3.grid(row=0, column=0, sticky=N + S)
        jtask_subframe4.grid(row=1, column=0)
        buttons_frame.pack(fill='both')

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header, width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply widgets to Job Frame
        #     Job Name Input Box
        jname_label = Label(jname_frame, text='Job Name:')
        jname_txtbox = Entry(jname_frame, textvariable=self.job_name, width=48)
        jname_label.grid(row=0, column=0, padx=6, pady=5, sticky='e')
        jname_txtbox.grid(row=0, column=1, padx=0, pady=5, sticky='e')

        #     Timeout Dropdown Menu's
        timeout_label = Label(jname_frame, text="Timeout:")
        to_hh_dropdown = OptionMenu(jname_frame, self.timeout_hh, *range(0, 24))
        to_hh_label = Label(jname_frame, text='HH')
        to_mm_dropdown = OptionMenu(jname_frame, self.timeout_mm, *range(0, 60))
        to_mm_label = Label(jname_frame, text='MM')
        timeout_label.grid(row=0, column=2, padx=13, pady=5, sticky='w')
        to_hh_dropdown.grid(row=0, column=3, padx=0, pady=5, sticky='w')
        to_hh_label.grid(row=0, column=4, padx=0, pady=5, sticky='w')
        to_mm_dropdown.grid(row=0, column=5, padx=0, pady=5, sticky='w')
        to_mm_label.grid(row=0, column=6, padx=0, pady=5, sticky='w')

        # Apply widgets to EDistro Frame
        #     Email To Input Box
        eto_label = Label(edistro_frame, text='Email To:')
        eto_txtbox = Entry(edistro_frame, textvariable=self.email_to, width=34)
        eto_label.grid(row=0, column=0, padx=8, pady=5, sticky='e')
        eto_txtbox.grid(row=0, column=1, padx=13, pady=5, sticky='e')

        #     Email CC Input Box
        ecc_label = Label(edistro_frame, text='Email Cc:')
        ecc_txtbox = Entry(edistro_frame, textvariable=self.email_cc, width=34)
        ecc_label.grid(row=0, column=2, padx=8, pady=5, sticky='w')
        ecc_txtbox.grid(row=0, column=3, padx=13, pady=5, sticky='w')

        # Apply widgets to Job Schedule Frame
        #     Start Date Date Entry
        js_date_label = Label(jschedule_subframe, text='Start Date:')
        self.js_date_entry = DateEntry(jschedule_subframe, textvariable=self.js_date, width=12,
                                       background='darkblue', foreground='white', borderwidth=2)
        js_date_label.grid(row=0, column=0, padx=4, pady=5, sticky='e')
        self.js_date_entry.grid(row=0, column=1, columnspan=3, padx=2, pady=5, sticky='w')

        js_hh_dropdown = OptionMenu(jschedule_subframe, self.job_hh, *range(0, 24))
        js_hh_label = Label(jschedule_subframe, text='HH')
        js_mm_dropdown = OptionMenu(jschedule_subframe, self.job_mm, *range(0, 60))
        js_mm_label = Label(jschedule_subframe, text='MM')
        js_hh_dropdown.grid(row=1, column=0, padx=0, pady=5, sticky='e')
        js_hh_label.grid(row=1, column=1, padx=0, pady=5, sticky='e')
        js_mm_dropdown.grid(row=1, column=2, padx=0, pady=5, sticky='e')
        js_mm_label.grid(row=1, column=3, padx=0, pady=5, sticky='e')

        #     Frequency Radio Buttons
        js_freq_rad1 = Radiobutton(jschedule_subframe5, text='Daily', variable=self.js_frequency, value=0,
                                   command=lambda: self.job_freq_toggle(None))
        js_freq_rad2 = Radiobutton(jschedule_subframe5, text='Weekly', variable=self.js_frequency, value=1,
                                   command=lambda: self.job_freq_toggle(None))
        js_freq_rad3 = Radiobutton(jschedule_subframe5, text='Bi-Weekly', variable=self.js_frequency, value=2,
                                   command=lambda: self.job_freq_toggle(None))
        js_freq_rad4 = Radiobutton(jschedule_subframe5, text='Monthly', variable=self.js_frequency, value=3,
                                   command=lambda: self.job_freq_toggle(None))
        js_freq_rad1.grid(row=0, column=0, padx=2, pady=4, sticky='w')
        js_freq_rad2.grid(row=0, column=1, padx=2, pady=4, sticky='w')
        js_freq_rad3.grid(row=1, column=0, padx=2, pady=4, sticky='w')
        js_freq_rad4.grid(row=1, column=1, padx=2, pady=4, sticky='w')

        #    Include Stored Procedure Attachments Checkbox
        js_missed_run_chkbox = Checkbutton(jschedule_subframe5, text='Exec Missed Run on Startup',
                                           variable=self.js_missed_run, wraplength=160) # was 80 for wraplength
        js_missed_run_chkbox.grid(row=2, column=0, columnspan=2, padx=2, pady=4, sticky='w')

        #     Job Schedule List Box
        js_yscrollbar = Scrollbar(jschedule_subframe2, orient="vertical")
        js_xscrollbar = Scrollbar(jschedule_subframe2, orient="horizontal")
        self.js_list_box = Listbox(jschedule_subframe2, selectmode=SINGLE, width=25, height=6,
                                   yscrollcommand=js_yscrollbar.set, xscrollcommand=js_xscrollbar.set)
        js_yscrollbar.config(command=self.js_list_box.yview)
        js_xscrollbar.config(command=self.js_list_box.xview)
        self.js_list_box.grid(row=0, column=0, rowspan=4, padx=8, pady=5)
        js_yscrollbar.grid(row=0, column=1, rowspan=4, sticky=N + S)
        js_xscrollbar.grid(row=4, column=0, columnspan=2, sticky=E + W)
        self.js_list_box.bind("<Down>", self.js_list_down)
        self.js_list_box.bind("<Up>", self.js_list_up)
        self.js_list_box.bind('<<ListboxSelect>>', self.js_list_select)

        #     Add/Del Buttons
        js_add_button = Button(jschedule_subframe3, text='Add', width=4, command=self.js_add)
        js_del_button = Button(jschedule_subframe3, text='Del', width=4, command=self.js_del)
        js_mod_button = Button(jschedule_subframe3, text='Modify', width=11, command=self.js_mod)
        js_add_button.grid(row=0, column=0, sticky='w', pady=5)
        js_del_button.grid(row=0, column=1, sticky='w', padx=6, pady=5)
        js_mod_button.grid(row=1, column=0, columnspan=2, sticky='w', pady=5)

        #    Include Monday - Sunday Checkboxes
        self.js_monday_chkbox = Checkbutton(jschedule_subframe4, text='Monday', variable=self.js_monday)
        self.js_tuesday_chkbox = Checkbutton(jschedule_subframe4, text='Tuesday', variable=self.js_tuesday)
        self.js_wednesday_chkbox = Checkbutton(jschedule_subframe4, text='Wednesday', variable=self.js_wednesday)
        self.js_thursday_chkbox = Checkbutton(jschedule_subframe4, text='Thursday', variable=self.js_thursday)
        self.js_friday_chkbox = Checkbutton(jschedule_subframe4, text='Friday', variable=self.js_friday)
        self.js_saturday_chkbox = Checkbutton(jschedule_subframe4, text='Saturday', variable=self.js_saturday)
        self.js_sunday_chkbox = Checkbutton(jschedule_subframe4, text='Sunday', variable=self.js_sunday)

        # Apply widgets to Job Task Frame
        #     Job Exec Filepath/Stored Procedure Entry
        self.js_task_label = Label(jtask_subframe, text='Filepath:')
        js_task_entry = Entry(jtask_subframe, textvariable=self.js_task, width=15)
        self.js_task_label.grid(row=0, column=0, padx=3, pady=5, sticky='e')
        js_task_entry.grid(row=0, column=1, padx=7, pady=5, sticky='w')

        #     Params/Filename Prefix Entry
        self.js_fprefix_label = Label(jtask_subframe, text='Param(s):')
        self.js_fprefix_entry = Entry(jtask_subframe, textvariable=self.js_fprefix, width=15)
        self.js_fprefix_label.grid(row=1, column=0, padx=2, pady=5, sticky='e')
        self.js_fprefix_entry.grid(row=1, column=1, padx=7, pady=5, sticky='w')

        #     PowerShell Command Entry Box
        js_shellcomm_label = Label(jtask_subframe, text='Shell Com:')
        self.js_shellcomm_entry = Entry(jtask_subframe, textvariable=self.js_shell_comm, width=15)
        js_shellcomm_label.grid(row=2, column=0, padx=2, pady=5, sticky='e')
        self.js_shellcomm_entry.grid(row=2, column=1, padx=7, pady=5, sticky='w')

        #     Job Task List Box
        jst_yscrollbar = Scrollbar(jtask_subframe2, orient="vertical")
        jst_xscrollbar = Scrollbar(jtask_subframe2, orient="horizontal")
        self.jst_list_box = Listbox(jtask_subframe2, selectmode=SINGLE, width=25, height=11,
                                    yscrollcommand=jst_yscrollbar.set, xscrollcommand=jst_xscrollbar.set)
        jst_yscrollbar.config(command=self.jst_list_box.yview)
        jst_xscrollbar.config(command=self.jst_list_box.xview)
        self.jst_list_box.grid(row=0, column=0, rowspan=4, padx=8, pady=5)
        jst_yscrollbar.grid(row=0, column=1, rowspan=4, sticky=N + S)
        jst_xscrollbar.grid(row=4, column=0, columnspan=2, sticky=E + W)
        self.jst_list_box.bind("<Down>", self.jst_list_down)
        self.jst_list_box.bind("<Up>", self.jst_list_up)
        self.jst_list_box.bind('<<ListboxSelect>>', self.jst_list_select)

        #    Include Stored Procedure Attachments Checkbox
        self.js_file_dir = Button(jtask_subframe3, text='Find File', width=13, command=self.jst_dir)
        self.js_file_dir.grid(row=0, column=0, padx=3, pady=4, sticky='w')

        #    Include Stored Procedure Attachments Checkbox
        self.js_attach_chkbox = Checkbutton(jtask_subframe3, text='Attach SP File', variable=self.js_attach,
                                            command=lambda: self.job_attach_toggle(None))

        #     Job Task Type Radio Buttons
        js_task_type_rad1 = Radiobutton(jtask_subframe3, text='Exec Program', variable=self.js_task_type,
                                        value=0, command=lambda: self.job_task_type_toggle(None))
        js_task_type_rad2 = Radiobutton(jtask_subframe3, text='Exec Stored Proc', variable=self.js_task_type,
                                        value=1, command=lambda: self.job_task_type_toggle(None))
        js_task_type_rad1.grid(row=1, column=0, padx=3, pady=4, sticky='w')
        js_task_type_rad2.grid(row=2, column=0, padx=3, pady=4, sticky='w')

        #     Add/Del Buttons
        jst_add_button = Button(jtask_subframe4, text='Add', width=5, command=self.jst_add)
        jst_del_button = Button(jtask_subframe4, text='Del', width=5, command=self.jst_del)
        jst_mod_button = Button(jtask_subframe4, text='Modify Task', width=13, command=self.jst_mod)
        jst_up_button = Button(jtask_subframe4, text='Task Up', width=13, command=self.jst_up)
        jst_down_button = Button(jtask_subframe4, text='Task Down', width=13, command=self.jst_down)
        jst_add_button.grid(row=0, column=0, sticky='w', pady=5)
        jst_del_button.grid(row=0, column=1, sticky='w', pady=5, padx=6)
        jst_mod_button.grid(row=1, column=0, columnspan=2, sticky='w', pady=5)
        jst_up_button.grid(row=2, column=0, columnspan=2, sticky='w', pady=5)
        jst_down_button.grid(row=3, column=0, columnspan=2, sticky='w', pady=5)

        # Apply Buttons to the Buttons Frame
        #     Add Job
        add_job_button = Button(buttons_frame, text=button_name, width=23, command=self.add_job)
        add_job_button.grid(row=0, column=0, pady=6, padx=15, sticky='w')

        if self.job:
            #     View Jobs
            view_jobs_button = Button(buttons_frame, text='Delete Job', width=23, command=self.del_job)
            view_jobs_button.grid(row=0, column=1, pady=6, padx=10)

        #     Cancel Button
        cancel_button = Button(buttons_frame, text='Cancel', width=23, command=self.cancel)
        cancel_button.grid(row=0, column=2, pady=6, padx=cancel_xpos, sticky='w')

        self.fill_gui()

    # Function to fill GUI textbox fields
    def fill_gui(self):
        if self.job:
            job_schedule = copy.deepcopy(self.job['Job_Schedule'])
            job_list = copy.deepcopy(self.job['Job_List'])

            self.job_name.set(self.job['Job_Name'])
            self.email_to.set(self.job['To_Distro'])
            self.email_cc.set(self.job['CC_Distro'])
            self.timeout_hh.set(self.job['Job_Timeout'][0])
            self.timeout_mm.set(self.job['Job_Timeout'][1])

            self.freq_list, self.freq_days_of_week_list, self.freq_start_dt_list, self.freq_missed_run_list,\
            self.freq_prev_run_list, self.freq_next_run_list = zip(*job_schedule)
            self.sub_job_list, self.sub_job_type_list = zip(*job_list)

            self.freq_list = list(self.freq_list)
            self.freq_days_of_week_list = list(self.freq_days_of_week_list)
            self.freq_start_dt_list = list(self.freq_start_dt_list)
            self.freq_missed_run_list = list(self.freq_missed_run_list)
            self.freq_prev_run_list = list(self.freq_prev_run_list)
            self.freq_next_run_list = list(self.freq_next_run_list)
            self.sub_job_list = list(self.sub_job_list)
            self.sub_job_type_list = list(self.sub_job_type_list)

            for freq, start_time in zip(self.freq_list, self.freq_start_dt_list):
                self.js_list_box.insert('end', '{0} - {1}'.format(start_time.__format__("%Y%m%d %H:%M"),
                                                                  freq_name(freq)))

            self.js_list_sel = self.js_list_box.size() - 1
            self.js_list_box.select_set(self.js_list_box.size() - 1)

            for sub_job, sub_job_type in zip(self.sub_job_list, self.sub_job_type_list):
                self.jst_list_box.insert('end', '{0} <{1}>'.format(sub_job_type, os.path.basename(sub_job[0])))

            self.jst_list_sel = self.jst_list_box.size() - 1
            self.jst_list_box.select_set(self.jst_list_box.size() - 1)

        if not self.timeout_hh.get():
            self.timeout_hh.set(0)

        if not self.timeout_mm.get():
            self.timeout_mm.set(0)

    def job_task_type_toggle(self, event):
        self.job_attach_toggle(None)
        self.js_task.set('')
        self.js_fprefix.set('')

        if self.js_task_type.get() == 1:
            self.js_task_label.configure(text='Stored Proc:')
            self.js_fprefix_label.configure(text='File Prefix:')
            self.js_file_dir.grid_remove()
            self.js_attach_chkbox.grid(row=0, column=0, padx=3, pady=4, sticky='w')
            self.js_shellcomm_entry.configure(state=DISABLED)
        else:
            self.js_attach_chkbox.grid_remove()
            self.js_file_dir.grid(row=0, column=0, padx=3, pady=4, sticky='w')
            self.js_task_label.configure(text='Filepath:')
            self.js_fprefix_label.configure(text='Param(s):')
            self.js_shellcomm_entry.configure(state=NORMAL)

    def job_freq_toggle(self, event):
        self.js_monday.set(0)
        self.js_tuesday.set(0)
        self.js_wednesday.set(0)
        self.js_thursday.set(0)
        self.js_friday.set(0)
        self.js_saturday.set(0)
        self.js_sunday.set(0)

        if self.js_frequency.get() == 1 or self.js_frequency.get() == 2:
            self.js_monday_chkbox.grid(row=0, column=0, padx=2, pady=3, sticky='w')
            self.js_tuesday_chkbox.grid(row=1, column=0, padx=2, pady=3, sticky='w')
            self.js_wednesday_chkbox.grid(row=2, column=0, padx=2, pady=3, sticky='w')
            self.js_thursday_chkbox.grid(row=3, column=0, padx=2, pady=3, sticky='w')
            self.js_friday_chkbox.grid(row=4, column=0, padx=2, pady=3, sticky='w')
            self.js_saturday_chkbox.grid(row=5, column=0, padx=2, pady=3, sticky='w')
            self.js_sunday_chkbox.grid(row=6, column=0, padx=2, pady=3, sticky='w')
        else:
            self.js_monday_chkbox.grid_remove()
            self.js_tuesday_chkbox.grid_remove()
            self.js_wednesday_chkbox.grid_remove()
            self.js_thursday_chkbox.grid_remove()
            self.js_friday_chkbox.grid_remove()
            self.js_saturday_chkbox.grid_remove()
            self.js_sunday_chkbox.grid_remove()

    def job_attach_toggle(self, event):
        if self.js_attach.get() == 1 or self.js_task_type.get() == 0:
            self.js_fprefix_entry.configure(state=NORMAL)
        else:
            self.js_fprefix_entry.configure(state=DISABLED)

    # Function adjusts selection of item when user clicks item (JS List)
    def js_list_select(self, event):
        if self.js_list_box and self.js_list_box.curselection() \
                and -1 < self.js_list_sel < self.js_list_box.size() - 1:
            self.js_list_sel = self.js_list_box.curselection()[0]

    # Function adjusts selection of item when user presses down key JS List)
    def js_list_down(self, event):
        if self.js_list_sel < self.js_list_box.size() - 1:
            self.js_list_box.select_clear(self.js_list_sel)
            self.js_list_sel += 1
            self.js_list_box.select_set(self.js_list_sel)

    # Function adjusts selection of item when user presses up key (JS List)
    def js_list_up(self, event):
        if self.js_list_sel > 0:
            self.js_list_box.select_clear(self.js_list_sel)
            self.js_list_sel -= 1
            self.js_list_box.select_set(self.js_list_sel)

    # Function adjusts selection of item when user clicks item (JST List)
    def jst_list_select(self, event):
        if self.jst_list_box and self.jst_list_box.curselection() \
                and -1 < self.jst_list_sel < self.jst_list_box.size() - 1:
            self.jst_list_sel = self.jst_list_box.curselection()[0]

    # Function adjusts selection of item when user presses down key JST List)
    def jst_list_down(self, event):
        if self.jst_list_sel < self.jst_list_box.size() - 1:
            self.jst_list_box.select_clear(self.jst_list_sel)
            self.jst_list_sel += 1
            self.jst_list_box.select_set(self.jst_list_sel)

    # Function adjusts selection of item when user presses up key (JST List)
    def jst_list_up(self, event):
        if self.jst_list_sel > 0:
            self.jst_list_box.select_clear(self.jst_list_sel)
            self.jst_list_sel -= 1
            self.jst_list_box.select_set(self.jst_list_sel)

    def js_add(self):
        if self.js_date.get() and self.job_hh.get() and self.job_mm.get():
            days_of_week = []
            my_freq = self.js_frequency.get()
            start_time = datetime.datetime.combine(self.js_date_entry.get_date(), datetime
                                                   .time(hour=int(self.job_hh.get()),
                                                         minute=int(self.job_mm.get())))

            for freq, freq_start_dt in zip(self.freq_list, self.freq_start_dt_list):
                if freq == my_freq and freq_start_dt == start_time:
                    return

            days_of_week.append([0, self.js_monday.get()])
            days_of_week.append([1, self.js_tuesday.get()])
            days_of_week.append([2, self.js_wednesday.get()])
            days_of_week.append([3, self.js_thursday.get()])
            days_of_week.append([4, self.js_friday.get()])
            days_of_week.append([5, self.js_saturday.get()])
            days_of_week.append([6, self.js_sunday.get()])

            self.freq_days_of_week_list.append(days_of_week)
            self.freq_list.append(my_freq)
            self.freq_start_dt_list.append(start_time)
            self.freq_missed_run_list.append(self.js_missed_run.get())
            self.freq_prev_run_list.append(None)

            if start_time > datetime.datetime.now() and (my_freq == 0 or my_freq == 3):
                self.freq_next_run_list.append(start_time)
            else:
                self.freq_next_run_list.append(next_run_date(start_time, my_freq, days_of_week))

            self.js_frequency.set(0)
            self.js_date.set(datetime.datetime.today().__format__('%m/%d/%Y'))
            self.job_hh.set('')
            self.job_mm.set('')
            self.js_missed_run.set(0)
            self.job_freq_toggle(None)

            self.js_list_box.insert('end', '{0} - {1}'.format(start_time.__format__("%Y%m%d %H:%M"),
                                                              freq_name(my_freq)))

            if self.js_list_box.curselection():
                self.js_list_box.selection_clear(self.js_list_box.curselection())

            self.js_list_sel = self.js_list_box.size() - 1
            self.js_list_box.select_set(self.js_list_box.size() - 1)
            messagebox.showinfo('Schedule Added!', 'You have successfully added schedule {0} - {1}'
                                .format(start_time.__format__("%Y%m%d %H:%M"), freq_name(my_freq)),
                                parent=self.main)

    def js_mod(self):
        if self.js_list_box.curselection() and self.js_list_box.size() > 0:
            selection = self.js_list_box.curselection()[0]

            self.js_frequency.set(self.freq_list[selection])
            self.js_missed_run.set(self.freq_missed_run_list[selection])
            start_time = self.freq_start_dt_list[selection]
            self.js_date.set(start_time.__format__('%m/%d/%Y'))
            self.job_hh.set(start_time.hour)
            self.job_mm.set(start_time.minute)
            self.job_freq_toggle(None)

            for day_of_week in self.freq_days_of_week_list[selection]:
                if day_of_week[0] == 0:
                    self.js_monday.set(day_of_week[1])
                elif day_of_week[0] == 1:
                    self.js_tuesday.set(day_of_week[1])
                elif day_of_week[0] == 2:
                    self.js_wednesday.set(day_of_week[1])
                elif day_of_week[0] == 3:
                    self.js_thursday.set(day_of_week[1])
                elif day_of_week[0] == 4:
                    self.js_friday.set(day_of_week[1])
                elif day_of_week[0] == 5:
                    self.js_saturday.set(day_of_week[1])
                elif day_of_week[0] == 6:
                    self.js_sunday.set(day_of_week[1])

            self.js_del()

    def js_del(self):
        if self.js_list_box.curselection():
            del self.freq_list[self.js_list_box.curselection()[0]]
            del self.freq_days_of_week_list[self.js_list_box.curselection()[0]]
            del self.freq_start_dt_list[self.js_list_box.curselection()[0]]
            del self.freq_missed_run_list[self.js_list_box.curselection()[0]]
            del self.freq_prev_run_list[self.js_list_box.curselection()[0]]
            del self.freq_next_run_list[self.js_list_box.curselection()[0]]

            self.js_list_box.delete(self.js_list_box.curselection(), self.js_list_box.curselection())

            if self.js_list_box.size() > 0:
                self.js_list_sel -= 1
                self.js_list_box.select_set(self.js_list_sel)
            else:
                self.js_list_sel = 0

    def jst_add(self):
        sub_job_list = []
        i = -1
        sel = -1

        if self.js_task.get() and self.js_task_type.get() == 0:
            sub_job_list = [self.js_task.get(), self.js_fprefix.get(), self.js_shell_comm.get()]
        elif self.js_task.get() and self.js_task_type.get() == 1 and self.js_attach.get() == 0:
            sub_job_list = [self.js_task.get(), False, None]
        elif self.js_task.get() and self.js_task_type.get() == 1 and self.js_attach.get() == 1 \
                and self.js_fprefix.get():
            sub_job_list = [self.js_task.get(), True, self.js_fprefix.get()]

        if len(sub_job_list) > 0:
            if self.js_task_type.get() == 0:
                sj_type = 'Program'
                sj_name = os.path.basename(self.js_task.get())
            else:
                sj_type = 'Stored Procedure'
                sj_name = self.js_task.get()

            for sub_job, sub_job_type in zip(self.sub_job_list, self.sub_job_type_list):
                i += 1

                if sub_job_type == sj_type and sub_job[0].lower() == sub_job_list[0].lower()\
                        and sub_job[1] == sub_job_list[1] and (
                        sj_type == 'Program' or (sj_type == 'Stored Procedure' and sub_job[2] == sub_job_list[2])):
                    return
                elif sub_job[0].lower() == sub_job_list[0].lower():
                    sel = i
                    break

            self.js_task.set('')
            self.js_fprefix.set('')
            self.js_task_type.set(0)
            self.js_attach.set(0)

            self.job_task_type_toggle(None)

            if sel > -1:
                self.sub_job_type_list[sel] = sj_type
                self.sub_job_list[sel] = sub_job_list

                self.jst_list_box.delete(sel)
                self.jst_list_box.insert(sel, '{0} <{1}>'.format(sj_type, sj_name))

                if self.jst_list_box.curselection():
                    self.jst_list_box.selection_clear(self.jst_list_box.curselection())

                self.jst_list_sel = sel
                self.jst_list_box.select_set(sel)
                messagebox.showinfo('Sub-Job updated!', 'You have successfully updated Sub-Job %s' % sj_name,
                                    parent=self.main)
            else:
                self.sub_job_type_list.append(sj_type)
                self.sub_job_list.append(sub_job_list)

                self.jst_list_box.insert('end', '{0} <{1}>'.format(sj_type, sj_name))

                if self.jst_list_box.curselection():
                    self.jst_list_box.selection_clear(self.jst_list_box.curselection())

                self.jst_list_sel = self.jst_list_box.size() - 1
                self.jst_list_box.select_set(self.jst_list_box.size() - 1)
                messagebox.showinfo('Sub-Job Added!', 'You have successfully added Sub-Job %s' % sj_name,
                                    parent=self.main)

    def jst_mod(self):
        if self.jst_list_box.curselection() and self.jst_list_box.size() > 0:
            selection = self.jst_list_box.curselection()[0]
            sj_type = self.sub_job_type_list[selection]
            sub_job = self.sub_job_list[selection]

            if sj_type == 'Program':
                self.js_task_type.set(0)
                self.job_task_type_toggle(None)
                self.js_task.set(sub_job[0])
                self.js_fprefix.set(sub_job[1])
            else:
                self.js_task_type.set(1)
                self.job_task_type_toggle(None)
                self.js_task.set(sub_job[0])

                if sub_job[1]:
                    self.js_attach.set(1)
                    self.js_fprefix.set(sub_job[2])
                else:
                    self.js_attach.set(0)
                    self.js_fprefix.set('')

            self.jst_del()

    def jst_del(self):
        if self.jst_list_box.curselection():
            del self.sub_job_list[self.jst_list_box.curselection()[0]]
            del self.sub_job_type_list[self.jst_list_box.curselection()[0]]

            self.jst_list_box.delete(self.jst_list_box.curselection(), self.jst_list_box.curselection())

            if self.jst_list_box.size() > 0:
                self.jst_list_sel -= 1
                self.jst_list_box.select_set(self.jst_list_sel)
            else:
                self.jst_list_sel = 0

    def jst_up(self):
        selection = self.jst_list_box.curselection()

        if selection and selection[0] > 0:
            tmp = copy.deepcopy(self.sub_job_list[selection[0] - 1])
            tmp2 = copy.deepcopy(self.sub_job_type_list[selection[0] - 1])
            self.sub_job_list[selection[0] - 1] = self.sub_job_list[selection[0]]
            self.sub_job_type_list[selection[0] - 1] = self.sub_job_type_list[selection[0]]
            self.sub_job_list[selection[0]] = tmp
            self.sub_job_type_list[selection[0]] = tmp2
            tmp = self.jst_list_box.get(selection[0])
            self.jst_list_box.delete(selection[0], selection[0])
            self.jst_list_box.insert(selection[0] - 1, tmp)
            self.jst_list_sel -= 1
            self.jst_list_box.select_set(self.jst_list_sel)

    def jst_down(self):
        selection = self.jst_list_box.curselection()

        if selection and selection[0] < self.jst_list_box.size() - 1:
            tmp = copy.deepcopy(self.sub_job_list[selection[0] + 1])
            tmp2 = copy.deepcopy(self.sub_job_type_list[selection[0] + 1])
            self.sub_job_list[selection[0] + 1] = self.sub_job_list[selection[0]]
            self.sub_job_type_list[selection[0] + 1] = self.sub_job_type_list[selection[0]]
            self.sub_job_list[selection[0]] = tmp
            self.sub_job_type_list[selection[0]] = tmp2
            tmp = self.jst_list_box.get(selection[0])
            self.jst_list_box.delete(selection[0], selection[0])
            self.jst_list_box.insert(selection[0] + 1, tmp)
            self.jst_list_sel += 1
            self.jst_list_box.select_set(self.jst_list_sel)

    # Function to search Directory
    def jst_dir(self):
        file_name = None

        if self.js_task.get() and os.path.exists(self.js_task.get()):
            init_dir = os.path.dirname(self.js_task.get())

            if init_dir != self.js_task.get():
                file_name = os.path.basename(self.js_task.get())
        else:
            init_dir = '/'

        file = filedialog.askopenfilename(initialdir=init_dir, title='Select Executable File',
                                          initialfile=file_name, parent=self.main)

        if file:
            self.js_task.set(file)

    # Function to Add/Modify a Job when button is pressed
    def add_job(self):
        if not self.job_name.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Job Name',
                                 parent=self.main)
        elif not self.email_to.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Email To',
                                 parent=self.main)
        elif len(self.freq_list) < 1:
            messagebox.showerror('Field Empty Error!', 'You haven''t added a Job Schedule. Plz add one',
                                 parent=self.main)
        elif len(self.sub_job_list) < 1:
            messagebox.showerror('Field Empty Error!', 'You haven''t added a Job Task. Plz add one',
                                 parent=self.main)
        elif not self.timeout_hh.get() > 0 and not self.timeout_mm.get() > 0:
            messagebox.showerror('Timeout Fields Error!', 'You haven''t added timeout time. Plz add one',
                                 parent=self.main)
        else:
            i = -1
            line_num = i
            new_config = dict()
            configs = global_objs['Local_Settings'].grab_item('Job_Configs')

            if configs:
                for config in configs:
                    i += 1

                    if self.job and config['Job_Name'] == self.job['Job_Name']:
                        line_num = i
                        break
                    elif not self.job and config['Job_Name'].lower() == self.job_name.get().lower():
                        messagebox.showerror('Field Error!',
                                             'Job Name already exist as a Job. Plz choose different Name',
                                             parent=self.main)
                        return

                if self.job and line_num < 0:
                    messagebox.showerror('Job Not Found Error!', 'Unable to find Job in Job List', parent=self.main)
                    return
            else:
                configs = []

            new_config['Job_Name'] = self.job_name.get()
            new_config['Job_Controls'] = [True, False, 0]
            new_config['To_Distro'] = self.email_to.get()
            new_config['CC_Distro'] = self.email_cc.get()
            new_config['Job_Timeout'] = [self.timeout_hh.get(), self.timeout_mm.get()]
            new_config['Job_Schedule'] = zip(self.freq_list, self.freq_days_of_week_list,
                                             self.freq_start_dt_list, self.freq_missed_run_list,
                                             self.freq_prev_run_list, self.freq_next_run_list)
            new_config['Job_List'] = zip(self.sub_job_list, self.sub_job_type_list)

            if self.job:
                configs[line_num] = new_config
            else:
                configs.append(new_config)

            add_setting('Local_Settings', configs, 'Job_Configs', False)

            if self.job and self.job_name.get().lower() != self.job['Job_Name'].lower():
                files = list(pl.Path(joblogsdir).glob('%s*' % self.job['Job_Name']))

                for file in files:
                    ext = os.path.splitext(file)[1]

                    if ext in ['.bak', '.dat', '.dir']:
                        try:
                            os.rename(file, os.path.join(os.path.dirname(file), '{0}.{1}'.format(self.job_name.get(),
                                                                                                 ext)))
                        except:
                            pass

            self.destroy = True
            self.class_obj.fill_gui(True)
            self.main.destroy()

    def del_job(self):
        myresponse = messagebox.askokcancel(
            'Delete Notice!',
            'Deleting this job will lose this job forever. Would you like to proceed?',
            parent=self.main)

        if myresponse:
            configs = global_objs['Local_Settings'].grab_item('Job_Configs')

            if configs:
                for config in configs:
                    if config['Job_Name'] == self.job['Job_Name']:
                        configs.remove(config)

                        if len(configs) > 0:
                            add_setting('Local_Settings', configs, 'Job_Configs', False)
                        else:
                            add_setting('Local_Settings', None, 'Job_Configs', False)

                        files = list(pl.Path(joblogsdir).glob('%s*' % self.job['Job_Name']))

                        for file in files:
                            ext = os.path.splitext(file)[1]

                            if ext in ['.bak', '.dat', '.dir']:
                                try:
                                    os.remove(file)
                                except:
                                    pass

                        self.destroy = True
                        self.class_obj.fill_gui(True)
                        self.main.destroy()
                        return

            messagebox.showerror('Job Not Found Error!', 'Unable to find Job in Job List', parent=self.main)

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        self.main.destroy()


class JobListGUI:
    job_list_box = None
    job_obj = None
    job_log_obj = None
    job_log = None
    job_log_button = None
    mod_job_button = None
    job_action_button = None
    job_instant_action_button = None
    job_list_sel = 0

    def __init__(self, root):
        self.main = Toplevel(root)
        self.configs = global_objs['Local_Settings'].grab_item('Job_Configs')
        self.main.iconbitmap(icon_path)
        self.config_selected = None

        self.prev_run = StringVar()
        self.next_run = StringVar()

    def build_gui(self):
        # Set GUI Geometry and GUI Title
        self.main.geometry('357x302+630+290')
        self.main.title('Jobs List')
        self.main.resizable(False, False)

        # Create Frames for GUI
        header_frame = Frame(self.main)
        job_list_frame = LabelFrame(self.main, text='Jobs', width=245, height=240)
        job_list_lframe = Frame(job_list_frame, width=122, height=240)
        job_list_rframe = Frame(job_list_frame, width=122, height=240)
        buttons_frame = Frame(self.main)

        # Apply Frames to GUI
        header_frame.pack(fill='both')
        job_list_frame.pack(fill='both')
        job_list_lframe.grid(row=0, column=0, sticky=W)
        job_list_rframe.grid(row=0, column=1, sticky=W + N + S)
        buttons_frame.pack(fill='both')

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text='Please choose a job to take action on', width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Widgets to Job List Left Frame
        #     Job List List Box
        job_yscrollbar = Scrollbar(job_list_lframe, orient="vertical")
        job_xscrollbar = Scrollbar(job_list_lframe, orient="horizontal")
        self.job_list_box = Listbox(job_list_lframe, selectmode=SINGLE, width=25, height=12,
                                    yscrollcommand=job_yscrollbar.set, xscrollcommand=job_xscrollbar.set)
        job_yscrollbar.config(command=self.job_list_box.yview)
        job_xscrollbar.config(command=self.job_list_box.xview)
        self.job_list_box.grid(row=0, column=0, rowspan=4, padx=8, pady=5)
        job_yscrollbar.grid(row=0, column=1, rowspan=4, sticky=N + S)
        job_xscrollbar.grid(row=4, column=0, columnspan=2, sticky=E + W)
        self.job_list_box.bind("<Down>", self.job_list_down)
        self.job_list_box.bind("<Up>", self.job_list_up)
        self.job_list_box.bind('<<ListboxSelect>>', self.job_list_select)

        # Apply Widgets to Job List Left Frame
        #     Prev Run Time Input Box
        prev_run_label = Label(job_list_rframe, text='Prev Run:')
        prev_run_txtbox = Entry(job_list_rframe, textvariable=self.prev_run, width=15, state=DISABLED)
        prev_run_label.grid(row=0, column=0, padx=2, pady=5, sticky='e')
        prev_run_txtbox.grid(row=0, column=1, padx=3, pady=5, sticky='e')

        #     Next Run Time Input Box
        next_run_label = Label(job_list_rframe, text='Next Run:')
        next_run_txtbox = Entry(job_list_rframe, textvariable=self.next_run, width=15, state=DISABLED)
        next_run_label.grid(row=1, column=0, padx=2, pady=5, sticky='e')
        next_run_txtbox.grid(row=1, column=1, padx=3, pady=5, sticky='w')

        #     Modify Job Button
        self.mod_job_button = Button(job_list_rframe, text='Modify Job', width=21, command=self.modify_job)
        self.mod_job_button.grid(row=2, column=0, columnspan=2, pady=6, padx=2, sticky='w')

        self.job_log_button = Button(job_list_rframe, text='View Job Log', width=21, command=self.view_job_log)
        self.job_log_button.grid(row=3, column=0, columnspan=2, pady=4, padx=2, sticky='w')

        self.job_action_button = Button(job_list_rframe, text='Disable Job', width=21, command=self.job_action)
        self.job_action_button.grid(row=4, column=0, columnspan=2, pady=4, padx=2, sticky='w')

        self.job_instant_action_button = Button(job_list_rframe, text='Start Job', width=21,
                                                command=self.job_instant_action)
        self.job_instant_action_button.grid(row=5, column=0, columnspan=2, pady=4, padx=2, sticky='w')

        # Apply Buttons to the Buttons Frame
        #     Add Job
        add_job_button = Button(buttons_frame, text='Add New Job', width=15, command=self.add_job)
        add_job_button.grid(row=0, column=0, pady=6, padx=15, sticky='w')

        #     Cancel Button
        cancel_button = Button(buttons_frame, text='Cancel', width=15, command=self.cancel)
        cancel_button.grid(row=0, column=1, pady=6, padx=85, sticky='e')

        self.fill_gui()

    def fill_gui(self, select_prev_item=False):
        if not select_prev_item:
            global_objs['Local_Settings'].read_shelf()

        self.configs = global_objs['Local_Settings'].grab_item('Job_Configs')

        if self.job_list_box.size() > 0:
            self.job_list_box.delete(0, self.job_list_box.size() - 1)

        if self.configs:
            for config in self.configs:
                self.job_list_box.insert('end', config['Job_Name'])

            if select_prev_item:
                self.job_list_box.select_clear(self.job_list_sel)
                self.job_list_box.activate(self.job_list_sel)
                self.job_list_box.select_set(self.job_list_sel)
            else:
                self.job_list_box.select_set(0)

            self.list_toggle()
            self.mod_job_button.configure(state=NORMAL)
            self.job_action_button.configure(state=NORMAL)
            self.job_instant_action_button.configure(state=NORMAL)
        else:
            self.job_log_button.configure(state=DISABLED)
            self.mod_job_button.configure(state=DISABLED)
            self.job_action_button.configure(state=DISABLED)
            self.job_instant_action_button.configure(state=DISABLED)

    # Function adjusts selection of item when user clicks item (Job List)
    def job_list_select(self, event):
        if self.job_list_box.size() > 0 and self.job_list_box.curselection():
            self.job_list_sel = self.job_list_box.curselection()[0]

        self.list_toggle()

    # Function adjusts selection of item when user presses down key (Job List)
    def job_list_down(self, event):
        if self.job_list_sel < self.job_list_box.size() - 1:
            self.job_list_box.select_clear(self.job_list_sel)
            self.job_list_sel += 1
            self.job_list_box.select_set(self.job_list_sel)
            # self.job_list_box.activate(self.job_list_sel)

        self.list_toggle()

    # Function adjusts selection of item when user presses up key (Job List)
    def job_list_up(self, event):
        if self.job_list_sel > 0:
            self.job_list_box.select_clear(self.job_list_sel)
            self.job_list_sel -= 1
            self.job_list_box.select_set(self.job_list_sel)
            # self.job_list_box.activate(self.job_list_sel)

        self.list_toggle()

    # Function to show Add Job GUI
    def add_job(self):
        if self.job_obj:
            self.job_obj.cancel()

        self.job_obj = JobGUI(self, self.main)
        self.job_obj.build_gui()

    def list_toggle(self):
        if self.job_list_box.size() > 0 and self.job_list_sel > -1:
            config_found = None
            job_name = self.job_list_box.get(self.job_list_sel)

            if not self.config_selected or self.config_selected != job_name:
                self.config_selected = job_name

                if os.path.exists(os.path.join(joblogsdir, '%s.bak' % job_name)):
                    self.job_log = ShelfHandle(os.path.join(joblogsdir, job_name))
                    self.job_log.read_shelf()

                    if self.job_log and self.job_log.grab_list():
                        self.job_log_button.configure(state=NORMAL)
                    else:
                        self.job_log_button.configure(state=DISABLED)
                else:
                    self.job_log_button.configure(state=DISABLED)

            for config in self.configs:
                if config['Job_Name'].lower() == job_name.lower():
                    config_found = config

            if config_found:
                if config_found['Job_Controls'][0]:
                    action_name = 'Disable'
                else:
                    action_name = 'Enable'

                if config_found['Job_Controls'][1]:
                    if config_found['Job_Controls'][2] == 1:
                        instant_action_name = 'Stopping'
                    else:
                        instant_action_name = 'Stop'
                else:
                    if config_found['Job_Controls'][2] == 2:
                        instant_action_name = 'Starting'
                    else:
                        instant_action_name = 'Start'

                self.job_action_button.configure(text='%s Job' % action_name)
                self.job_instant_action_button.configure(text='%s Job' % instant_action_name)

                max_prev_run = None
                min_next_run = None
                job_schedule = copy.deepcopy(config_found['Job_Schedule'])

                for freq, days_of_week, start_dt, missed_run, prev_run, next_run in job_schedule:
                    if not max_prev_run or (max_prev_run and prev_run and max_prev_run < prev_run):
                        max_prev_run = prev_run

                    if not min_next_run or (min_next_run and next_run and min_next_run > next_run):
                        min_next_run = next_run

                if max_prev_run:
                    self.prev_run.set(max_prev_run.__format__("%Y%m%d %H:%M"))
                else:
                    self.prev_run.set('')

                if min_next_run:
                    self.next_run.set(min_next_run.__format__("%Y%m%d %H:%M"))
                else:
                    self.next_run.set('')

                return

        self.job_log_button.configure(state=DISABLED)
        self.prev_run.set('')
        self.next_run.set('')

    def modify_job(self):
        if self.job_list_box.curselection():
            config_found = None
            global_objs['Local_Settings'].read_shelf()
            self.configs = global_objs['Local_Settings'].grab_item('Job_Configs')
            job_name = self.job_list_box.get(self.job_list_box.curselection())

            for config in self.configs:
                if config['Job_Name'].lower() == job_name.lower():
                    config_found = config

            if config_found:
                if self.job_obj:
                    self.job_obj.cancel()

                self.job_obj = JobGUI(self, self.main, config_found)
                self.job_obj.build_gui()
            else:
                messagebox.showerror('Job Not Found Error!', 'Unable to find Job in Job List', parent=self.main)

    def view_job_log(self):
        if self.job_list_box.curselection() and self.job_log and self.job_log.grab_list():
            if self.job_log_obj:
                self.job_log_obj.cancel()

            job_name = self.job_list_box.get(self.job_list_box.curselection())
            self.job_log_obj = JobLogGUI(self, self.main, job_name, self.job_log)
            self.job_log_obj.build_gui()

    def job_action(self):
        if self.job_list_box.curselection():
            sel = None
            config_found = None
            job_name = self.job_list_box.get(self.job_list_box.curselection())
            global_objs['Local_Settings'].read_shelf()
            self.configs = global_objs['Local_Settings'].grab_item('Job_Configs')

            for line, config in enumerate(self.configs):
                if config['Job_Name'].lower() == job_name.lower():
                    sel = line
                    config_found = config
                    break

            if config_found:
                job_status = config_found['Job_Controls'][1]
                job_action_overide = 0

                if config_found['Job_Controls'][0]:
                    job_enabled = False
                    button_name = 'Enable'
                    action_name = 'Disabled'
                    self.job_instant_action_button.configure(state=DISABLED)
                else:
                    job_enabled = True
                    button_name = 'Disable'
                    action_name = 'Enabled'
                    self.job_instant_action_button.configure(state=NORMAL)
                    freq_list, freq_days_of_week_list, freq_start_dt_list, freq_missed_run_list,\
                    freq_prev_run_list, freq_next_run_list = zip(*copy.deepcopy(config_found['Job_Schedule']))
                    freq_list = list(freq_list)
                    freq_days_of_week_list = list(freq_days_of_week_list)
                    freq_start_dt_list = list(freq_start_dt_list)
                    freq_missed_run_list = list(freq_missed_run_list)
                    freq_prev_run_list = list(freq_prev_run_list)
                    freq_next_run_list = list(freq_next_run_list)

                    for line in range(len(freq_list)):
                        freq_next_run_list[line] = next_run_date(freq_start_dt_list[line], freq_list[line],
                                                                 freq_days_of_week_list[line])
                    config_found['Job_Schedule'] = zip(freq_list, freq_days_of_week_list, freq_start_dt_list,
                                                       freq_missed_run_list, freq_prev_run_list, freq_next_run_list)

                if job_status:
                    myresponse = messagebox.askokcancel('Job Running Notice!',
                                                        "Job '{0}' is currently running. Would you like to stop job?"
                                                        .format(config_found['Job_Name']), parent=self.main)

                    if not myresponse:
                        job_action_overide = 2

                config_found['Job_Controls'] = [job_enabled, job_status, job_action_overide]
                self.configs[sel] = config_found
                add_setting('Local_Settings', self.configs, 'Job_Configs', False)
                self.job_action_button.configure(text='%s Job' % button_name)

                messagebox.showinfo('Job {0}!', "Job '{1}' has been {0}"
                                    .format(action_name, config_found['Job_Name']), parent=self.main)

    def job_instant_action(self):
        if self.job_list_box.curselection():
            sel = None
            config_found = None
            global_objs['Local_Settings'].read_shelf()
            self.configs = global_objs['Local_Settings'].grab_item('Job_Configs')
            job_name = self.job_list_box.get(self.job_list_box.curselection())

            for line, config in enumerate(self.configs):
                if config['Job_Name'].lower() == job_name.lower():
                    sel = line
                    config_found = config
                    break

            if config_found and config_found['Job_Controls'][0]:
                job_status = config_found['Job_Controls'][1]
                job_instant_action = config_found['Job_Controls'][2]

                if job_status or job_instant_action == 2:
                    if job_status:
                        job_instant_action = 1
                    else:
                        job_instant_action = 0

                    button_text = 'Start Job'
                    action_text = 'Stopped'
                else:
                    job_instant_action = 2
                    button_text = 'Stop Job'
                    action_text = 'Started'

                config_found['Job_Controls'] = [True, job_status, job_instant_action]
                self.configs[sel] = config_found
                add_setting('Local_Settings', self.configs, 'Job_Configs', False)
                self.job_instant_action_button.configure(text=button_text)

                messagebox.showinfo('Job {0}!', "Job '{1}' has been {0}"
                                    .format(action_text, config_found['Job_Name']), parent=self.main)

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        self.main.destroy()


class JobLogGUI:
    dates_list_box = None
    history_list_box = None
    dates_list_sel = 0
    history_list_sel = 0

    def __init__(self, class_obj, root, job_name, job_log):
        self.root = root
        self.main = Toplevel(root)
        self.job_log = job_log
        self.job_name = job_name
        self.class_obj = class_obj
        self.destroy = False
        self.main.iconbitmap(icon_path)

        self.main.bind('<Destroy>', self.gui_cleanup)

    def gui_cleanup(self, event):
        if self.root and not self.destroy:
            self.destroy = True
            self.class_obj.fill_gui(True)

    def build_gui(self):
        # Set GUI Geometry and GUI Title
        self.main.geometry('484x247+630+290')
        self.main.title("'{0}' Log History".format(self.job_name))
        self.main.resizable(False, False)

        # Create Frames for GUI
        log_list_frame = Frame(self.main, width=245, height=240)
        log_list_lframe = LabelFrame(log_list_frame, text='Log Dates', width=122, height=240)
        log_list_rframe = LabelFrame(log_list_frame, text='Log History', width=122, height=240)
        buttons_frame = Frame(self.main)

        # Apply Frames to GUI
        log_list_frame.pack(fill='both')
        log_list_lframe.grid(row=0, column=0, sticky=W)
        log_list_rframe.grid(row=0, column=1, sticky=W + N + S)
        buttons_frame.pack(fill='both')

        # Apply Widgets to Log List Left Frame
        #     Log Dates List Box
        log_yscrollbar = Scrollbar(log_list_lframe, orient="vertical")
        log_xscrollbar = Scrollbar(log_list_lframe, orient="horizontal")
        self.dates_list_box = Listbox(log_list_lframe, selectmode=SINGLE, width=15, yscrollcommand=log_yscrollbar.set,
                                      xscrollcommand=log_xscrollbar.set)
        log_yscrollbar.config(command=self.dates_list_box.yview)
        log_xscrollbar.config(command=self.dates_list_box.xview)
        self.dates_list_box.grid(row=0, column=0, rowspan=4, padx=8, pady=5)
        log_yscrollbar.grid(row=0, column=1, rowspan=4, sticky=N + S)
        log_xscrollbar.grid(row=4, column=0, columnspan=2, sticky=E + W)
        self.dates_list_box.bind("<Down>", self.dates_list_down)
        self.dates_list_box.bind("<Up>", self.dates_list_up)
        self.dates_list_box.bind('<<ListboxSelect>>', self.dates_list_select)

        # Apply Widgets to Log List Left Frame
        #     Log Dates List Box
        log_yscrollbar = Scrollbar(log_list_rframe, orient="vertical")
        log_xscrollbar = Scrollbar(log_list_rframe, orient="horizontal")
        self.history_list_box = Listbox(log_list_rframe, selectmode=SINGLE, width=52, yscrollcommand=log_yscrollbar.set,
                                        xscrollcommand=log_xscrollbar.set)
        log_yscrollbar.config(command=self.history_list_box.yview)
        log_xscrollbar.config(command=self.history_list_box.xview)
        self.history_list_box.grid(row=0, column=0, rowspan=4, padx=8, pady=5)
        log_yscrollbar.grid(row=0, column=1, rowspan=4, sticky=N + S)
        log_xscrollbar.grid(row=4, column=0, columnspan=2, sticky=E + W)
        self.history_list_box.bind("<Down>", self.history_list_down)
        self.history_list_box.bind("<Up>", self.history_list_up)
        self.history_list_box.bind('<<ListboxSelect>>', self.history_list_select)

        #     Export Log Button
        export_button = Button(buttons_frame, text='Export Log', width=13, command=self.export_to_file)
        export_button.grid(row=0, column=0, pady=6, padx=16, sticky='e')

        #     Copy To Clipboard Button
        copy_button = Button(buttons_frame, text='Copy Log', width=13, command=self.copy_to_clipboard)
        copy_button.grid(row=0, column=1, pady=6, padx=8, sticky='e')

        #     Copy To Clipboard Button
        delete_button = Button(buttons_frame, text='Delete Log', width=13, command=self.delete_log)
        delete_button.grid(row=0, column=2, pady=6, padx=8, sticky='e')

        #     Cancel Button
        cancel_button = Button(buttons_frame, text='Cancel', width=13, command=self.cancel)
        cancel_button.grid(row=0, column=3, pady=6, padx=8, sticky='e')

        self.fill_gui()

    def fill_gui(self):
        for key in self.job_log.get_keys():
            self.dates_list_box.insert(0, key)

        self.dates_list_box.select_set(0)
        self.select_toggle()

    # Function adjusts selection of item when user clicks item (Dates List)
    def dates_list_select(self, event):
        if self.dates_list_box.size() > 0 and self.dates_list_box.curselection():
            if self.dates_list_box.curselection()[0] != self.dates_list_sel:
                self.dates_list_sel = self.dates_list_box.curselection()[0]
                self.select_toggle()
        elif self.dates_list_box.size() < 1:
            self.select_toggle()
            # self.dates_list_box.select_set(self.dates_list_sel)

    # Function adjusts selection of item when user presses down key (Dates List)
    def dates_list_down(self, event):
        if self.dates_list_sel < self.dates_list_box.size() - 1:
            self.dates_list_box.select_clear(self.dates_list_sel)
            self.dates_list_sel += 1
            self.dates_list_box.select_set(self.dates_list_sel)

        self.select_toggle()

    # Function adjusts selection of item when user presses up key (Dates List)
    def dates_list_up(self, event):
        if self.dates_list_sel > 0:
            self.dates_list_box.select_clear(self.dates_list_sel)
            self.dates_list_sel -= 1
            self.dates_list_box.select_set(self.dates_list_sel)

        self.select_toggle()

    # Function adjusts selection of item when user clicks item (Dates List)
    def history_list_select(self, event):
        if self.history_list_box.size() > 0 and self.history_list_box.curselection():
            self.history_list_sel = self.history_list_box.curselection()[0]
            self.history_list_box.select_set(self.history_list_sel)

    # Function adjusts selection of item when user presses down key (Dates List)
    def history_list_down(self, event):
        if self.history_list_sel < self.history_list_box.size() - 1:
            self.history_list_box.select_clear(self.history_list_sel)
            self.history_list_sel += 1
            self.history_list_box.select_set(self.history_list_sel)

    # Function adjusts selection of item when user presses up key (Dates List)
    def history_list_up(self, event):
        if self.history_list_sel > 0:
            self.history_list_box.select_clear(self.history_list_sel)
            self.history_list_sel -= 1
            self.history_list_box.select_set(self.history_list_sel)

    def select_toggle(self):
        if self.history_list_box.size() > 0:
            self.history_list_box.delete(0, self.history_list_box.size() - 1)

        if self.dates_list_box.size() > 0 and self.dates_list_sel > -1:
            for val in self.job_log.grab_item(self.dates_list_box.get(self.dates_list_sel)):
                self.history_list_box.insert(0, ' - '.join(val))

    # Function to Copy Log into Clipboard when button is pressed
    def copy_to_clipboard(self):
        if self.history_list_box.size() > 0:
            text = '\n'.join(self.history_list_box.get(0, self.history_list_box.size() - 1))
            pyperclip.copy(text)
            messagebox.showinfo('Log Copied!', 'Your log has been copied to clipboard', parent=self.main)

    # Function to Write Log into Flat-File when button is pressed
    def export_to_file(self):
        if self.history_list_box.size() > 0:
            hist_list = self.history_list_box.get(0, self.history_list_box.size() - 1)
            hist_date = str(self.dates_list_box.get(self.dates_list_sel))
            export_file = os.path.join(joblogsexpdir, '{0}_{1}.txt'.format(hist_date, self.job_name))

            with portalocker.Lock(export_file, 'w') as f:
                f.write('\n'.join(hist_list))

            messagebox.showinfo('Log Exported!', 'Your log has been exported to %s' % os.path.basename(export_file),
                                parent=self.main)

    def delete_log(self):
        hist_date = str(self.dates_list_box.get(self.dates_list_sel))
        self.job_log.del_item(hist_date)
        self.job_log.write_shelf()
        self.dates_list_box.delete(self.dates_list_sel)

        if self.dates_list_box.size() > 0:
            if self.dates_list_sel > 0:
                self.dates_list_sel -= 1

            self.dates_list_box.select_set(self.dates_list_sel)

        global_objs['Local_Settings'].read_shelf()
        self.destroy = True
        self.class_obj.fill_gui(True)

        self.select_toggle()

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        self.main.destroy()


# Function to fill textbox in GUI
def fill_textbox(setting_list, val, key):
    assert (key and val and setting_list)
    item = global_objs[setting_list].grab_item(key)

    if isinstance(item, CryptHandle):
        val.set(item.decrypt_text())


# Function to add setting to Local_Settings shelf files
def add_setting(setting_list, val, key, encrypt=True):
    assert (key and setting_list)

    global_objs[setting_list].del_item(key)

    if val:
        global_objs[setting_list].add_item(key=key, val=val, encrypt=encrypt)

    global_objs[setting_list].write_shelf()
    global_objs[setting_list].backup()


def next_date(start_date, days):
    date_part = datetime.datetime.now() - start_date
    days_part = days + (floor(date_part.days / days) * days)

    if days_part < 0:
        days_part = 0

    return days_part


def next_run_date_list(start_date, days, days_of_week):
    next_dates = []

    for args in days_of_week:
        if args[1] == 1:
            if args[0] < start_date.weekday():
                start_date2 = start_date + datetime.timedelta(days=7 - (start_date.weekday() - args[0]))
            elif args[0] > start_date.weekday():
                start_date2 = start_date + datetime.timedelta(days=args[0] - start_date.weekday())
            else:
                start_date2 = start_date

            ndate = start_date2 + datetime.timedelta(days=next_date(start_date2, days))

            if ndate > datetime.datetime.now():
                next_dates.append(ndate)

    if len(next_dates) > 0:
        next_dates.sort()
        return next_dates[0]


def next_run_date(start_date, freq_type, days_of_week=None):
    if freq_type == 0:
        return start_date + datetime.timedelta(days=next_date(start_date, 1))
    elif freq_type == 1:
        return next_run_date_list(start_date, 7, days_of_week)
    elif freq_type == 2:
        return next_run_date_list(start_date, 14, days_of_week)
    elif freq_type == 3:
        now = datetime.datetime.now()

        if start_date > now:
            return start_date
        else:
            y = now.year - start_date.year
            m = now.month - start_date.month

            return start_date + relativedelta.relativedelta(months=1 + (y * 12 + m))


def freq_name(freq_type):
    if freq_type == 0:
        return 'Daily'
    elif freq_type == 1:
        return 'Weekly'
    elif freq_type == 2:
        return 'Bi-Weekly'
    elif freq_type == 3:
        return 'Monthly'


# Main loop routine to create GUI Settings
if __name__ == '__main__':
    if not os.path.exists(joblogsdir):
        os.makedirs(joblogsdir)

    if not os.path.exists(joblogsexpdir):
        os.makedirs(joblogsexpdir)

    obj = SettingsGUI()
    obj.build_gui()
