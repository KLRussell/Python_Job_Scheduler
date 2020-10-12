from KGlobal import Toolbox
from tkinter.messagebox import showerror, askokcancel, showinfo
from tkinter import *
from threading import Thread, Lock
from tkcalendar import DateEntry
from os.path import exists, dirname, basename, join
from copy import deepcopy
from babel import numbers
from gc import collect

import sys

if getattr(sys, 'frozen', False):
    app_path = sys.executable
    ico_dir = sys._MEIPASS
else:
    app_path = __file__
    ico_dir = dirname(__file__)

tool = Toolbox(app_path, logging_folder="01_Event_Logs", logging_base_name="Job_Scheduler", max_pool_size=100)
attach_dir = join(tool.local_config_dir, "02_Attachments")
temp_dir = join(tool.local_config_dir, '08_Temp_Files')
job_logs_dir = join(tool.local_config_dir, "05_Job_Logs")
job_logs_export_dir = join(tool.local_config_dir, "06_Job_Logs_Export")
icon_path = join(ico_dir, 'New_Job_Scheduler.ico')

email_engine = tool.default_exchange_conn(auto_renew=True)
local_config = tool.local_config


class JobLog(Toplevel):
    __export_button = None
    __copy_button = None
    __delete_button = None
    __date_list = None
    __history_list = None
    __log_date = None

    """
        Class to pull job logs by date in a GUI window by using Toplevel in tkinter
    """

    def __init__(self, parent, job_name):
        Toplevel.__init__(self)
        from KGlobal.data import DataConfig
        self.__parent = parent
        self.__job_name = job_name
        self.__job_log = DataConfig(file_dir=job_logs_dir, file_name_prefix=job_name.lower(), file_ext='log')
        self.iconbitmap(icon_path)
        self.bind('<Destroy>', self.__fix_focus)
        self.__build()
        self.__load_gui()

    def __fix_focus(self, event):
        if event and hasattr(event, 'widget') and event.widget == self:
            self.__parent.load_gui(self.__job_name)

    def __build(self):
        # Set GUI Geometry and GUI Title
        self.geometry('484x247+630+290')
        self.title("'%s' Log History" % self.__job_name)
        self.resizable(False, False)

        # Set GUI Frames
        main_frame = Frame(self)
        date_frame = LabelFrame(main_frame, text='Dates')
        history_frame = LabelFrame(main_frame, text='History')
        buttons_frame = Frame(self)

        # Apply Frames into GUI
        main_frame.pack(fill="both")
        date_frame.grid(row=0, column=0, sticky=W)
        history_frame.grid(row=0, column=1, sticky=W + N + S)
        buttons_frame.pack(fill="both")

        # Apply Widgets to date_frame
        #     Log Date List
        xbar = Scrollbar(date_frame, orient='horizontal')
        ybar = Scrollbar(date_frame, orient='vertical')
        self.__date_list = Listbox(date_frame, selectmode=SINGLE, width=15, yscrollcommand=ybar,
                                   xscrollcommand=xbar)
        xbar.config(command=self.__date_list.xview)
        ybar.config(command=self.__date_list.yview)
        self.__date_list.grid(row=0, column=0, padx=8, pady=5)
        xbar.grid(row=1, column=0, sticky=W + E)
        ybar.grid(row=0, column=1, sticky=N + S)
        self.__date_list.bind("<Down>", self.__list_action)
        self.__date_list.bind("<Up>", self.__list_action)
        self.__date_list.bind('<<ListboxSelect>>', self.__list_action)

        #     Log History List
        xbar = Scrollbar(history_frame, orient='horizontal')
        ybar = Scrollbar(history_frame, orient='vertical')
        self.__history_list = Listbox(history_frame, selectmode=SINGLE, width=52, yscrollcommand=ybar,
                                      xscrollcommand=xbar)
        xbar.config(command=self.__history_list.xview)
        ybar.config(command=self.__history_list.yview)
        self.__history_list.grid(row=0, column=0, padx=8, pady=5)
        xbar.grid(row=1, column=0, sticky=W + E)
        ybar.grid(row=0, column=1, sticky=N + S)
        self.__history_list.bind("<Down>", self.__list_action)
        self.__history_list.bind("<Up>", self.__list_action)

        # Apply Widgets to buttons_frame
        #     Export Log Button
        self.__export_button = Button(buttons_frame, text='Export Date', width=13, command=self.__export_date)
        self.__export_button.grid(row=0, column=0, pady=6, padx=16, sticky='e')

        #     Copy To Clipboard Button
        self.__copy_button = Button(buttons_frame, text='Copy Date', width=13, command=self.__copy_date)
        self.__copy_button.grid(row=0, column=1, pady=6, padx=8, sticky='e')

        #     Copy To Clipboard Button
        self.__delete_button = Button(buttons_frame, text='Delete Date', width=13, command=self.__delete_date)
        self.__delete_button.grid(row=0, column=2, pady=6, padx=8, sticky='e')

        #     Cancel Button
        button = Button(buttons_frame, text='Exit', width=13, command=self.destroy)
        button.grid(row=0, column=3, pady=6, padx=8, sticky='e')

    def __load_gui(self):
        if self.__date_list.size() > 0:
            self.__date_list.delete(0, self.__date_list.size() - 1)

        if self.__job_log:
            for log_date in self.__job_log.keys():
                self.__date_list.insert('end', log_date)

            self.after_idle(self.__date_list.select_set, 0)

        self.after_idle(self.__load_history)

    def __load_history(self):
        state = DISABLED
        selections = self.__date_list.curselection()

        if (selections or self.__date_list.size() < 1) and self.__history_list.size() > 0:
            self.__log_date = None
            self.__history_list.delete(0, self.__history_list.size() - 1)

        if selections:
            state = NORMAL
            self.__log_date = self.__date_list.get(selections[0])

            if self.__job_log[self.__log_date]:
                for log_entry in self.__job_log[self.__log_date]:
                    self.__history_list.insert('end', ' - '.join(log_entry))

        self.__copy_button.configure(state=state)
        self.__export_button.configure(state=state)
        self.__delete_button.configure(state=state)

    def __list_action(self, event):
        widget = event.widget

        if widget.size() > 0:
            selections = widget.curselection()

            if selections and hasattr(event, 'keysym') and (event.keysym == 'Up' and selections[0] > 0):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] - 1)
            elif selections and hasattr(event, 'keysym') and (event.keysym == 'Down'
                                                              and selections[0] < widget.size() - 1):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] + 1)
            elif not selections and hasattr(event, 'keysym') and event.keysym in ('Up', 'Down'):
                self.after_idle(widget.select_set, 0)

            if self.__date_list == widget:
                self.after_idle(self.__load_history)

    def __export_date(self):
        if self.__history_list.size() > 0:
            from portalocker import Lock
            hist_list = self.__history_list.get(0, self.__history_list.size() - 1)
            export_file = join(job_logs_export_dir, '{0}_{1}.txt'.format(self.__log_date, self.__job_name))

            with Lock(export_file, 'w') as f:
                f.write('\n'.join(hist_list))

            showinfo('Log Exported!', 'Your log has been exported to %s' % basename(export_file), parent=self)

    def __copy_date(self):
        if self.__history_list.size() > 0:
            from pyperclip import copy
            copy_text = '\n'.join(self.__history_list.get(0, self.__history_list.size() - 1))
            copy(copy_text)
            showinfo('Log Copied!', 'Your log has been copied to clipboard', parent=self)

    def __delete_date(self):
        if self.__history_list.size() > 0 and self.__log_date:
            myresponse = askokcancel(
                'Delete Notice!',
                'Deleting this job log will lose this job log forever. Would you like to proceed?',
                parent=self)

            if myresponse:
                del self.__job_log[self.__log_date]
                self.__job_log.sync()
                self.__load_gui()


class JobConsole(Toplevel):
    """
        Class to show job schedule main event logging in a GUI window by using Toplevel in tkinter
    """

    def __init__(self, parent):
        Toplevel.__init__(self)
        self.__print_lock = Lock()
        self.__print_queue = list()
        self.__console_text = None
        self.__log_class = None
        self.__parent = parent
        self.__js_info = None
        self.__prelog_list = list()
        self.bind('<Destroy>', self.__cleanup)
        self.iconbitmap(icon_path)
        self.__build()
        self.__gui_fill()
        self.__start_idle()

    def log_setup(self, log_class):
        if log_class and hasattr(log_class, 'gc'):
            self.__log_class = log_class

            if self.__js_info:
                for line in self.__js_info:
                    log_class.write_to_log(line, print_only=True)

    def __cleanup(self, event):
        if event and hasattr(event, 'widget') and event.widget == self:
            if self.__parent and hasattr(self.__parent, 'job_console'):
                self.__parent.job_console = None

            if self.__log_class and hasattr(self.__log_class, 'gc'):
                self.__log_class.gc = None

    def __start_idle(self):
        self.after(5, self.__on_idle)

    def print_gui(self, msg, sep='\n'):
        with self.__print_lock:
            self.__print_queue.append(str(msg) + sep)

    def __build(self):
        my_frame = Frame(self)
        button_frame = Frame(self)
        my_frame.pack(fill="both")
        button_frame.pack(fill="both")

        self.geometry('660x445+400+170')
        self.title('Job Scheduler Console')
        self.resizable(False, False)

        xbar = Scrollbar(my_frame, orient='horizontal')
        ybar = Scrollbar(my_frame, orient='vertical')
        self.__console_text = Text(my_frame, bg="black", fg="white", wrap="none", yscrollcommand=ybar,
                                   xscrollcommand=xbar)
        xbar.config(command=self.__console_text.xview)
        ybar.config(command=self.__console_text.yview)
        self.__console_text.grid(row=0, column=0)
        xbar.grid(row=1, column=0, sticky=W + E)
        ybar.grid(row=0, column=1, sticky=N + S)

        # Apply button Widgets to button_frame
        #     Start Scheduler Button
        button = Button(button_frame, text='Start Scheduler', width=13, command=self.__js_start)
        button.pack(side=LEFT, padx=7, pady=5)

        #     Close Console Button
        button = Button(button_frame, text='Close Console', width=13, command=self.destroy)
        button.pack(side=RIGHT, padx=7, pady=5)

        #     Stop Scheduler Button
        button = Button(button_frame, text='Stop Scheduler', width=13, command=self.__js_stop)
        button.pack(side=TOP, pady=5)

    def __gui_fill(self):
        if self.__parent and hasattr(self.__parent, 'js_alive') and self.__parent.js_alive:
            self.__js_info = list(["Job Scheduler Thread is currently running"])
            self.__js_info += self.__parent.js_info()

    def __on_idle(self):
        with self.__print_lock:
            for msg in self.__print_queue:
                self.__console_text.insert(END, msg)
                self.__console_text.see(END)

            self.__print_queue = []

        self.after(5, self.__on_idle)

    def __js_start(self):
        self.__parent.start_js()

    def __js_stop(self):
        self.__parent.stop_js()


class EmailModify(Toplevel):
    """
            Class to adjust cc_email and to_email for a job profile by using Toplevel in tkinter
    """

    def __init__(self, parent):
        Toplevel.__init__(self)

        self.__header = ['Welcome to Email Distro Setup!', 'Please fill out the information below:']
        self.__parent = parent
        self.__cc_email = StringVar()
        self.__to_email = StringVar()

        self.iconbitmap(icon_path)
        self.__build()
        self.__load_gui()

    @property
    def to_email(self):
        if self.__to_email.get():
            return self.__to_email.get()
        else:
            return None

    @to_email.setter
    def to_email(self, to_email):
        if to_email is None:
            self.__to_email.set('')
        else:
            self.__to_email.set(to_email)

    @property
    def cc_email(self):
        if self.__cc_email.get():
            return self.__cc_email.get()
        else:
            return None

    @cc_email.setter
    def cc_email(self, cc_email):
        if cc_email is None:
            self.__cc_email.set('')
        else:
            self.__cc_email.set(cc_email)

    def __build(self):
        # Set GUI Geometry and GUI Title
        self.geometry('280x155+670+350')
        self.title('Email Distro Setup')
        self.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self)
        email_frame = LabelFrame(self, text='Email Distro Settings')
        buttons_frame = Frame(self)

        # Apply Frames into GUI
        header_frame.pack(fill="both")
        email_frame.pack(fill="both")
        buttons_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self, text='\n'.join(self.__header), width=500, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Widgets to email_frame
        #    Apply To Email Entry Box Widget
        to_email_label = Label(email_frame, text='To Email:')
        to_email_entry = Entry(email_frame, textvariable=self.__to_email)
        to_email_label.grid(row=0, column=0, padx=4, pady=5)
        to_email_entry.grid(row=0, column=1, columnspan=3, padx=4, pady=5)

        #    Apply To Email Entry Box Widget
        cc_email_label = Label(email_frame, text='Cc Email:')
        cc_email_entry = Entry(email_frame, textvariable=self.__cc_email)
        cc_email_label.grid(row=1, column=0, padx=4, pady=5)
        cc_email_entry.grid(row=1, column=1, columnspan=3, padx=4, pady=5)

        # Apply Widgets to buttons_frame
        #     Save button
        save_button = Button(self, text='Save', width=13, command=self.__save)
        save_button.pack(in_=buttons_frame, side=LEFT, padx=7, pady=7)

        #     Cancel button
        cancel_button = Button(self, text='Cancel', width=13, command=self.destroy)
        cancel_button.pack(in_=buttons_frame, side=RIGHT, padx=7, pady=7)

    def __load_gui(self):
        if self.__parent and self.__parent.to_email:
            self.to_email = self.__parent.to_email

        if self.__parent and self.__parent.cc_email:
            self.cc_email = self.__parent.cc_email

    def __save(self):
        if not self.to_email:
            showerror('Field Empty Error!', 'No value has been inputed for To Email', parent=self)
        elif self.to_email.find('@') < 0:
            showerror('Email To Address Error!', '@ is not in the email to address field', parent=self)
        elif self.cc_email and self.cc_email.find('@') < 0:
            showerror('Email CC Address Error!', '@ is not in the email cc address field', parent=self)
        else:
            self.__parent.to_email = self.to_email
            self.__parent.cc_email = self.cc_email
            self.destroy()


class JobProfile(object):
    __schedule_date_entry = None
    __schedule_list = None
    __monday_chkbox = None
    __tuesday_chkbox = None
    __wednesday_chkbox = None
    __thursday_chkbox = None
    __friday_chkbox = None
    __saturday_chkbox = None
    __sunday_chkbox = None
    __task_label = None
    __task_param_label = None
    __task_param_entry = None
    __task_scomm_entry = None
    __task_list = None
    __task_file_dir = None
    __task_mod_button = None
    __task_up_button = None
    __task_down_button = None
    __task_del_button = None
    __schedule_mod_button = None
    __schedule_del_button = None

    """
            Class to add or modify a job profile by using Toplevel in tkinter
    """

    def __init__(self, parent, job_profile=None):
        self.__main = Toplevel()
        self.__parent = parent
        self.__job_profile = job_profile
        self.__main.bind('<Destroy>', self.__fix_focus)
        self.__schedules = list()
        self.__tasks = list()
        self.__to_email = None
        self.__cc_email = None
        self.__job_name = StringVar()
        self.__schedule_date = StringVar()
        self.__task = StringVar()
        self.__task_param = StringVar()
        self.__task_scomm = StringVar()
        self.__timeout_hh = IntVar()
        self.__timeout_mm = IntVar()
        self.__schedule_hh = IntVar()
        self.__schedule_mm = IntVar()
        self.__schedule_freq = IntVar()
        self.__monday = IntVar()
        self.__tuesday = IntVar()
        self.__wednesday = IntVar()
        self.__thursday = IntVar()
        self.__friday = IntVar()
        self.__saturday = IntVar()
        self.__sunday = IntVar()
        self.__task_attach = IntVar()
        self.__task_type = IntVar()

        if job_profile:
            self.__button_name = 'Modify Job'
            self.__title = 'Modify Existing Job'
            self.__header = ["Please Modify/Delete Job below.", "Press 'Modify Job' or 'Delete Job' when finished"]
        else:
            self.__button_name = 'Add Job'
            self.__title = 'Add New Job'
            self.__header = ["Please Add a new Job below.", "Press 'Add Job' when finished"]

        self.__main.iconbitmap(icon_path)
        self.__build()
        self.__load_gui()

    def __del__(self):
        self.__cleanup()

    def __cleanup(self):
        self.__clean_objs(self)

    def __clean_objs(self, obj):
        keys = list(vars(obj).keys())

        for key in keys:
            if key == '_JobProfile__main' and hasattr(vars(obj)[key], '__dict__'):
                self.__clean_objs(vars(obj)[key])

            del vars(obj)[key]

        collect()

    @property
    def to_email(self):
        return self.__to_email

    @to_email.setter
    def to_email(self, to_email):
        self.__to_email = to_email

    @property
    def cc_email(self):
        return self.__cc_email

    @cc_email.setter
    def cc_email(self, cc_email):
        self.__cc_email = cc_email

    @property
    def task(self):
        if self.__task.get():
            return self.__task.get()
        else:
            return None

    @task.setter
    def task(self, task):
        if task is None:
            self.__task.set('')
        else:
            self.__task.set(task)

    @property
    def task_param(self):
        if self.__task_param.get():
            return self.__task_param.get()
        else:
            return None

    @task_param.setter
    def task_param(self, task_param):
        if task_param is None:
            self.__task_param.set('')
        else:
            self.__task_param.set(task_param)

    @property
    def task_scomm(self):
        if self.__task_scomm.get():
            return self.__task_scomm.get()
        else:
            return None

    @task_scomm.setter
    def task_scomm(self, task_scomm):
        if task_scomm is None:
            self.__task_scomm.set('')
        else:
            self.__task_scomm.set(task_scomm)

    def __fix_focus(self, event):
        if event and hasattr(event, 'widget') and event.widget == self.__main:
            if self.__job_profile and 'Job_Name' in self.__job_profile.keys():
                self.__parent.load_gui(self.__job_profile['Job_Name'])
            elif self.__job_name.get():
                self.__parent.load_gui(self.__job_name.get())

    def __build(self):
        # Set GUI Geometry and GUI Title
        self.__main.geometry('657x520+500+90')
        self.__main.title(self.__title)
        self.__main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.__main)
        main_frame = Frame(self.__main)
        name_frame = LabelFrame(main_frame, text='Information')
        schedule_frame = LabelFrame(main_frame, text='Schedule(s)')
        start_frame = Frame(schedule_frame)
        freq_frame = Frame(schedule_frame)
        schedule_list_frame = Frame(schedule_frame)
        schedule_right_frame = Frame(schedule_frame)
        task_frame = LabelFrame(main_frame, text='Task(s)')
        task_field_frame = Frame(task_frame)
        task_list_frame = Frame(task_frame)
        task_button_frame = Frame(task_frame)
        button_frame = Frame(self.__main)

        # Apply Frames into GUI
        header_frame.pack(fill="both")
        main_frame.pack(fill="both")
        name_frame.grid(row=0, column=0, columnspan=2, sticky=W + E)
        schedule_frame.grid(row=1, column=0)
        start_frame.grid(row=0, column=0)
        freq_frame.grid(row=1, column=0)
        schedule_list_frame.grid(row=2, column=0)
        schedule_right_frame.grid(row=0, column=1, rowspan=3, sticky=N + S)
        task_frame.grid(row=1, column=1, sticky=N + S)
        task_field_frame.grid(row=0, column=0)
        task_list_frame.grid(row=1, column=0, sticky=N + S)
        task_button_frame.grid(row=0, column=1, rowspan=2, sticky=N + S)
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.__main, text='\n'.join(self.__header), width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply widgets to Name Frame
        #     Job Name Input Box
        label = Label(name_frame, text='Job Name:')
        txtbox = Entry(name_frame, textvariable=self.__job_name, width=55)
        label.grid(row=0, column=0, padx=6, pady=5, sticky='e')
        txtbox.grid(row=0, column=1, padx=0, pady=5, sticky='e')

        #     Timeout Dropdown Menu's
        label = Label(name_frame, text="Timeout:")
        hh_dropdown = OptionMenu(name_frame, self.__timeout_hh, *range(0, 24))
        hh_label = Label(name_frame, text='HH')
        mm_dropdown = OptionMenu(name_frame, self.__timeout_mm, *range(0, 60))
        mm_label = Label(name_frame, text='MM')
        label.grid(row=0, column=2, padx=13, pady=5, sticky='w')
        hh_dropdown.grid(row=0, column=3, padx=0, pady=5, sticky='w')
        hh_label.grid(row=0, column=4, padx=0, pady=5, sticky='w')
        mm_dropdown.grid(row=0, column=5, padx=0, pady=5, sticky='w')
        mm_label.grid(row=0, column=6, padx=0, pady=5, sticky='w')

        # Apply widgets to Start Frame
        #     Start Date Widgets
        date_label = Label(start_frame, text='Start Date:')
        self.__schedule_date_entry = DateEntry(start_frame, textvariable=self.__schedule_date, width=12,
                                               background='darkblue', foreground='white', borderwidth=2)
        date_label.grid(row=0, column=0, padx=4, pady=5, sticky='e')
        self.__schedule_date_entry.grid(row=0, column=1, columnspan=3, padx=2, pady=5, sticky='w')
        hh_dropdown = OptionMenu(start_frame, self.__schedule_hh, *range(0, 24))
        hh_label = Label(start_frame, text='HH')
        mm_dropdown = OptionMenu(start_frame, self.__schedule_mm, *range(0, 60))
        mm_label = Label(start_frame, text='MM')
        hh_dropdown.grid(row=1, column=0, padx=0, pady=5, sticky='e')
        hh_label.grid(row=1, column=1, padx=0, pady=5, sticky='e')
        mm_dropdown.grid(row=1, column=2, padx=0, pady=5, sticky='e')
        mm_label.grid(row=1, column=3, padx=0, pady=5, sticky='e')

        # Apply widgets to Freq Frame
        #     Frequency Radio Buttons
        radio1 = Radiobutton(freq_frame, text='Daily', variable=self.__schedule_freq, value=0,
                             command=lambda: self.__day_to_day_toggle(None))
        radio2 = Radiobutton(freq_frame, text='Weekly', variable=self.__schedule_freq, value=1,
                             command=lambda: self.__day_to_day_toggle(None))
        radio3 = Radiobutton(freq_frame, text='Bi-Weekly', variable=self.__schedule_freq, value=2,
                             command=lambda: self.__day_to_day_toggle(None))
        radio4 = Radiobutton(freq_frame, text='Monthly', variable=self.__schedule_freq, value=3,
                             command=lambda: self.__day_to_day_toggle(None))
        radio1.grid(row=0, column=0, padx=2, pady=4, sticky='w')
        radio2.grid(row=0, column=1, padx=2, pady=4, sticky='w')
        radio3.grid(row=1, column=0, padx=2, pady=4, sticky='w')
        radio4.grid(row=1, column=1, padx=2, pady=4, sticky='w')

        # Apply List Widget to schedule_list_frame
        #     Schedule List
        xbar = Scrollbar(schedule_list_frame, orient='horizontal')
        ybar = Scrollbar(schedule_list_frame, orient='vertical')
        self.__schedule_list = Listbox(schedule_list_frame, selectmode=SINGLE, width=30, yscrollcommand=ybar,
                                       xscrollcommand=xbar)
        xbar.config(command=self.__schedule_list.xview)
        ybar.config(command=self.__schedule_list.yview)
        self.__schedule_list.grid(row=0, column=0, padx=8, pady=5)
        xbar.grid(row=1, column=0, sticky=W + E)
        ybar.grid(row=0, column=1, sticky=N + S)
        self.__schedule_list.bind("<<ListboxSelect>>", self.__list_action)
        self.__schedule_list.bind("<Down>", self.__list_action)
        self.__schedule_list.bind("<Up>", self.__list_action)

        # Apply List Widget to schedule_list_frame
        #     Add/Del Buttons
        add_button = Button(schedule_right_frame, text='Add', width=5)
        self.__schedule_del_button = Button(schedule_right_frame, text='Del', width=4)
        self.__schedule_mod_button = Button(schedule_right_frame, text='Modify', width=12)
        add_button.grid(row=0, column=0, sticky='w', pady=5)
        self.__schedule_del_button.grid(row=0, column=1, sticky='w', padx=6, pady=5)
        self.__schedule_mod_button.grid(row=1, column=0, columnspan=2, sticky='w', pady=5)
        add_button.bind("<ButtonRelease-1>", self.__schedule_action)
        self.__schedule_mod_button.bind("<1>", self.__schedule_action)
        self.__schedule_del_button.bind("<ButtonRelease-1>", self.__schedule_action)

        #    Monday - Sunday Checkboxes
        self.__monday_chkbox = Checkbutton(schedule_right_frame, text='Monday', variable=self.__monday)
        self.__tuesday_chkbox = Checkbutton(schedule_right_frame, text='Tuesday', variable=self.__tuesday)
        self.__wednesday_chkbox = Checkbutton(schedule_right_frame, text='Wednesday', variable=self.__wednesday)
        self.__thursday_chkbox = Checkbutton(schedule_right_frame, text='Thursday', variable=self.__thursday)
        self.__friday_chkbox = Checkbutton(schedule_right_frame, text='Friday', variable=self.__friday)
        self.__saturday_chkbox = Checkbutton(schedule_right_frame, text='Saturday', variable=self.__saturday)
        self.__sunday_chkbox = Checkbutton(schedule_right_frame, text='Sunday', variable=self.__sunday)

        # Apply widgets to task_field_frame
        #     Exec Filepath/Stored Procedure Entry
        self.__task_label = Label(task_field_frame, text='Filepath:')
        entry = Entry(task_field_frame, textvariable=self.__task, width=15)
        self.__task_label.grid(row=0, column=0, padx=3, pady=5, sticky='e')
        entry.grid(row=0, column=1, padx=7, pady=5, sticky='w')

        #     Params/Filename Prefix Entry
        self.__task_param_label = Label(task_field_frame, text='Param(s):')
        self.__task_param_entry = Entry(task_field_frame, textvariable=self.__task_param, width=15)
        self.__task_param_label.grid(row=1, column=0, padx=2, pady=5, sticky='e')
        self.__task_param_entry.grid(row=1, column=1, padx=7, pady=5, sticky='w')

        #     PowerShell Command Entry Box
        label = Label(task_field_frame, text='Shell Com:')
        self.__task_scomm_entry = Entry(task_field_frame, textvariable=self.__task_scomm, width=15)
        label.grid(row=2, column=0, padx=2, pady=5, sticky='e')
        self.__task_scomm_entry.grid(row=2, column=1, padx=7, pady=5, sticky='w')

        # Apply widgets to task_list_frame
        #     Task List Box
        xbar = Scrollbar(task_list_frame, orient='horizontal')
        ybar = Scrollbar(task_list_frame, orient='vertical')
        self.__task_list = Listbox(task_list_frame, selectmode=SINGLE, width=30, height=15, yscrollcommand=ybar,
                                   xscrollcommand=xbar)
        xbar.config(command=self.__task_list.xview)
        ybar.config(command=self.__task_list.yview)
        self.__task_list.grid(row=0, column=0, padx=8, pady=5)
        xbar.grid(row=1, column=0, sticky=W + E)
        ybar.grid(row=0, column=1, sticky=N + S)
        self.__task_list.bind("<<ListboxSelect>>", self.__list_action)
        self.__task_list.bind("<Down>", self.__list_action)
        self.__task_list.bind("<Up>", self.__list_action)

        # Apply widgets to task_button_frame
        #    Task Directory Button
        self.__task_file_dir = Button(task_button_frame, text='Find File', width=13, command=self.__task_dir_show)
        self.__task_file_dir.grid(row=0, column=0, columnspan=2, padx=3, pady=4, sticky='w')

        #     Task Type Radio Buttons
        radio1 = Radiobutton(task_button_frame, text='Exec Program', variable=self.__task_type,
                             value=0, command=lambda: self.__task_type_toggle(None))
        radio2 = Radiobutton(task_button_frame, text='Exec Stored Proc', variable=self.__task_type,
                             value=1, command=lambda: self.__task_type_toggle(None))
        radio1.grid(row=1, column=0, columnspan=2, padx=3, pady=4, sticky='w')
        radio2.grid(row=2, column=0, columnspan=2, padx=3, pady=4, sticky='w')

        #     Task Buttons
        add_button = Button(task_button_frame, text='Add', width=5)
        self.__task_del_button = Button(task_button_frame, text='Del', width=5)
        self.__task_mod_button = Button(task_button_frame, text='Modify Task', width=13)
        self.__task_up_button = Button(task_button_frame, text='Task Up', width=13)
        self.__task_down_button = Button(task_button_frame, text='Task Down', width=13)
        add_button.grid(row=3, column=0, sticky='w', padx=3, pady=5)
        self.__task_del_button.grid(row=3, column=1, sticky='w', pady=5, padx=6)
        self.__task_mod_button.grid(row=4, column=0, columnspan=2, padx=3, sticky='w', pady=5)
        self.__task_up_button.grid(row=5, column=0, columnspan=2, padx=3, sticky='w', pady=5)
        self.__task_down_button.grid(row=6, column=0, columnspan=2, padx=3, sticky='w', pady=5)
        add_button.bind("<ButtonRelease-1>", self.__task_action)
        self.__task_mod_button.bind("<1>", self.__task_action)
        self.__task_up_button.bind("<1>", self.__task_action)
        self.__task_down_button.bind("<1>", self.__task_action)
        self.__task_del_button.bind("<ButtonRelease-1>", self.__task_action)

        # Apply Buttons to the button_frame
        #     Add/Modify Job Button
        button = Button(button_frame, text=self.__button_name, width=23, command=self.__job_submit)
        button.grid(row=0, column=0, pady=6, padx=15, sticky='w')

        #     Mail Settings GUI Button
        button = Button(button_frame, text='Mail Settings', width=23, command=self.__mail_settings)
        button.grid(row=0, column=1, pady=6, padx=32)

        #     Cancel Button
        button = Button(button_frame, text='Cancel', width=23, command=self.__main.destroy)
        button.grid(row=0, column=2, pady=6, padx=30, sticky='e')

    def __load_gui(self):
        if self.__job_profile:
            if 'Job_Name' in self.__job_profile.keys():
                self.__job_name.set(self.__job_profile['Job_Name'])

            if 'Timeout_HH' in self.__job_profile.keys():
                self.__timeout_hh.set(self.__job_profile['Timeout_HH'])

            if 'Timeout_MM' in self.__job_profile.keys():
                self.__timeout_mm.set(self.__job_profile['Timeout_MM'])

            if 'To_Email' in self.__job_profile.keys():
                self.to_email = self.__job_profile['To_Email']

            if 'Cc_Email' in self.__job_profile.keys():
                self.cc_email = self.__job_profile['Cc_Email']

            if 'Schedules' in self.__job_profile.keys():
                self.__schedules = self.__job_profile['Schedules']

            if 'Tasks' in self.__job_profile.keys():
                self.__tasks = self.__job_profile['Tasks']

            if self.__schedules:
                for schedule in self.__schedules:
                    self.__schedule_list.insert('end', schedule['Schedule_Name'])

            if self.__tasks:
                for task in self.__tasks:
                    self.__task_list.insert('end', task['Task_Name'])

        self.__button_toggle()

    def __list_action(self, event):
        widget = event.widget

        if widget.size() > 0:
            selections = widget.curselection()

            if selections and hasattr(event, 'keysym') and (event.keysym == 'Up' and selections[0] > 0):
                self.__main.after_idle(widget.select_clear, selections[0])
                self.__main.after_idle(widget.select_set, selections[0] - 1)
            elif selections and hasattr(event, 'keysym') and (event.keysym == 'Down'
                                                              and selections[0] < widget.size() - 1):
                self.__main.after_idle(widget.select_clear, selections[0])
                self.__main.after_idle(widget.select_set, selections[0] + 1)
            elif not selections and hasattr(event, 'keysym') and event.keysym in ('Up', 'Down'):
                self.__main.after_idle(widget.select_set, 0)

        self.__main.after_idle(self.__button_toggle)

    def __button_toggle(self):
        task_state = DISABLED
        schedule_state = DISABLED

        if self.__schedule_list.size() > 0 and self.__schedule_list.curselection():
            schedule_state = NORMAL

        if self.__task_list.size() > 0 and self.__task_list.curselection():
            task_state = NORMAL

        self.__schedule_mod_button.configure(state=schedule_state)
        self.__schedule_del_button.configure(state=schedule_state)
        self.__task_mod_button.configure(state=task_state)
        self.__task_up_button.configure(state=task_state)
        self.__task_down_button.configure(state=task_state)
        self.__task_del_button.configure(state=task_state)

    def __day_to_day_toggle(self, event):
        if self.__schedule_freq.get() in (1, 2):
            self.__monday_chkbox.grid(row=2, column=0, columnspan=2, padx=2, pady=3, sticky='w')
            self.__tuesday_chkbox.grid(row=3, column=0, columnspan=2, padx=2, pady=3, sticky='w')
            self.__wednesday_chkbox.grid(row=4, column=0, columnspan=2, padx=2, pady=3, sticky='w')
            self.__thursday_chkbox.grid(row=5, column=0, columnspan=2, padx=2, pady=3, sticky='w')
            self.__friday_chkbox.grid(row=6, column=0, columnspan=2, padx=2, pady=3, sticky='w')
            self.__saturday_chkbox.grid(row=7, column=0, columnspan=2, padx=2, pady=3, sticky='w')
            self.__sunday_chkbox.grid(row=8, column=0, columnspan=2, padx=2, pady=3, sticky='w')
        else:
            self.__monday_chkbox.grid_remove()
            self.__tuesday_chkbox.grid_remove()
            self.__wednesday_chkbox.grid_remove()
            self.__thursday_chkbox.grid_remove()
            self.__friday_chkbox.grid_remove()
            self.__saturday_chkbox.grid_remove()
            self.__sunday_chkbox.grid_remove()

    def __task_type_toggle(self, event):
        self.task = None
        self.task_param = None
        self.task_scomm = None

        if self.__task_type.get() == 1:
            self.__task_label.configure(text='Stored Proc:')
            self.__task_param_label.configure(text='Tab Name:')
            self.__task_file_dir.configure(state=DISABLED)
            self.__task_scomm_entry.configure(state=DISABLED)
        else:
            self.__task_label.configure(text='Filepath:')
            self.__task_param_label.configure(text='Param(s):')
            self.__task_file_dir.configure(state=NORMAL)
            self.__task_scomm_entry.configure(state=NORMAL)

    def __task_dir_show(self):
        from tkinter.filedialog import askopenfilename
        file_name = None

        if self.task and exists(self.task):
            init_dir = dirname(self.task)

            if init_dir != self.task:
                file_name = basename(self.task)
        else:
            init_dir = '/'

        file = askopenfilename(initialdir=init_dir, title='Select Executable File', initialfile=file_name, parent=self)

        if file:
            self.task = file

    def __schedule_action(self, event):
        selections = self.__schedule_list.curselection()
        widget = event.widget

        if widget.cget('text') == 'Add':
            self.__add_schedule()
        elif widget.cget('text') == 'Modify' and selections:
            self.__modify_schedule(selections[0])
        elif widget.cget('text') == 'Del' and selections:
            self.__delete_schedule(selections[0])

    def __task_action(self, event):
        selections = self.__task_list.curselection()
        widget = event.widget

        if widget.cget('text') == 'Add':
            self.__add_task()
        elif widget.cget('text') == 'Modify Task' and selections:
            self.__modify_task(selections[0])
        elif widget.cget('text') == 'Del' and selections:
            self.__delete_task(selections[0])
        elif widget.cget('text') == 'Task Up' and selections:
            self.__task_up(selections[0])
        elif widget.cget('text') == 'Task Down' and selections:
            self.__task_down(selections[0])

    def __task_up(self, selection):
        if selection > 0:
            prev_task = deepcopy(self.__tasks[selection - 1])
            self.__tasks[selection - 1] = deepcopy(self.__tasks[selection])
            self.__tasks[selection] = prev_task
            self.__task_list.delete(selection)
            self.__task_list.insert(selection - 1, self.__tasks[selection - 1]['Task_Name'])
            self.__main.after_idle(self.__task_list.select_clear, selection)
            self.__main.after_idle(self.__task_list.select_set, selection - 1)

    def __task_down(self, selection):
        if selection < self.__task_list.size() - 1:
            prev_task = deepcopy(self.__tasks[selection + 1])
            self.__tasks[selection + 1] = deepcopy(self.__tasks[selection])
            self.__tasks[selection] = prev_task
            self.__task_list.delete(selection)
            self.__task_list.insert(selection + 1, self.__tasks[selection + 1]['Task_Name'])
            self.__main.after_idle(self.__task_list.select_clear, selection)
            self.__main.after_idle(self.__task_list.select_set, selection + 1)

    def __add_task(self):
        if self.__task_type.get() == 1 and not self.task:
            showerror('Field Empty Error!', 'No value has been inputed for Stored Proc', parent=self)
        elif self.__task_type.get() == 0 and not self.task:
            showerror('Field Empty Error!', 'No value has been inputed for Filepath', parent=self)
        else:
            new_task = dict()
            new_task['Task_Type'] = self.__task_type.get()
            new_task['Task'] = self.task

            if self.__tasks:
                for task in self.__tasks:
                    if task['Task'] == new_task['Task']:
                        if self.__task_type.get() == 1:
                            showerror('Task Exists!', 'Stored Proc already exists!', parent=self)
                        else:
                            showerror('Task Exists!', 'Filepath already exists!', parent=self)

                        return

            if new_task['Task_Type'] == 0:
                new_task['Task'] = str(new_task['Task']).replace('/', '\\')
                new_task['Task_Name'] = '{0} ({1})'.format(basename(new_task['Task']), hash(new_task['Task']))
                new_task['Params'] = self.task_param
            else:
                new_task['Task_Name'] = new_task['Task']
                new_task['Tab_Name'] = self.task_param

            new_task['Task_SComm'] = self.task_scomm
            self.__tasks.append(new_task)
            self.__task_list.insert('end', new_task['Task_Name'])
            self.__task_type.set(0)
            self.task = None
            self.task_param = None
            self.task_scomm = None
            self.__task_type_toggle(None)

    def __modify_task(self, selection):
        task_name = self.__task_list.get(selection)

        if self.__tasks:
            for task in self.__tasks:
                if task['Task_Name'] == task_name:
                    self.__task_type.set(task['Task_Type'])
                    self.__task_type_toggle(None)
                    self.task = task['Task']
                    self.task_scomm = task['Task_SComm']

                    if self.__task_type.get() == 0:
                        self.task_param = task['Params']
                    else:
                        self.task_param = task['Tab_Name']

                    self.__tasks.remove(task)
                    break

        self.__task_list.delete(selection)

    def __delete_task(self, selection):
        myresponse = askokcancel(
            'Delete Notice!',
            'Deleting this task will lose this task forever. Would you like to proceed?',
            parent=self)

        if myresponse:
            task_name = self.__task_list.get(selection)

            if self.__tasks:
                for task in self.__tasks:
                    if task['Task_Name'] == task_name:
                        self.__tasks.remove(task)
                        break

            self.__task_list.delete(selection)

    def __add_schedule(self):
        if not self.__schedule_date.get():
            showerror('Field Empty Error!', 'No value has been inputed for Schedule Date', parent=self)
        elif self.__schedule_freq.get() in (1, 2) and self.__monday.get() != 1 and self.__tuesday.get() != 1 \
                and self.__wednesday.get() != 1 and self.__thursday.get() != 1 and self.__friday.get() != 1 \
                and self.__saturday.get() != 1 and self.__sunday.get() != 1:
            showerror('Selection Error!', 'No Monday-Friday checkbox selection for frequency', parent=self)
        else:
            from datetime import datetime, time
            new_schedule = dict()
            new_schedule['Start_Date'] = self.__schedule_date_entry.get_date()
            new_schedule['Start_HH'] = self.__schedule_hh.get()
            new_schedule['Start_MM'] = self.__schedule_mm.get()
            new_schedule['Frequency'] = self.__schedule_freq.get()
            new_schedule['Start_Datetime'] = datetime.combine(new_schedule['Start_Date'], time(
                hour=new_schedule['Start_HH'], minute=new_schedule['Start_MM']))

            if new_schedule['Start_Datetime'] <= datetime.now():
                showerror('Start Date & Time Error!',
                          'Start Date and Time is before now. Please select something in future', parent=self)
                return

            if new_schedule['Frequency'] == 0:
                freq_name = 'Daily'
            elif new_schedule['Frequency'] == 1:
                freq_name = 'Weekly'
            elif new_schedule['Frequency'] == 2:
                freq_name = 'Bi-Weekly'
            elif new_schedule['Frequency'] == 3:
                freq_name = 'Monthly'
            else:
                freq_name = None

            new_schedule['Schedule_Name'] = '{0} - {1}'.format(new_schedule['Start_Datetime'], freq_name)

            if self.__schedules:
                for schedule in self.__schedules:
                    if schedule['Schedule_Name'] == new_schedule['Schedule_Name']:
                        showerror('Schedule Exists!', 'Start Date & Frequency already exists!', parent=self)
                        return

            if self.__schedule_freq.get() in (1, 2):
                new_schedule['Frequency_DOW'] = [self.__monday.get(), self.__tuesday.get(), self.__wednesday.get(),
                                                 self.__thursday.get(), self.__friday.get(), self.__saturday.get(),
                                                 self.__sunday.get()]
            self.__schedules.append(new_schedule)
            self.__schedule_list.insert('end', new_schedule['Schedule_Name'])
            self.__schedule_date.set(datetime.today().__format__('%m/%d/%Y'))
            self.__schedule_hh.set(0)
            self.__schedule_mm.set(0)
            self.__schedule_freq.set(0)
            self.__monday.set(0)
            self.__tuesday.set(0)
            self.__wednesday.set(0)
            self.__thursday.set(0)
            self.__friday.set(0)
            self.__saturday.set(0)
            self.__sunday.set(0)
            self.__day_to_day_toggle(None)

    def __modify_schedule(self, selection):
        schedule_name = self.__schedule_list.get(selection)

        if self.__schedules:
            for schedule in self.__schedules:
                if schedule['Schedule_Name'] == schedule_name:
                    self.__schedule_date_entry.set_date(schedule['Start_Date'])
                    self.__schedule_hh.set(schedule['Start_HH'])
                    self.__schedule_mm.set(schedule['Start_MM'])
                    self.__schedule_freq.set(schedule['Frequency'])

                    if schedule['Frequency'] in (1, 2):
                        self.__monday.set(schedule['Frequency_DOW'][0])
                        self.__tuesday.set(schedule['Frequency_DOW'][1])
                        self.__wednesday.set(schedule['Frequency_DOW'][2])
                        self.__thursday.set(schedule['Frequency_DOW'][3])
                        self.__friday.set(schedule['Frequency_DOW'][4])
                        self.__saturday.set(schedule['Frequency_DOW'][5])
                        self.__sunday.set(schedule['Frequency_DOW'][6])

                    self.__schedules.remove(schedule)
                    self.__day_to_day_toggle(None)
                    break

        self.__schedule_list.delete(selection)

    def __delete_schedule(self, selection):
        myresponse = askokcancel(
            'Delete Notice!',
            'Deleting this schedule will lose this schedule forever. Would you like to proceed?',
            parent=self)

        if myresponse:
            schedule_name = self.__schedule_list.get(selection)

            if self.__schedules:
                for schedule in self.__schedules:
                    if schedule['Schedule_Name'] == schedule_name:
                        self.__schedules.remove(schedule)
                        break

            self.__schedule_list.delete(selection)

    def __job_submit(self):
        if not self.__job_name.get():
            showerror('Field Empty Error!', 'No value has been inputed for Job Name', parent=self)
        elif self.__timeout_hh.get() == 0 and self.__timeout_mm.get() == 0:
            showerror('Field Empty Error!', 'No value has been inputed for Job Timeout HH or MM', parent=self)
        elif not self.to_email:
            showerror('Settings Empty Error!', 'Mail Settings hasn''t been established', parent=self)
        elif not self.__schedules:
            showerror('Schedule Empty Error!', 'A Schedule hasn''t been added', parent=self)
        elif not self.__tasks:
            showerror('Task Empty Error!', 'A Task hasn''t been added', parent=self)
        else:
            from New_Job_Scheduler_Class import get_next_run
            job_profile = dict()
            job_profile['Job_Name'] = self.__job_name.get()
            job_profile['Enabled'] = True
            job_profile['Manual_Start'] = False
            job_profile['Manual_Stop'] = False
            job_profile['Running'] = False
            job_profile['Timeout_HH'] = self.__timeout_hh.get()
            job_profile['Timeout_MM'] = self.__timeout_mm.get()
            job_profile['To_Email'] = self.to_email
            job_profile['Cc_Email'] = self.cc_email
            job_profile['Schedules'] = self.__schedules
            job_profile['Prev_Run'] = None
            job_profile['Next_Run'] = get_next_run(job_profile=job_profile)
            job_profile['Tasks'] = self.__tasks

            if local_config['Jobs']:
                jobs = local_config['Jobs']
            else:
                jobs = dict()

            jobs[self.__job_name.get()] = job_profile

            if self.__job_profile and 'Job_Name' in self.__job_profile.keys() and \
                    self.__job_profile['Job_Name'] != self.__job_name.get():
                del jobs[self.__job_profile['Job_Name']]

            local_config['Jobs'] = jobs
            local_config.sync()
            self.__parent.js_instance.job_configs = jobs
            self.__main.destroy()

    def __mail_settings(self):
        EmailModify(parent=self)


class JL(Tk):
    __job_list = None
    __job_log_button = None
    __modify_button = None
    __delete_button = None
    __disable_button = None
    __start_button = None
    __job_console = None
    __js_start_datetime = None
    __js_thread = None
    __js_thread_watcher = None
    __js_thread_watcher_on = False

    """
            Class to show list of job profiles by using Toplevel in tkinter
    """

    def __init__(self):
        Tk.__init__(self)
        from New_Job_Scheduler_Class import JobScheduler
        self.__js_instance = JobScheduler(parent=self)
        self.start_js()
        self.__next_run = StringVar()
        self.__prev_run = StringVar()
        self.iconbitmap(icon_path)
        self.__build()
        self.load_gui()
        self.bind('<Destroy>', self.stop_js)

    @property
    def js_instance(self):
        return self.__js_instance

    @property
    def js_thread(self):
        return self.__js_thread

    @js_thread.setter
    def js_thread(self, js_thread):
        self.__js_thread = js_thread

    @property
    def js_alive(self):
        if self.__js_thread:
            return self.__js_thread.is_alive()

    @property
    def js_start_datetime(self):
        return self.__js_start_datetime

    @property
    def job_console(self):
        return self.__job_console

    @job_console.setter
    def job_console(self, job_console):
        self.__job_console = job_console

    def start_js(self, start_watch=True):
        from datetime import datetime

        if not self.__js_thread or not self.__js_thread.is_alive():
            self.__js_start_datetime = datetime.now()
            self.__js_thread = Thread(target=self.__js_instance.start)
            self.__js_thread.daemon = True
            self.__js_thread.start()

        if start_watch and (not self.__js_thread_watcher or not self.__js_thread_watcher.is_alive()):
            self.__js_thread_watcher_on = True
            self.__js_thread_watcher = Thread(target=self.__watch_js)
            self.__js_thread_watcher.daemon = True
            self.__js_thread_watcher.start()

    def js_info(self):
        if self.__js_instance:
            return self.__js_instance.info()

    def stop_js(self, event=None):
        if not event or (event and hasattr(event, 'widget') and event.widget == self):
            from time import sleep
            self.__js_thread_watcher_on = False

            while self.__js_thread_watcher and self.__js_thread_watcher.is_alive():
                sleep(1)

            if self.__js_thread_watcher:
                self.__js_thread_watcher = None

            if self.__js_thread and self.__js_thread.is_alive():
                self.__js_instance.exit()
            elif self.__js_thread:
                self.__js_thread = None

            if self.__js_start_datetime:
                self.__js_start_datetime = None

    def __watch_js(self):
        from time import sleep
        from datetime import datetime

        start_time = datetime.now()

        while self.__js_thread_watcher_on:
            if self.__js_thread and not self.__js_thread.is_alive():
                if int(divmod((datetime.now() - start_time).total_seconds(), 60)[0] % 60) < 6:
                    tool.write_to_log("Job Scheduler ran into an issue within 5 minutes of startup. Shutting down")
                    self.__js_thread_watcher_on = False
                    self.__js_thread = None
                    break
                else:
                    tool.write_to_log("Job Scheduler ran into an issue. Restarting Scheduler in 5 minutes")
                    sleep(300)
                    self.start_js(start_watch=False)
            elif not self.__js_thread:
                self.__js_thread_watcher_on = False
                break

            sleep(1)

        self.__js_thread_watcher = None

    def __build(self):
        header = ['Welcome to Job Scheduler Control!', 'Feel free to control a job or open console']

        # Set GUI Geometry and GUI Title
        self.geometry('420x285+630+290')
        self.title('Job Scheduler GUI')
        self.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self)
        main_frame = LabelFrame(self, text='Jobs')
        list_frame = Frame(main_frame)
        run_frame = Frame(main_frame)
        control_frame = Frame(main_frame)
        button_frame = Frame(self)

        # Apply Frames into GUI
        header_frame.pack(fill="both")
        main_frame.pack(fill="both")
        list_frame.grid(row=0, column=0, rowspan=2)
        run_frame.grid(row=0, column=1, padx=5)
        control_frame.grid(row=1, column=1, padx=5)
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self, text='\n'.join(header), width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply List Widget to list_frame
        #     Job List
        xbar = Scrollbar(list_frame, orient='horizontal')
        ybar = Scrollbar(list_frame, orient='vertical')
        self.__job_list = Listbox(list_frame, selectmode=SINGLE, width=30, yscrollcommand=ybar,
                                  xscrollcommand=xbar)
        xbar.config(command=self.__job_list.xview)
        ybar.config(command=self.__job_list.yview)
        self.__job_list.grid(row=0, column=0, padx=8, pady=5)
        xbar.grid(row=1, column=0, sticky=W + E)
        ybar.grid(row=0, column=1, sticky=N + S)
        self.__job_list.bind("<Down>", self.__list_action)
        self.__job_list.bind("<Up>", self.__list_action)
        self.__job_list.bind('<<ListboxSelect>>', self.__list_action)

        # Apply button Widgets to run_frame
        #     Prev Run Time Input Box
        prev_run_label = Label(run_frame, text='Prev Run:')
        prev_run_txtbox = Entry(run_frame, textvariable=self.__prev_run, width=20, state=DISABLED)
        prev_run_label.grid(row=0, column=0, padx=2, pady=5, sticky='e')
        prev_run_txtbox.grid(row=0, column=1, padx=3, pady=5, sticky='e')

        #     Next Run Time Input Box
        next_run_label = Label(run_frame, text='Next Run:')
        next_run_txtbox = Entry(run_frame, textvariable=self.__next_run, width=20, state=DISABLED)
        next_run_label.grid(row=1, column=0, padx=2, pady=5, sticky='e')
        next_run_txtbox.grid(row=1, column=1, padx=3, pady=5, sticky='w')

        # Apply button Widgets to control_frame
        #     Access Column Migration Buttons
        self.__job_log_button = Button(control_frame, text='Job Log', width=24)
        self.__modify_button = Button(control_frame, text='Modify Job', width=10)
        self.__delete_button = Button(control_frame, text='Delete Job', width=10, command=self.__delete_job)
        self.__disable_button = Button(control_frame, text='Disable Job', width=10)
        self.__start_button = Button(control_frame, text='Start Job', width=10)
        self.__job_log_button.grid(row=2, column=0, columnspan=2, padx=7, pady=7)
        self.__modify_button.grid(row=3, column=0, padx=7, pady=7)
        self.__delete_button.grid(row=3, column=1, padx=7, pady=7)
        self.__start_button.grid(row=4, column=0, padx=7, pady=7)
        self.__disable_button.grid(row=4, column=1, padx=7, pady=7)
        self.__job_log_button.bind("<1>", self.__job_action)
        self.__modify_button.bind("<1>", self.__job_action)
        self.__disable_button.bind("<ButtonRelease-1>", self.__job_action)
        self.__start_button.bind("<ButtonRelease-1>", self.__job_action)

        # Apply button Widgets to button_frame
        #     Mail Settings Button
        button = Button(button_frame, text='Console', width=13, command=self.__console)
        button.grid(row=0, column=0, padx=23, pady=5)

        #     Error Profiles Button
        button = Button(button_frame, text='Add Job', width=12, command=self.__add_job)
        button.grid(row=0, column=1, padx=18, pady=5)

        #     Cancel Button
        button = Button(button_frame, text='Exit Program', width=13, command=self.destroy)
        button.grid(row=0, column=2, padx=18, pady=5)

    def load_gui(self, load_job_name=None):
        if self.__job_list.size() > 0:
            self.__job_list.delete(0, self.__job_list.size() - 1)

        jobs = local_config['Jobs']
        job_list_pos = 0

        if jobs:
            for job_name in jobs.keys():
                self.__job_list.insert('end', job_name)

                if load_job_name == job_name:
                    job_list_pos = self.__job_list.size() - 1

            self.after_idle(self.__job_list.select_set, job_list_pos)
            self.after_idle(self.__job_button_state)

    def __list_action(self, event):
        widget = event.widget

        if widget.size() > 0:
            selections = widget.curselection()

            if selections and hasattr(event, 'keysym') and (event.keysym == 'Up' and selections[0] > 0):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] - 1)
            elif selections and hasattr(event, 'keysym') and (event.keysym == 'Down'
                                                              and selections[0] < widget.size() - 1):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] + 1)
            elif not selections and hasattr(event, 'keysym') and event.keysym in ('Up', 'Down'):
                self.after_idle(widget.select_set, 0)

            self.after_idle(self.__job_button_state)

    def __job_action(self, event):
        selections = self.__job_list.curselection()

        if selections:
            widget = event.widget
            jobs = local_config['Jobs']
            job_name = self.__job_list.get(selections[0])

            if widget.cget('text') == 'Job Log' and job_name in jobs.keys() and not jobs[job_name]['Running'] and\
                    not jobs[job_name]['Manual_Start'] and exists(join(job_logs_dir, '%s.log'.lower() % job_name)):
                JobLog(parent=self, job_name=job_name)
            elif widget.cget('text') == 'Modify Job' and job_name in jobs.keys():
                JobProfile(parent=self, job_profile=deepcopy(jobs[job_name]))
            elif widget.cget('text') == 'Enable Job' and job_name in jobs.keys():
                self.__enable_job(job_name=job_name, job_profile=jobs[job_name])
            elif widget.cget('text') == 'Disable Job' and job_name in jobs.keys():
                self.__disable_job(job_name=job_name, job_profile=jobs[job_name])
            elif widget.cget('text') in ('Start Job', 'Stop Job') and job_name in jobs.keys():
                self.__start_stop_job(job_name=job_name, job_profile=jobs[job_name])

    def __enable_job(self, job_name, job_profile):
        if job_profile and not job_profile['Manual_Start'] and not job_profile['Manual_Stop']\
                and not job_profile['Running']:
            from New_Job_Scheduler_Class import get_next_run
            jobs = local_config['Jobs']
            job_profile['Enabled'] = True
            job_profile['Manual_Start'] = False
            job_profile['Running'] = False
            job_profile['Next_Run'] = get_next_run(job_profile=job_profile)
            jobs[job_name] = job_profile
            local_config['Jobs'] = jobs
            local_config.sync()
            self.__job_button_state()

    def __disable_job(self, job_name, job_profile):
        if job_profile and not job_profile['Manual_Start'] and not job_profile['Manual_Stop']\
                and not job_profile['Running']:
            jobs = local_config['Jobs']
            job_profile['Enabled'] = False
            job_profile['Manual_Start'] = False
            job_profile['Manual_Stop'] = False
            job_profile['Running'] = False
            job_profile['Next_Run'] = None
            jobs[job_name] = job_profile
            local_config['Jobs'] = jobs
            local_config.sync()
            self.__job_button_state()

    def __start_stop_job(self, job_name, job_profile):
        if job_profile and not job_profile['Manual_Start'] and not job_profile['Running']:
            jobs = local_config['Jobs']
            job_profile['Manual_Start'] = True
            job_profile['Manual_Stop'] = False
            jobs[job_name] = job_profile
            local_config['Jobs'] = jobs
            local_config.sync()
            self.__job_button_state()
        elif job_profile and (job_profile['Manual_Start'] or job_profile['Running']):
            jobs = local_config['Jobs']
            job_profile['Manual_Start'] = False

            if self.js_alive is None:
                job_profile['Manual_Stop'] = False
                job_profile['Running'] = False
            else:
                job_profile['Manual_Stop'] = True

            jobs[job_name] = job_profile
            local_config['Jobs'] = jobs
            local_config.sync()
            self.__job_button_state()

    def __delete_job(self):
        selections = self.__job_list.curselection()

        if selections:
            myresponse = askokcancel(
                'Delete Notice!',
                'Deleting this job will lose this job profile forever. Would you like to proceed?',
                parent=self)

            if myresponse:
                jobs = local_config['Jobs']
                job_name = self.__job_list.get(selections[0])

                if job_name in jobs.keys():
                    del jobs[job_name]

                    if len(jobs) > 0:
                        local_config['Jobs'] = jobs
                    else:
                        del local_config['Jobs']

                    local_config.sync()

                self.__job_list.delete(selections[0], selections[0])

                if self.__job_list.size() > 0 and selections[0] == 0:
                    self.after_idle(self.__job_list.select_set, 0)
                elif self.__job_list.size() > 0:
                    self.after_idle(self.__job_list.select_set, selections[0] - 1)

    def __job_button_state(self):
        handle_dict = dict()
        selections = self.__job_list.curselection()
        handle_dict['ED_Name'] = 'Disable Job'
        handle_dict['SS_Name'] = 'Start Job'
        handle_dict['Job_Log'] = DISABLED
        handle_dict['Modify_Job'] = NORMAL
        handle_dict['Delete_Job'] = NORMAL
        handle_dict['ED_Job'] = NORMAL
        handle_dict['SS_Job'] = NORMAL
        handle_dict['Next_Run'] = ''
        handle_dict['Prev_Run'] = ''

        if selections:
            jobs = local_config['Jobs']
            job_name = self.__job_list.get(selections[0])

            if exists(join(job_logs_dir, '%s.log'.lower() % job_name)):
                handle_dict['Job_Log'] = NORMAL

            if job_name in jobs.keys():
                job_profile = jobs[job_name]

                if job_profile['Next_Run']:
                    handle_dict['Next_Run'] = job_profile['Next_Run'].__format__("%Y%m%d %I:%M:%S %p")

                if job_profile['Prev_Run']:
                    handle_dict['Prev_Run'] = job_profile['Prev_Run'].__format__("%Y%m%d %I:%M:%S %p")

                if not job_profile['Enabled']:
                    handle_dict['ED_Name'] = 'Enable Job'
                    handle_dict['Modify_Job'] = DISABLED
                    handle_dict['SS_Job'] = DISABLED

                if job_profile['Running'] or job_profile['Manual_Start']:
                    handle_dict['SS_Name'] = 'Stop Job'
                    handle_dict['ED_Job'] = DISABLED
                    handle_dict['Modify_Job'] = DISABLED
                    handle_dict['Delete_Job'] = DISABLED
                    handle_dict['Job_Log'] = DISABLED
            else:
                self.__job_list.delete(selections[0], selections[0])

                if self.__job_list.size() > 0 and selections[0] == 0:
                    self.after_idle(self.__job_list.select_set, 0)
                elif self.__job_list.size() > 0:
                    self.after_idle(self.__job_list.select_set, selections[0] - 1)

                return
        elif self.__job_list.size() < 1:
            handle_dict['Modify_Job'] = DISABLED
            handle_dict['Delete_Job'] = DISABLED
            handle_dict['ED_Job'] = DISABLED
            handle_dict['SS_Job'] = DISABLED

        self.__job_log_button.configure(state=handle_dict['Job_Log'])
        self.__modify_button.configure(state=handle_dict['Modify_Job'])
        self.__delete_button.configure(state=handle_dict['Delete_Job'])
        self.__disable_button.configure(state=handle_dict['ED_Job'])
        self.__start_button.configure(state=handle_dict['SS_Job'])
        self.__disable_button.configure(text=handle_dict['ED_Name'])
        self.__start_button.configure(text=handle_dict['SS_Name'])
        self.__next_run.set(handle_dict['Next_Run'])
        self.__prev_run.set(handle_dict['Prev_Run'])

    def __add_job(self):
        JobProfile(parent=self)

    def __console(self):
        if not self.__job_console:
            self.__job_console = JobConsole(parent=self)
            tool.gui_console(gui_obj=self.__job_console)
        else:
            self.after(1, lambda: self.__job_console.deiconify())
            self.after(1, lambda: self.__job_console.focus_set())
            self.after(1, lambda: self.__job_console.geometry('+400+170'))


class JobList(JL):
    def __init__(self):
        JL.__init__(self)
        self.mainloop()
