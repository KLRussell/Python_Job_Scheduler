from Global import grabobjs
from Global import ShelfHandle
from Global import SQLHandle
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from time import sleep
from multiprocessing import Process
from multiprocessing.managers import BaseManager
from subprocess import Popen, PIPE
from Job_Scheduler_Settings import next_run_date
from Job_Scheduler_Settings import add_setting

import smtplib
import zipfile
import os
import pandas as pd
import datetime
import traceback
import copy
import sys
import portalocker
import atexit
import pathlib as pl

if getattr(sys, 'frozen', False):
    from multiprocessing import freeze_support
    application_path = sys.executable
else:
    application_path = __file__

curr_dir = os.path.dirname(os.path.abspath(application_path))
main_dir = os.path.dirname(curr_dir)
batcheddir = os.path.join(main_dir, '02_Attachments')
joblogsdir = os.path.join(main_dir, '05_Job_Logs')
global_objs = grabobjs(main_dir, 'Job_Scheduler')


class Email:
    def __init__(self, job_config, job_results, attach=None, error_msg=None):
        self.email_server = global_objs['Settings'].grab_item('Email_Server')
        self.email_port = global_objs['Settings'].grab_item('Email_Port')
        self.email_user = global_objs['Settings'].grab_item('Email_User')
        self.email_pass = global_objs['Settings'].grab_item('Email_Pass')
        self.email_from = '{0}@{1}'.format(self.email_user.decrypt_text(),
                                           '.'.join(str(self.email_server.decrypt_text()).split('.')[1:]))
        self.email_to = job_config['To_Distro']
        self.email_cc = job_config['CC_Distro']
        self.server = None
        self.message = MIMEMultipart()
        self.body = list()
        self.file = attach
        self.body.append("Hello {0},\n".format(self.email_to.split('@')[0].title()))

        if error_msg:
            self.subject = "<Job Failed> Job \"{0}\"".format(job_config['Job_Name'])
            sub_body = "Job \"{0}\", {1}".format(job_config['Job_Name'], error_msg)
        else:
            self.subject = "<Job Succeeded> Job \"{0}\"".format(job_config['Job_Name'])
            sub_body = "Job \"{0}\" completed successfully".format(job_config['Job_Name'])

        self.body.append("{0}. Total job runtime was {1}.\n"
                         .format(sub_body, self.parse_time(job_config['Start_Time'], datetime.datetime.now())))
        self.parse_results(job_results)
        self.body.append("\nYour's Truly,\n")
        self.body.append("The CDA's")
        self.email_close()

    @staticmethod
    def parse_time(start_time, end_time):
        duration = end_time - start_time
        duration_in_s = duration.total_seconds()
        days = int(divmod(duration_in_s, 86400)[0])
        hours = int(divmod(duration_in_s, 3600)[0] % 24)
        minutes = int(divmod(duration_in_s, 60)[0] % 60)
        seconds = int(duration.seconds % 60)
        milliseconds = int(divmod(duration.microseconds, 1000)[0] % 1000)

        date_list = []

        if days > 0:
            if days == 1:
                date_list.append('{0} Day'.format(days))
            else:
                date_list.append('{0} Days'.format(days))

        if hours > 0:
            if hours == 1:
                date_list.append('{0} Hour'.format(hours))
            else:
                date_list.append('{0} Hours'.format(hours))

        if days < 1 and minutes > 0:
            if minutes == 1:
                date_list.append('{0} Minute'.format(minutes))
            else:
                date_list.append('{0} Minutes'.format(minutes))

        if days < 1 and hours < 1 and seconds > 0:
            if seconds == 1:
                date_list.append('{0} Second'.format(seconds))
            else:
                date_list.append('{0} Seconds'.format(seconds))

        if len(date_list) < 1 and milliseconds > 0:
            if milliseconds == 1:
                date_list.append('{0} Millisecond'.format(milliseconds))
            else:
                date_list.append('{0} Milliseconds'.format(milliseconds))

        return ' and '.join(date_list)

    def parse_results(self, job_results):
        for sub_job_name, sub_job_type, sub_start_time, sub_end_time, sub_error in job_results:
            if sub_end_time:
                self.body.append('\t\u2022  {0} "{1}" <Succeeded Task> [{2}]'.format(
                    sub_job_type, sub_job_name, self.parse_time(sub_start_time, sub_end_time)))
            else:
                self.body.append('\t\u2022  {0} "{1}" <Failed Task> [Err Code: {2}]'.format(sub_job_type, sub_job_name,
                                                                                            sub_error))

    def email_connect(self):
        try:
            self.server = smtplib.SMTP(str(self.email_server.decrypt_text()), int(self.email_port.decrypt_text()))

            try:
                self.server.ehlo()
                self.server.starttls()
                self.server.ehlo()
                self.server.login(self.email_user.decrypt_text(), self.email_pass.decrypt_text())
            except:
                self.email_close()
        except:
            pass

    def email_send(self):
        self.server.sendmail(self.email_from, self.email_to, str(self.message))

    def email_close(self):
        if self.server:
            try:
                self.server.quit()
            except:
                pass
            finally:
                del self.server

    def package_email(self):
        self.message['To'] = self.email_to
        self.message['Date'] = formatdate(localtime=True)

        if self.email_cc:
            self.message['Cc'] = self.email_cc

        self.message['Subject'] = self.subject
        self.message.attach(MIMEText('\n'.join(self.body)))

        if self.file:
            part = MIMEBase('application', "octet-stream")
            zip_filepath = self.zip_file()
            zf = open(zip_filepath, 'rb')

            try:
                part.set_payload(zf.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename="{0}"'.format(
                    os.path.basename(zip_filepath)))
                self.message.attach(part)
            finally:
                zf.close()

    def zip_file(self):
        zip_filepath = None
        i = 1

        while not zip_filepath:
            if i > 1:
                zip_filepath = os.path.join(os.path.dirname(self.file), '{0}{1}.zip'.format(
                                                     os.path.splitext(os.path.basename(self.file))[0], i))
            else:
                zip_filepath = os.path.join(os.path.dirname(self.file), '{0}.zip'.format(
                    os.path.splitext(os.path.basename(self.file))[0]))

            if os.path.exists(zip_filepath):
                i += 1
                zip_filepath = None

        zip_file = zipfile.ZipFile(zip_filepath, mode='w')

        try:
            zip_file.write(self.file, os.path.basename(self.file))
        finally:
            zip_file.close()
            os.remove(self.file)

        return zip_filepath


class JobConfig(object):
    def __init__(self, job_config, time_out):
        self.sub_job_name = []
        self.sub_job_type = []
        self.job_files = []
        self.file_path = None
        self.sub_start_time = []
        self.sub_end_time = []
        self.sub_error = []
        self.data = pd.DataFrame()
        self.asql = SQLHandle(logobj=global_objs['Event_Log'], settingsobj=global_objs['Settings'])
        self.job_config = job_config
        self.job_log = ShelfHandle(os.path.join(joblogsdir, job_config['Job_Name']))
        self.job_log.read_shelf()
        self.time_out = time_out

    def close_job(self):
        fconfig = None
        fline = None
        global_objs['Local_Settings'].read_shelf()
        my_configs = global_objs['Local_Settings'].grab_item('Job_Configs')

        for my_line, my_config in enumerate(my_configs):
            if my_config['Job_Name'].lower() == self.job_config['Job_Name'].lower():
                fconfig = my_config
                fline = my_line
                break

        if fconfig:
            fconfig['Job_Controls'] = [fconfig['Job_Controls'][0], False, 0]
            my_configs[fline] = fconfig
            add_setting('Local_Settings', my_configs, 'Job_Configs', False)

    def job_name(self):
        return self.job_config['Job_Name']

    def job_log_item(self, log_item):
        assert log_item

        log_items = self.job_log.grab_item(datetime.datetime.now().__format__("%Y%m%d"))
        self.job_log.del_item(datetime.datetime.now().__format__("%Y%m%d"))

        if not log_items:
            log_items = []

        log_items.append([datetime.datetime.now().__format__("%I:%M:%S %p"), log_item])
        self.job_log.add_item(key=datetime.datetime.now().__format__("%Y%m%d"), val=log_items, encrypt=False)
        self.job_log.write_shelf()

    def start_job(self):
        self.job_log_item("Starting Job '{0}'".format(self.job_config['Job_Name']))

        for sub_job, sub_job_type in copy.deepcopy(self.job_config['Job_List']):
            job_name = os.path.basename(sub_job[0])
            self.job_log_item("Processing {0} '{1}'".format(sub_job_type, job_name))

            self.sub_job_name.append(job_name)
            self.sub_job_type.append(sub_job_type)
            self.sub_start_time.append(datetime.datetime.now())

            try:
                if sub_job_type == 'Stored Procedure':
                    self.exec_proc(sub_job[0])

                    if self.data[1]:
                        self.job_log_item("{0} '{1}' failed [ECode {2}] - {3}"
                                          .format(sub_job_type, sub_job[0], self.data[1], self.data[2]))
                        self.sub_error.append(self.data[1])
                        self.sub_end_time.append(None)
                    else:
                        if not self.data[0].empty:
                            self.job_log_item("Stored Procedure '{0}' found {1} items to batch into an excel attachment"
                                              .format(sub_job[0], len(self.data[0])))
                            self.job_files.append([sub_job[0], sub_job[2], self.data[0]])
                        elif sub_job[1]:
                            self.job_log_item("Stored Procedure '{0}' found no items to batch into an excel attachment"
                                              .format(sub_job[0]))

                        self.sub_end_time.append(datetime.datetime.now())
                        self.sub_error.append(None)
                elif os.path.exists(sub_job[0]):
                    lines = []
                    ext = os.path.splitext(sub_job[0])[1].lower()

                    if sub_job[2]:
                        proc = Popen(['powershell.exe', "{2} '{0}' {1}".format(sub_job[0], sub_job[1], sub_job[2])],
                                     stdin=PIPE, stdout=PIPE, stderr=PIPE)
                        stdout, stderr = proc.communicate()
                    elif ext == '.py':
                        proc = Popen(['powershell.exe', "python '{0}' {1}".format(sub_job[0], sub_job[1])], stdin=PIPE,
                                     stdout=PIPE, stderr=PIPE)
                        stdout, stderr = proc.communicate()
                    elif ext == '.ps1':
                        proc = Popen(['powershell.exe', ". '{0}' {1}".format(sub_job[0], sub_job[1])], stdin=PIPE,
                                     stdout=PIPE, stderr=PIPE)
                        stdout, stderr = proc.communicate()
                    elif ext == '.vbs':
                        proc = Popen(['powershell.exe', "cscript '{0}' {1}".format(sub_job[0], sub_job[1])], stdin=PIPE,
                                     stdout=PIPE, stderr=PIPE)
                        stdout, stderr = proc.communicate()
                    elif ext == '.exe':
                        proc = Popen('{0} {1}'.format(sub_job[0], sub_job[1]), stdin=PIPE, stdout=PIPE)
                        stdout, stderr = proc.communicate()
                    elif ext == '.sql':
                        proc = None
                        stderr = None
                        self.exec_sql(sub_job[0])
                    else:
                        proc = Popen(['powershell.exe', "'{0}' {1}".format(sub_job[0], sub_job[1])],
                                     stdin=PIPE, stdout=PIPE, stderr=PIPE)
                        stdout, stderr = proc.communicate()

                    if ext == '.sql':
                        if self.data[1]:
                            self.job_log_item("{0} '{1}' failed [ECode {2}] - {3}"
                                              .format(sub_job_type, job_name, self.data[1], self.data[2]))
                            self.sub_error.append(self.data[1])
                            self.sub_end_time.append(None)
                        else:
                            if not self.data[0].empty:
                                self.job_log_item(
                                    "Program '{0}' found {1} items to batch into an excel attachment"
                                    .format(job_name, len(self.data[0])))
                                self.job_files.append([job_name, job_name, self.data[0]])

                            self.sub_end_time.append(datetime.datetime.now())
                            self.sub_error.append(None)
                    elif proc:
                        for my_line in stderr.decode("utf-8").split('\n'):
                            lines.append(my_line.rstrip())

                        if proc.returncode == 0:
                            self.sub_error.append(None)
                            self.sub_end_time.append(datetime.datetime.now())
                        else:
                            self.job_log_item("{0} '{1}' failed [ECode {2}] - {3}"
                                              .format(sub_job_type, job_name, proc.returncode,
                                                      '. '.join(lines)))
                            self.sub_error.append(proc.returncode)
                            self.sub_end_time.append(None)
                else:
                    self.job_log_item("{0} '{1}' failed [ECode 00x01] - Filepath was not found for file"
                                      .format(sub_job_type, job_name))
                    self.sub_end_time.append(None)
                    self.sub_error.append('00x01')
            except Exception as e:
                if len(self.sub_end_time) < len(self.sub_start_time):
                    self.sub_end_time.append(None)
                    self.sub_error.append(type(e).__name__)

                self.job_log_item("{0} '{1}' failed [ECode {2}] - {3}"
                                  .format(sub_job_type, job_name, type(e).__name__, str(e)))
                self.list_chksum(type(e).__name__)
                global_objs['Event_Log'].write_log(traceback.format_exc(), 'critical')
                pass

    def get_timeout(self):
        if self.time_out:
            return (self.time_out - datetime.datetime.now()).total_seconds()
        else:
            return 0

    def exec_sql(self, sql_path):
        self.job_log_item("Executing SQL File '{0}'".format(os.path.basename(sql_path)))

        with portalocker.Lock(sql_path, 'r') as f:
            query = f.read()

        if query:
            self.asql.connect('alch', query_time_out=self.get_timeout())
            self.data = self.asql.execute(str_txt=query, execute=True, ret_err=True)
            self.asql.close_conn()
        else:
            self.data = [pd.DataFrame(), '00x03', 'File has no data to execute']

    def exec_proc(self, job):
        self.job_log_item("Executing Stored Procedure '{0}'".format(job))
        self.asql.connect('alch', query_time_out=self.get_timeout())
        self.data = self.asql.execute(str_txt=job, execute=True, proc=True, ret_err=True)
        self.asql.close_conn()

    def export_files(self):
        if len(self.job_files) > 0:
            for file_num in range(0, 10000000000):
                self.file_path = os.path.join(batcheddir, '{0}_{1}_{2}'.format(
                    datetime.datetime.now().__format__("%Y%m%d"), self.job_config['Job_Name'], file_num))
                if not os.path.exists('%s.zip' % self.file_path):
                    self.file_path = '%s.xlsx' % self.file_path
                    break

            self.job_log_item("Writing {0} files into '{1}'".format(len(self.job_files), self.file_path))
            with pd.ExcelWriter(self.file_path) as writer:
                for file in self.job_files:
                    if not file[1]:
                        file[1] = file[0]

                    file[2].to_excel(writer, index=False, sheet_name=file[1])

            self.job_files = []

    def check_lists(self):
        for sub_job, sub_job_type in copy.deepcopy(self.job_config['Job_List']):
            found_item = False

            for sj_name, sj_type in zip(self.sub_job_name, self.sub_job_type):
                if sj_name == os.path.basename(sub_job[0]) and sj_type == sub_job_type:
                    found_item = True
                    break

            if not found_item:
                self.sub_job_name.append(os.path.basename(sub_job[0]))
                self.sub_job_type.append(sub_job_type)
                self.sub_start_time.append(None)
                self.job_log_item("{0} '{1}' failed [ECode 00x02] - This task never processed"
                                  .format(sub_job_type, os.path.basename(sub_job[0])))
                self.sub_end_time.append(None)
                self.sub_error.append('00x02')

    def check_error_list(self):
        failed_tasks = 0

        for error in self.sub_error:
            if error:
                failed_tasks += 1

        if failed_tasks > 0:
            return "Failed {0} task(s) during execution".format(failed_tasks)
        else:
            return None

    def list_chksum(self, err_code='00x02'):
        while len(self.sub_end_time) < len(self.sub_start_time):
            self.sub_end_time.append(None)

        eline = len(self.sub_error)

        while len(self.sub_error) < len(self.sub_start_time):
            self.sub_error.append(err_code)

            if err_code == '00x02':
                self.job_log_item("{0} '{1}' failed [ECode 00x02] - This task never processed"
                                  .format(self.sub_job_type[eline], self.sub_job_name[eline]))
                eline += 1

    def send_email(self, error_msg=None):
        email_trys = -1
        self.list_chksum()
        self.check_lists()

        if not error_msg:
            error_msg = self.check_error_list()

        if error_msg:
            self.job_log_item("Job '{0}' Failed job execution. Sending e-mail".format(self.job_config['Job_Name']))
        else:
            self.job_log_item("Job '{0}' Completed successfully. Sending e-mail".format(self.job_config['Job_Name']))

        package = zip(self.sub_job_name, self.sub_job_type, self.sub_start_time, self.sub_end_time, self.sub_error)
        obj = Email(job_config=self.job_config, job_results=package, attach=self.file_path, error_msg=error_msg)
        obj.email_connect()

        while email_trys < 4:
            email_trys += 1

            try:
                obj.package_email()
                obj.email_send()
                self.job_log_item("Email has been successfully sent")
                email_trys = 5
            except Exception as e:
                self.job_log_item("Job '{0}' failed e-mail sending [ECode {1}] - {2}"
                                  .format(self.job_config['Job_Name'], type(e).__name__, str(e)))
                global_objs['Event_Log'].write_log(traceback.format_exc(), 'critical')
                sleep(5)
                pass
            finally:
                obj.email_close()
                del obj

    def close_conn(self):
        if self.asql:
            self.asql.close_conn()


def check_settings():
    if not os.path.exists(batcheddir):
        os.makedirs(batcheddir)

    if not os.path.exists(joblogsdir):
        os.makedirs(joblogsdir)

    if not global_objs['Settings'].grab_item('Server') \
            or not global_objs['Settings'].grab_item('Database') \
            or not global_objs['Settings'].grab_item('Email_Server') \
            or not global_objs['Settings'].grab_item('Email_Port') \
            or not global_objs['Settings'].grab_item('Email_User') \
            or not global_objs['Settings'].grab_item('Email_Pass'):
        return False
    else:
        return True


def repackage_freq(my_config, freq_line, next_run):
    job_schedule = copy.deepcopy(my_config['Job_Schedule'])
    freq_list, days_of_week, freq_start_dt_list, freq_missed_run_list, prev_run_list, next_run_list = zip(*job_schedule)
    prev_run_list = list(prev_run_list)
    next_run_list = list(next_run_list)
    prev_run_list[freq_line] = datetime.datetime.now()
    next_run_list[freq_line] = next_run
    my_config['Job_Schedule'] = zip(freq_list, days_of_week, freq_start_dt_list, freq_missed_run_list, prev_run_list,
                                    next_run_list)
    my_config['Job_Controls'] = [True, True, 0]
    return my_config


def val_exec(my_config, sstarted):
    freq_line = -1
    job_schedule = copy.deepcopy(my_config['Job_Schedule'])
    job_controls = my_config['Job_Controls']

    if job_controls[0]:
        for freq, days_of_week, freq_start_dt, freq_missed_run, prev_run, next_run in job_schedule:
            freq_line += 1

            if datetime.datetime.now().strftime('%Y%m%d %H:%M') == next_run.strftime('%Y%m%d %H:%M') \
                    or (freq_missed_run == 1 and next_run < datetime.datetime.now()) \
                    or (not job_controls[1] and job_controls[2] == 2) or (sstarted and job_controls[1]):
                return repackage_freq(my_config, freq_line, next_run_date(freq_start_dt, freq, days_of_week))

    if job_controls[2] == 1:
        stop_job_list.append(my_config['Job_Name'].lower())


def exec_job(class_obj):
    if class_obj:
        try:
            class_obj.start_job()
            class_obj.export_files()
            class_obj.send_email()
        finally:
            class_obj.close_job()
            class_obj.close_conn()
            del class_obj


def exec_email(class_obj, err_msg):
    if class_obj:
        try:
            class_obj.export_files()
            class_obj.send_email(error_msg=err_msg)
        finally:
            class_obj.close_job()
            class_obj.close_conn()
            del class_obj


def watch_jobs(job_thread, job_obj, job_timeout):
    if len(stop_job_list) > 0:
        for stop_job_name in stop_job_list:
            if stop_job_name == job_obj.job_name().lower():
                global_objs['Event_Log'].write_log("Job '{0}' was forcebly requested to be stopped by GUI"
                                                   .format(job_obj.job_name()))

                if job_thread.is_alive():
                    job_thread.terminate()

                jw = Process(target=exec_email, args=[job_obj,
                                                      "Failed execution because it was requested to be stopped by GUI"])
                jw.daemon = True
                jw.start()
                stop_job_list.remove(stop_job_name)
                return True

    if job_thread.exitcode is None and not job_thread.is_alive():
        global_objs['Event_Log'].write_log("Job '{0}' failed execution as if it was never executed"
                                           .format(job_obj.job_name()))
        jw = Process(target=exec_email, args=[job_obj, "Failed execution as if it was never executed"])
        jw.daemon = True
        jw.start()
        return True
    elif job_thread.exitcode == 0 and not job_thread.is_alive():
        global_objs['Event_Log'].write_log("Job '{0}' Completed!".format(job_obj.job_name()))
        return True
    elif job_thread.exitcode is not None:
        global_objs['Event_Log'].write_log("Job '{0}' failed execution because it ran into Error Code {1}"
                                           .format(job_obj.job_name(), job_thread.exitcode))

        if job_thread.is_alive():
            job_thread.terminate()

        jw = Process(target=exec_email, args=[job_obj, "Failed execution because it ran into an Error Code %s"
                                              % job_thread.exitcode])
        jw.daemon = True
        jw.start()
        return True
    elif job_timeout and job_timeout < datetime.datetime.now():
        global_objs['Event_Log'].write_log("Failed execution because of Time-Out!")

        if job_thread.is_alive():
            job_thread.terminate()

        jw = Process(target=exec_email, args=[job_obj, "Timed-Out while processing"])
        jw.daemon = True
        jw.start()
        return True


def attach_cleanup():
    for file in list(pl.Path(batcheddir).glob('*.xlsx')):
        os.remove(file)


def exit_handler(jobs_to_kill):
    if len(jobs_to_kill) > 0:
        for job_item in jobs_to_kill:
            global_objs['Event_Log'].write_log("Job '{0}' was forcebly requested to be stopped by program exit"
                                               .format(job_item[1].job_name()))
            job_item[1].job_log_item("Failed execution because it was requested to be stopped by program exit")

            if job_item[0].is_alive():
                job_item[0].terminate()

            job_item[1].close_conn()

    attach_cleanup()
    sys.exit(0)


if __name__ == '__main__':
    jobs = []
    atexit.register(exit_handler, jobs)
    attach_cleanup()

    if getattr(sys, 'frozen', False):
        freeze_support()

    global_objs['Event_Log'].write_log("Initializing Job Scheduler Settings")

    if check_settings():
        script_started = True
        jw_thread = None
        stop_job_list = []

        try:
            BaseManager.register('JobConfig', JobConfig)
            manager = BaseManager()
            manager.start()
            global_objs['Event_Log'].write_log("Job Scheduler is now sniffing for Jobs to execute")

            while 1 != 0:
                global_objs['Local_Settings'].read_shelf()
                configs = global_objs['Local_Settings'].grab_item('Job_Configs')

                if configs:
                    update_settings = False

                    for line, config in enumerate(configs):
                        new_config = val_exec(config, script_started)

                        if new_config:
                            global_objs['Event_Log'].write_log("Starting Job '%s'" % config['Job_Name'])
                            update_settings = True

                            if int(new_config['Job_Timeout'][0]) > 0 or int(new_config['Job_Timeout'][1]) > 0:
                                timeout = datetime.datetime.now() + datetime.timedelta(
                                    hours=int(new_config['Job_Timeout'][0]), minutes=int(new_config['Job_Timeout'][1]))
                            else:
                                timeout = None

                            new_config['Start_Time'] = datetime.datetime.now()
                            myobj = manager.JobConfig(new_config, timeout)
                            configs[line] = new_config
                            j = Process(target=exec_job, args=[myobj])
                            j.daemon = True
                            j.start()
                            jw_thread = [j, myobj, timeout]
                            jobs.append(jw_thread)

                    if script_started:
                        script_started = False

                    if update_settings:
                        add_setting('Local_Settings', configs, 'Job_Configs', False)

                    if len(jobs) > 0:
                        for my_job in jobs:
                            if watch_jobs(my_job[0], my_job[1], my_job[2]):
                                jobs.remove(my_job)

                sleep(10)
        except:
            global_objs['Event_Log'].write_log(traceback.format_exc(), 'critical')
        finally:
            exit_handler(jobs)
    else:
        global_objs['Event_Log'].write_log(
            'Settings haven''t been established. Please run Job_Scheduler_Settings', 'error')
