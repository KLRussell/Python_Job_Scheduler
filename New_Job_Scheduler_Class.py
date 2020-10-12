from New_Job_Scheduler_Settings import tool, local_config, job_logs_dir, email_engine, temp_dir, attach_dir
from datetime import datetime, timedelta, time
from KGlobal.data import DataConfig
from subprocess import Popen, PIPE, STARTUPINFO, STARTF_USESHOWWINDOW
from os import makedirs, remove, listdir, unlink, stat, kill
from os.path import splitext, join, exists, basename, islink, isdir, isfile
from portalocker import Lock
from pandas import ExcelWriter
from traceback import format_exc
from copy import deepcopy
from threading import Thread

startupinfo = STARTUPINFO()
startupinfo.dwFlags |= STARTF_USESHOWWINDOW


class Email(object):
    __job_profile = None

    def __init__(self, job):
        self.__attachments = list()
        self.__task_profiles = list()
        self.__job_start = datetime.now()
        self.__job_profile = job
        self.__job_log = DataConfig(file_dir=job_logs_dir, file_name_prefix=self.__job_profile['Job_Name'].lower(),
                                    file_ext='log')
        self.__email = None
        self.__err_msg = None

        for task in self.__job_profile['Tasks']:
            self.__copy_task(task)

        if 'To_Email' in self.__job_profile.keys() and self.__job_profile['To_Email']:
            self.__to_email = self.__job_profile['To_Email'].replace(' ', '').replace(';', ',').split(',')
        else:
            self.__to_email = None

        if 'Cc_Email' in self.__job_profile.keys() and self.__job_profile['Cc_Email']:
            self.__cc_email = self.__job_profile['Cc_Email'].replace(' ', '').replace(';', ',').split(',')
        else:
            self.__cc_email = None

    @property
    def job_name(self):
        if self.__job_profile and self.__job_profile['Job_Name']:
            return self.__job_profile['Job_Name']

    @property
    def job_start(self):
        if self.__job_start is None:
            return datetime.now()
        else:
            return self.__job_start

    @property
    def task_profiles(self):
        return self.__task_profiles

    @task_profiles.setter
    def task_profiles(self, task_profile):
        self.__task_profiles = task_profile

    def write_job_log(self, message):
        log = self.__job_log[datetime.today().__format__("%Y%m%d")]

        if not log:
            log = list()

        log.append([datetime.now().__format__("%I:%M:%S %p"), message])
        self.__job_log[datetime.today().__format__("%Y%m%d")] = log
        self.__job_log.sync()

    def email_results(self, err_msg=None):
        if self.__job_profile and self.__task_profiles:
            try:
                self.write_job_log("Job '%s' is sending e-mail to distros" % self.__job_profile['Job_Name'])
                self.__err_msg = err_msg
                self.__package_tasks()
                self.__package_email()
                self.__package_attach()
                self.__email.send()
                self.write_job_log("Job '%s' is e-mail was sent successfully" % self.__job_profile['Job_Name'])
            except Exception as e:
                tool.write_to_log(format_exc(), 'critical')
                tool.write_to_log("Email - ECode '{0}', {1}".format(type(e).__name__, str(e)))

    def close_email(self):
        self.__email = None

    def __package_tasks(self):
        if self.__task_profiles:
            for task_profile in self.__task_profiles:
                if not task_profile['Task_End'] and not task_profile['Task_Error']:
                    if task_profile['Task_Type'] == 1:
                        task_type = 'Stored Proc'
                    else:
                        task_type = 'Program'

                    task_profile['Task_Error'] = ['x01', 'Task never processed']
                    self.write_job_log("{0} '{1}' failed [ECode {2}] - {3}".format(
                        task_type, task_profile['Task_Name'], task_profile['Task_Error'][0],
                        task_profile['Task_Error'][1]))

                    task_profile['Task_End'] = datetime.now()
                elif not task_profile['Task_End']:
                    task_profile['Task_End'] = datetime.now()

                if not task_profile['Task_Start']:
                    task_profile['Task_Start'] = task_profile['Task_End']

                if task_profile['Task_Error'] and not self.__err_msg:
                    self.__err_msg = 'failed execution because one or more tasks failed execution'

                if task_profile['Task_Attach']:
                    task_attach = list()

                    for attach in task_profile['Task_Attach']:
                        if not attach.empty:
                            task_attach.append(attach)

                    if task_attach:
                        self.__parse_attach(task_profile, task_attach)

    def __package_attach(self):
        if self.__attachments:
            from zipfile import ZipFile
            from exchangelib import FileAttachment
            from portalocker import Lock

            file_dir = join(attach_dir, datetime.today().__format__("%Y%m%d"))
            file_path = join(file_dir, '{0}_{1}.zip'.format(hash(datetime.now().__format__("%I:%M:%S %p")),
                                                            self.__job_profile['Job_Name']))

            if not exists(file_dir):
                makedirs(file_dir)

            if exists(file_path):
                try:
                    remove(file_path)
                except Exception as e:
                    tool.write_to_log("Attach File Rem - ECode '{0}', {1}".format(type(e).__name__, str(e)))
                    return

            zip_file = ZipFile(file_path, mode='w')

            try:
                for attachment in self.__attachments:
                    try:
                        if isfile(attachment) and exists(attachment) and stat(attachment).st_size > 0:
                            zip_file.write(attachment, basename(attachment))
                    except Exception as e:
                        self.write_job_log(format_exc())
                        tool.write_to_log("Job '{0}' failed to zip file due to ECode '{1}', {2}"
                                          .format(self.__job_profile['Job_Name'], type(e).__name__, str(e)))
                        pass
                    finally:
                        if isfile(attachment) and exists(attachment) and stat(attachment).st_size > 0:
                            remove(attachment)
            except Exception as e:
                self.write_job_log(format_exc())
                tool.write_to_log("Job '{0}' failed to zip file in mainloop due to ECode '{1}', {2}".format(
                    self.__job_profile['Job_Name'], type(e).__name__, str(e)))
                pass
            finally:
                zip_file.close()

            if exists(file_path) and stat(file_path).st_size > 0:
                with Lock(file_path, 'rb') as f:
                    self.__email.attach(FileAttachment(name=basename(file_path), content_type='zip', content=f.read(),
                                                       is_inline=False))
            elif exists(file_path):
                remove(file_path)

                try:
                    unlink(file_dir)
                except Exception as e:
                    tool.write_to_log("Dir Unlink - ECode '{0}', {1}".format(type(e).__name__, str(e)))
                    pass

    def __package_email(self):
        if self.__to_email and self.__job_profile['Job_Name']:
            from exchangelib import Message

            body = list()
            self.__email = Message(account=email_engine)
            names = [str(email.split('@')[0]).title() for email in self.__to_email]

            if self.__to_email:
                self.__email.to_recipients = self.__gen_email_list(self.__to_email)

            if self.__cc_email:
                self.__email.cc_recipients = self.__gen_email_list(self.__cc_email)

            body.append('Happy {0} {1},\n'.format(datetime.today().strftime("%A"), '; '.join(names)))

            if self.__err_msg:
                tool.write_to_log("Job '{0}' failed job execution due to {1}".format(
                    self.__job_profile['Job_Name'], self.__err_msg))
                self.__email.subject = "<Job Failed> Job \"%s\"" % self.__job_profile['Job_Name']
                sub_body = "Job \"{0}\", {1}".format(self.__job_profile['Job_Name'], self.__err_msg)
            else:
                tool.write_to_log("Job '%s' completed successfully" % self.__job_profile['Job_Name'])
                self.__email.subject = "Job \"%s\" completed successfully" % self.__job_profile['Job_Name']
                sub_body = "Job \"{0}\" completed successfully".format(self.__job_profile['Job_Name'])

            body.append("{0}. Total job runtime was {1}.\n".format(sub_body, parse_time(self.job_start,
                                                                                        datetime.now())))

            for task_profile in self.__task_profiles:
                if task_profile['Task_Type'] == 1:
                    task_type = 'Stored Proc'
                else:
                    task_type = 'Program'

                if task_profile['Task_Error']:
                    body.append('\t\u2022  {0} "{1}" <Failed Task> ({2}) [Err Code: {3}]'.format(
                        task_type, task_profile['Task_Name'],
                        parse_time(task_profile['Task_Start'], task_profile['Task_End']),
                        task_profile['Task_Error'][0]))
                else:
                    body.append('\t\u2022  {0} "{1}" <Succeeded Task> ({2})'.format(
                        task_type, task_profile['Task_Name'], parse_time(task_profile['Task_Start'],
                                                                         task_profile['Task_End'])))

            body.append("\nYour's Truly,\n")
            body.append("The BI Team")
            self.__email.body = '\n'.join(body)

    @staticmethod
    def __gen_email_list(emails):
        from exchangelib import Mailbox
        objs = list()

        for email in emails:
            objs.append(Mailbox(email_address=email))

        return objs

    def __parse_attach(self, task_profile, task_attach):
        file_dir = join(temp_dir, self.__job_profile['Job_Name'])
        file_path = join(file_dir, '{0}_{1}.xlsx'.format(hash(datetime.now().__format__("%I:%M:%S %p")),
                                                         task_profile['Task_Name']))

        remove_dir(file_dir)

        if not exists(file_dir):
            makedirs(file_dir)

        with ExcelWriter(file_path) as w:
            for num, df in zip(range(0, len(task_attach)), task_attach):
                if 'Tab_Name' in task_profile.keys() and not df.empty:
                    df.to_excel(w, index=False, sheet_name='{0}_{1}'.format(task_profile['Tab_Name'], num))
                elif not df.empty:
                    df.to_excel(w, index=False, sheet_name=str(num))

        if exists(file_path) and stat(file_path).st_size > 0:
            self.__attachments.append(file_path)

    def __copy_task(self, task):
        profile = deepcopy(task)
        profile['Task_Start'] = None
        profile['Task_End'] = None
        profile['Task_Error'] = None
        profile['Task_Attach'] = None
        self.__task_profiles.append(profile)


class Job(Email):
    def __init__(self, job):
        Email.__init__(self, job)
        self.__sql_engine = None
        self.__sub_proc = None
        self.__kill_thread = False
        self.__finished = False

    @property
    def thread_killed(self):
        return self.__kill_thread

    @property
    def tasks_finished(self):
        return self.__finished

    def job_info(self):
        return self.job_start

    def start_job(self):
        self.__finished = False

        for task in self.task_profiles:
            if self.__kill_thread:
                break

            if task['Task_Type'] == 1:
                task_type = 'Stored Proc'
            else:
                task_type = 'Program'

            try:
                self.write_job_log("Task {0} '{1}' is processing".format(task_type, task['Task_Name']))
                task['Task_Start'] = datetime.now()

                if task['Task_Type'] == 1:
                    self.__sql_execute(task, 'EXEC %s' % task['Task'])
                else:
                    self.__task_execute(task)
            except Exception as e:
                self.write_job_log(format_exc())
                task['Task_Error'] = [type(e).__name__, type(e).__name__, str(e)]
                pass
            finally:
                task['Task_End'] = datetime.now()

                if task['Task_Error']:
                    self.write_job_log("{0} '{1}' failed [ECode {2}] - {3}".format(
                        task_type, task['Task_Name'], task['Task_Error'][0], task['Task_Error'][1]))
                else:
                    self.write_job_log("Task {0} '{1}' finished".format(task_type, task['Task_Name']))

        self.__finished = True

    def terminate(self, kill_thread=False):
        from KGlobal.sql import SQLEngineClass
        self.__kill_thread = kill_thread

        if self.__sql_engine and isinstance(self.__sql_engine, SQLEngineClass):
            self.__sql_engine.close_connections(destroy_self=True)

        if self.__sub_proc and self.__sub_proc.poll() is None:
            tool.write_to_log("Job '{0}' - Attempting to kill process thread {1}".format(self.job_name,
                                                                                         self.__sub_proc.pid))

            try:
                kill(self.__sub_proc.pid, -9)
            except Exception as e:
                self.write_job_log(format_exc())
                tool.write_to_log("Job '{0}' kill ECode '{1}', {2}".format(self.job_name, type(e).__name__, str(e)))
                pass

        self.__sql_engine = None
        self.__sub_proc = None

    def __task_execute(self, task_profile):
        try:
            shell_line = None
            shell_comm = task_profile['Task_SComm']
            ext = splitext(task_profile['Task'])[1].lower()

            if shell_comm:
                shell_comm = ". '%s'" % shell_comm
            elif ext == '.py':
                shell_comm = 'python'
            elif ext == '.ps1':
                shell_comm = '.'
            elif ext == '.vbs':
                shell_comm = 'cscript'

            if task_profile['Params']:
                params = task_profile['Params']
            else:
                params = ''

            if ext == '.sql':
                with Lock(task_profile['Task'], 'r') as f:
                    query = f.read()

                self.__sql_execute(task_profile, query)
            elif ext == '.exe':
                shell_line = "'{0}' {1}".format(str(task_profile['Task']), params)
            elif shell_comm:
                shell_line = ['powershell.exe', "{0} '{1}' {2}".format(shell_comm, task_profile['Task'], params)]

            if shell_line:
                self.__sub_proc = Popen(shell_line, startupinfo=startupinfo, stdin=PIPE, stdout=PIPE, stderr=PIPE)
                stdout, stderr = self.__sub_proc.communicate()
                self.__process_std_output(task_profile['Task_Name'], stdout)
                self.__process_std_output(task_profile['Task_Name'], stderr, is_error=True)

                if self.__sub_proc.returncode:
                    task_profile['Task_Error'] = [self.__sub_proc.returncode, 'Ran into error while executing program']
        except Exception as e:
            self.write_job_log(format_exc())
            task_profile['Task_Error'] = [type(e).__name__, str(e)]
            pass

    def __sql_execute(self, task_profile, query_str):
        if not self.__sql_engine:
            self.__sql_engine = tool.default_sql_conn(new_instance=True)

        self.__sql_engine.sql_execute(query_str=query_str, execute=True, queue_cursor=True)
        sql_results = self.__sql_engine.wait_for_cursors()

        if sql_results and sql_results[0]:
            if hasattr(sql_results[0], 'results'):
                task_profile['Task_Attach'] = sql_results[0].results

            if hasattr(sql_results[0], 'errors'):
                task_profile['Task_Error'] = sql_results[0].errors

    def __process_std_output(self, task_name, std_output, is_error=False):
        if std_output:
            for line in std_output.decode("utf-8").split('\n'):
                if is_error:
                    self.write_job_log('[ERR] {0} => {1}'.format(task_name, line))
                else:
                    self.write_job_log('{0} => {1}'.format(task_name, line))


class JobScheduler(object):
    __watch_thread = None
    __running = False
    __jobs = dict()

    def __init__(self, parent):
        tool.write_to_log("Initializing Job Scheduler")
        self.__parent = parent

    def start(self):
        tool.write_to_log("Starting Job Scheduler Thread")
        from time import sleep
        self.__running = True
        self.__jobs = dict()

        try:
            initiated = True

            while self.__running:
                job_names = local_config['Jobs'].keys()

                for job_name in job_names:
                    if job_name in local_config['Jobs'].keys():
                        job_profile = local_config['Jobs'][job_name]
                        self.__job_validate(job_profile, initiated)
                        self.__job_watch(job_profile)

                initiated = False
                sleep(1)
        except Exception as e:
            tool.write_to_log(format_exc(), 'critical')
            tool.write_to_log("Job Main - ECode '{0}', {1}".format(type(e).__name__, str(e)))
        finally:
            self.__kill_all_jobs()
            tool.write_to_log("Job Scheduler Thread has been stopped")

    def info(self):
        info = list()

        if self.__jobs:
            for job_name, job in self.__jobs.items():
                info.append("Job '{0}' was started {1}".format(job_name, parse_time(job[1].job_info(), datetime.now())))

        return info

    def exit(self):
        tool.write_to_log("Exiting Job Scheduler Thread")
        self.__running = False

        if self.__parent and hasattr(self.__parent, 'js_thread'):
            self.__parent.js_thread.join()
            self.__parent.js_thread = None

    def __job_validate(self, job_profile, initiated=False):
        if job_profile['Enabled'] and ((initiated and (job_profile['Running'] or job_profile['Manual_Start'])) or
                                       (not job_profile['Running'] and not job_profile['Manual_Stop']
                                        and job_profile['Job_Name'] not in self.__jobs and
                                        (job_profile['Manual_Start'] or job_profile['Next_Run'].strftime('%Y%m%d %H:%M')
                                         <= datetime.now().strftime('%Y%m%d %H:%M')))):
            if (job_profile['Manual_Start'] or job_profile['Running']) and\
                    job_profile['Next_Run'].strftime('%Y%m%d') == datetime.today().strftime('%Y%m%d'):
                job_profile['Next_Run'] = get_next_run(job_profile=job_profile, skip_today=True)
            elif not job_profile['Manual_Start']:
                job_profile['Next_Run'] = get_next_run(job_profile=job_profile)

            job_profile['Manual_Start'] = False
            job_profile['Running'] = True
            job_profile['Prev_Run'] = datetime.now()
            self.__save_profile(job_profile)
            self.__start_job(job_profile)

    def __start_job(self, job_profile):
        tool.write_to_log("Job '{0}' is now starting".format(job_profile['Job_Name']))
        job_class = Job(deepcopy(job_profile))
        timeout = get_timeout(hours=job_profile['Timeout_HH'], minutes=job_profile['Timeout_MM'])
        thread = Thread(target=process_job, args=(job_class,))
        thread.daemon = True
        thread.start()
        self.__jobs[job_profile['Job_Name']] = [thread, job_class, timeout, self]

    @staticmethod
    def __kill_job(job, job_name, reason):
        if job and job_name and reason and not job[1].tasks_finished:
            message = "failed execution because %s" % reason

            try:
                if job[0].is_alive():
                    job[1].terminate(kill_thread=True)
                    job[0].join()

                job[1].email_results(message)
            except Exception as e:
                tool.write_to_log(format_exc(), 'critical')
                tool.write_to_log("Kill Job - ECode '{0}', {1}".format(type(e).__name__, str(e)))
                pass
            finally:
                job[1].close_email()

    def __kill_all_jobs(self):
        if self.__jobs:
            reason = 'the job scheduler was shut off'

            for job_name, job in self.__jobs.items():
                if job[0].is_alive():
                    self.__kill_job(job, job_name, reason)

            self.__jobs = dict()

    def __job_watch(self, job_profile):
        reason = None
        config_fix = False
        job = None

        if job_profile['Enabled'] and job_profile['Manual_Stop']:
            if job_profile['Job_Name'] in self.__jobs.keys():
                reason = 'it was manually stopped by GUI'
                job = self.__jobs[job_profile['Job_Name']]

            config_fix = True
        elif job_profile['Job_Name'] in self.__jobs.keys():
            job = self.__jobs[job_profile['Job_Name']]

            if not job[0].is_alive():
                config_fix = True
            elif job[2] < datetime.now():
                reason = 'it timed-out while processing'
                config_fix = True
            else:
                job = None

        if reason:
            thread = Thread(target=self.__kill_job, args=(job, job_profile['Job_Name'], reason))
            Thread.daemon = True
            thread.start()

        if job:
            del self.__jobs[job_profile['Job_Name']]

        if config_fix:
            job_profile['Manual_Start'] = False
            job_profile['Manual_Stop'] = False
            job_profile['Running'] = False
            self.__save_profile(job_profile)

    @staticmethod
    def __save_profile(job_profile):
        jobs = local_config['Jobs']

        if job_profile and jobs[job_profile['Job_Name']]:
            jobs[job_profile['Job_Name']] = job_profile
            local_config['Jobs'] = jobs
            local_config.sync()


def date_add(interval, increment, date):
    now = datetime.now()

    if interval == 'day' and increment > 0:
        return date + timedelta(days=increment)
    elif interval == 'week' and increment > 0:
        return date_add('day', increment * 7, date)
    elif interval == 'month' and increment > 0:
        from dateutil.relativedelta import relativedelta
        y = now.year - date.year
        m = now.month - date.month
        return date + relativedelta(months=increment + (y * 12 + m))
    elif interval == 'year' and increment > 0:
        from dateutil.relativedelta import relativedelta
        y = now.year - date.year
        return date + relativedelta(years=increment + (y * 12))
    elif interval in ('day', 'week', 'month', 'year') and increment == 0:
        return date
    else:
        raise ValueError("Error! %r is an invalid interval" % interval)


def get_next_run(job_profile, skip_today=False):
    def next_run(schedule_profile):
        def dow_next(run_date, freq_dow):
            weekday = run_date.weekday()
            dow = [run_date + timedelta(days=n - weekday) for n in range(0, len(freq_dow)) if freq_dow[n] == 1
                   and (weekday < n or (weekday == n and run_date + timedelta(days=n - weekday) > datetime.now()))]
            dow += [run_date + timedelta(days=7 - (weekday - n)) for n in range(0, len(freq_dow)) if freq_dow[n] == 1
                    and (weekday > n or (weekday == n and run_date + timedelta(days=n - weekday) <= datetime.now()))]
            dow.sort()
            return dow[0]

        if skip_today:
            now = date_add('day', 1, datetime.now())
            start_datetime = datetime.combine(datetime.date(now), time(hour=schedule_profile['Start_HH'],
                                                                       minute=schedule_profile['Start_MM']))
        else:
            now = datetime.now()
            start_datetime = datetime.combine(datetime.date(now), time(hour=schedule_profile['Start_HH'],
                                                                       minute=schedule_profile['Start_MM']))

        if schedule_profile['Start_Datetime'] > now:
            return schedule_profile['Start_Datetime']
        elif schedule_profile['Frequency'] == 0:
            return date_add('day', 1, start_datetime)
        elif schedule_profile['Frequency'] == 1:
            return dow_next(start_datetime, schedule_profile['Frequency_DOW'])
        elif schedule_profile['Frequency'] == 2:
            from math import floor
            next_date = dow_next(start_datetime, schedule_profile['Frequency_DOW'])
            offset = floor((next_date - schedule_profile['Start_Datetime']).days / 7) % 2
            return date_add('week', offset, next_date)
        elif schedule_profile['Frequency'] == 3:
            return date_add('month', 1, start_datetime)

    if job_profile and job_profile['Schedules']:
        next_run_list = list()

        for schedule in job_profile['Schedules']:
            next_run_list.append(next_run(schedule))

        next_run_list.sort()
        return next_run_list[0]


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

    if len(date_list) < 1:
        date_list.append('0 Seconds')

    if len(date_list) == 1:
        return date_list[0]
    elif len(date_list) > 1:
        last = date_list[-1]
        date_list.remove(date_list[-1])
        return '{0} and {1}'.format(', '.join(date_list), last)
    else:
        return '0 Seconds'


def get_timeout(hours, minutes):
    return datetime.now() + timedelta(hours=hours, minutes=minutes)


def remove_dir(dir_filepath):
    if exists(dir_filepath):
        from shutil import rmtree

        for filename in listdir(dir_filepath):
            file_path = join(dir_filepath, filename)
            try:
                if isfile(file_path) or islink(file_path):
                    unlink(file_path)
                elif isdir(file_path):
                    rmtree(file_path)
            except Exception as e:
                tool.write_to_log('Failed to delete %s. Reason: %s' % (file_path, e), 'warning')
                pass


def process_job(job_instance):
    if job_instance:
        job_instance.start_job()

        if not job_instance.thread_killed:
            try:
                job_instance.email_results()
            finally:
                job_instance.terminate()
                job_instance.close_email()
