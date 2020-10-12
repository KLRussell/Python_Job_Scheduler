from New_Job_Scheduler_Settings import JobList

import sys


if __name__ == '__main__':
    if getattr(sys, 'frozen', False):
        from multiprocessing import freeze_support

        freeze_support()

    JobList()
