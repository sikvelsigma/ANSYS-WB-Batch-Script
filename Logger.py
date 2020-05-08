# -*- coding: utf-8 -*-
""" Script by Toybich Egor
"""
from __future__ import print_function
from datetime import datetime
from datetime import timedelta

__version__ = '1.0.8'

class Logger(object):
    """
    A class to write into a log file
    """
    __version__ = '1.0.8'
    
    __files_in_use = set()
    
    __total_instances = 0
    
    # ---------------------------------------------------------------	
    # Public attributes
    # ---------------------------------------------------------------
    @property
    def total_instances(self):
        """How many loggers of this class were opened"""
        return self.__class__.__total_instances
    
    @property
    def files_in_use(self):
        """Log files used by this logger class"""
        return self.__class__.__files_in_use
    
    @property
    def filename(self):
        """Name of log file"""
        return self.__filename
        
    # ---------------------------------------------------------------		
    # Magic methods
    # ---------------------------------------------------------------	
    def __init__(self, filename='log.txt', alwaysnew=False, info=None, console=True):
        """If alwaysnew=True will overwrite already opened log files """
        self.__info = info
        self._console = console
        key = 'w'
        add_s = ''
        if not alwaysnew and filename in self.__class__.__files_in_use:
            key = 'a'
            add_s = '\n'
        self.__filename = filename
        
        with open(filename, key) as log_file:
            log_file.write('{}NEW LOGGER OPENED ({}); VERSION: {}\n\n'.format(add_s, self.total_instances+1, self.__version__))
            
        self.__class__.__files_in_use.add(filename)
        self.__class__.__total_instances += 1
        self.__logger_instance = self.__class__.__total_instances
        self.__init_time = datetime.now()
    def __repr__(self):
        return self.__filename
    
    # ---------------------------------------------------------------		
    # Public methods
    # ---------------------------------------------------------------	
    def runtime(self):
        t_run = datetime.now() - self.__init_time
        t_run = timedelta(days=t_run.days, seconds=t_run.seconds, microseconds=0)
        self.log('Runtime: {}'.format(t_run), 1)
    
    def blank(self):
        with open(self.__filename, 'a') as log_file:
            log_file.write('\n')	
    
    def log(self, msg_log, newln=0, info=''):
        """Use this method to write into log file"""
        info_msg = info if self.__info is None or info != '' else self.__info   
        info_msg = info_msg + ' ' if info_msg else ''
        
        
        msg_time='[{}]'.format(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        msg_by_line = str(msg_log).splitlines()
        with open(self.__filename, 'a') as log_file:
            for msg in msg_by_line:
                mf = '({}) {} {}{}\n'.format(self.__logger_instance, msg_time, info_msg, msg)
                log_file.write(mf)
                if self._console: print('{} {}{}'.format(msg_time, info_msg, msg))
            if newln == 1: log_file.write('\n')
        
    # ---------------------------------------------------------------		
    # Static methods
    # ---------------------------------------------------------------