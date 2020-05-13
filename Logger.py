# -*- coding: utf-8 -*-
""" Script by Toybich Egor
"""
from __future__ import print_function
from datetime import datetime
from datetime import timedelta

__version__ = '1.0.9'

class Logger(object):
    """
    A class to write into a log file
    
    Args:
        filename: str; name of a log file
        alwaysnew: bool; always start a new log file if a new logger is created with
                         the same log file
        info: str; add this strs at the beginning of each message
        console: bool; duplicate output to stdout 
    """
    __version__ = '1.0.9'
    
    __files_in_use = set()
    
    __total_instances = 0
    
    # ---------------------------------------------------------------	
    # Public attributes
    # ---------------------------------------------------------------
    @property
    def total_instances(self):
        """How many loggers of this class were opened"""
        return self.__total_instances
    
    @property
    def files_in_use(self):
        """Log files used by this logger class"""
        return self.__files_in_use
    
    @property
    def filename(self):
        """Name of a log file for this class instance"""
        return self.__filename
        
    # ---------------------------------------------------------------		
    # Magic methods
    # ---------------------------------------------------------------	
    def __init__(self, filename='log.txt', alwaysnew=False, info=None, console=True):
        self.__info = info
        self._console = console

        if alwaysnew or filename not in self.__class__.__files_in_use:
            with open(filename, 'w'): pass
            
        self.__filename = filename
        
        self.__files_in_use.add(filename)
        self.__total_instances += 1
        self.__logger_instance = self.__total_instances
        self.__init_time = datetime.now()
        
        self._log_('New Logger instance created ({})'.format(self.__total_instances), 1)
        self._log_('Class version: ' + self.__version__ )
        
        
    def __repr__(self):
        return self.__filename
    
    # ---------------------------------------------------------------		
    # Public methods
    # ---------------------------------------------------------------	
    def runtime(self):
        """Prints a total time since creation of this object"""
        t_run = datetime.now() - self.__init_time
        t_run = timedelta(days=t_run.days, seconds=t_run.seconds, microseconds=0)
        self._log_('Runtime: {}'.format(t_run), 1)
    
    def blank(self):
        """Prints an empty line in a log file"""
        with open(self.__filename, 'a') as log_file:
            log_file.write('\n')	   
    
    def log(self, msg_log, newln=0, info=''):
        """
        Use this method to write into log file
        
        Args:
            msg_log: str; log message
            newln: int; if 1 - add an empty line after message, 
                        if 2 - print raw message in a log file
            info: add this at message at the beginning of each line; default to self.info 
        """
        info_msg = info if self.__info is None or info != '' else self.__info   
        info_msg = info_msg + '|| ' if info_msg else ''
              
        msg_time='[{}]'.format(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        msg_actual = str(msg_log).splitlines() if newln != 2 else str(msg_log)
        with open(self.__filename, 'a') as log_file:
            if newln == 2:
                log_file.write(msg_actual)
                if self._console: print(msg_actual, end='')
            else: 
                for msg in msg_actual:
                    mf = '({}) {} {}{}\n'.format(self.__logger_instance, msg_time, info_msg, msg)
                    log_file.write(mf)
                    if self._console: print('{} {}{}'.format(msg_time, info_msg, msg))
                if newln == 1: log_file.write('\n')
                
    def _log_(self, msg, newline=0):
        """Prints message with info of Logger class"""
        self.log(msg, newline, info = str(self.__class__.__name__))
    # ---------------------------------------------------------------		
    # Static methods
    # ---------------------------------------------------------------