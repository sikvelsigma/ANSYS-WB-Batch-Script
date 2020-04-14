# -*- coding: utf-8 -*-
""" Script by Toybich Egor
"""
from datetime import datetime

__version__ = '1.0.7'

class Logger(object):
	"""
	A class to write into a log file
	"""
	__version__ = '1.0.7'
	
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
	def __init__(self, filename='log.txt', alwaysnew=False, info=None):
		"""If alwaysnew=True will overwrite already opened log files """
		self.__info = info
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
		
	def __repr__(self):
		return self.__filename
	
	# ---------------------------------------------------------------		
	# Public methods
	# ---------------------------------------------------------------	
	def blank(self):
		with open(self.__filename, 'a') as log_file:
			log_file.write('\n')	
	
	def log(self, msg_log, newln=0, info=''):
		"""Use this method to write into log file"""
		info_msg = info if self.__info is None or info != '' else self.__info
		
		msg_time='[{}]'.format(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
		msg_by_line = str(msg_log).splitlines()
		with open(self.__filename, 'a') as log_file:
			for msg in msg_by_line:
				mf = '({}) {} {} {}\n'.format(self.__logger_instance, msg_time, info_msg, msg)
				log_file.write(mf)
			if newln == 1: log_file.write('\n')
	# ---------------------------------------------------------------		
	# Static methods
	# ---------------------------------------------------------------