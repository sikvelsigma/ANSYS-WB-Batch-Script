# -*- coding: utf-8 -*-
""" Script by Toybich Egor
Note: Workbench uses IronPython (Python 2.7)!
"""
#__________________________________________________________
#__________________________________________________________
# import os
# import csv
from glob import glob
from functools import partial 
from collections import defaultdict

from csv import reader as csvreader
from csv import writer as csvwriter
from csv import QUOTE_MINIMAL

from glob import glob 

# Import global from main to access Workbench commands
import __main__ as workbench

# Import Logger module from working directory
def find_module(st_in):
	res = []
	stlist = st_in if isinstance(st_in, list) else [st_in]
	
	for st in stlist:
		try:
			srch = [f for f in glob('{}*.py'.format(st))]
			print('WBInterface| Found: {}'.format(srch))
			srch = srch[0] if srch[0] == '{}.py'.format(st) else srch[-1]
			srch = srch.replace('.py','')
		except:
			res.append(None)
		else:
			res.append(srch)
	return tuple(res) if len(stlist) > 1 else res[0]

log_module = find_module('Logger')
print('WBInterface| Using: {}'.format(log_module))
if log_module: exec('from {} import Logger'.format(log_module))

__version__ = '2.0.3'
#__________________________________________________________
class WBInterface(object):
	"""
	A class to open Workbench project/archive, input/output parameters
	and start calculations.
	"""
	__version__ = '2.0.3'
	
	# ---------------------------------------------------------------	
	# Public attributes
	# ---------------------------------------------------------------
	
	@property
	def DPs(self):
		"""Design Points list."""
		return self.__DPs
	
	@property
	def DPs_count(self):
		"""Design Points count (imported, present)."""
		return (self.__DPs_imported, self.__DPs_present)
	
	@property
	def filename(self):
		"""Opened project."""
		return self.__workfile	
		
	@property
	def out_file(self):
		"""Output file for parameters."""
		return self._out_file
	@out_file.setter
	def out_file(self, filename):
		"""Supports manual change of output file."""
		self._out_file = filename
		
	@property
	def full_file(self):
		"""Workbench project summary file."""
		return self._full_file
	@full_file.setter
	def full_file(self, filename):
		"""Supports manual change of summary file."""
		self._full_file = filename
	
	@property
	def logfile(self):
		"""Log file of script execution."""
		return self._logger.filename
		
	@property
	def isopen(self):
		"""Returns if there is an open Workbench project."""
		return self.__active
		
	@property
	def parameters(self):
		"""Returns IO parameters with values as dictionaries."""
		return (self._param_in_value, 	self._param_out_value)
			
	# ---------------------------------------------------------------		
	# Magic methods
	# ---------------------------------------------------------------
	
	def __init__(self, logger = None, out_file='output.txt', full_report_file='full_report.txt', 
				control_file_template='*_control.csv', input_file_template='*_input.csv', 
				csv_delim=',', csv_skip='no', loginfo=None):
		"""
		Constructor. All arguments have default values.
		
		Arg:
			logger (object with logger.log(str) method): Used for creating log file. Autocreates one if not defined.
			out_file (str): Output file name; defaults to 'output.txt'.
			full_report_file (str): Workbench parametric report file; defaults to 'full_report.txt'.
			control_file_template (str): String template for search of control file; defaults to '*_control.csv'.
			input_file_template (str): String template for search of input file; defaults to '*_input.csv'.
			csv_delim (str): Delimiter used in csv file; defaults to ','.
			csv_skip (str): Read as 'no parameters' if found; defaults to 'no'.
			loginfo (str): Prefix for logger to use; defaults to WBInterface
		
		Use method log() to write into a log file (see Logger class)
		"""
		self._logger = logger if logger is not None else Logger('log.txt')
		
		log_prefix = '{}||'.format(str(self.__class__.__name__)) if loginfo is None else loginfo
		
		# Partial logger method with prefix at the start of every line
		try:
			self._log_ = partial(self._logger.log, info=log_prefix)
		except:
			self._log_ = self._logger.log
		
		self.log = self._logger.log					#: original logger method
		self._out_file = out_file					#: file for outputting
		self._full_file = full_report_file			#: file for workbench parametric report
		self._control_file = None					#: csv file with io parameter list
		self._input_file = None						#: csv file with input parameters
		self._control_srch = control_file_template	#: key string to search for cotrol file
		self._input_srch = input_file_template		#: key string to search for input file
		self.__active = False						#: is a project open
		self._param_in = []							#: list of workbench input parameters
		self._param_out = []						#: list of workbench output parameters
		self._param_in_value = defaultdict(list)	#: dictionary with input parameters (keys=self._param_in)
		self._param_out_value = defaultdict(list)	#: dictionary with output parameters (keys=self._param_out)
		self._csv_delim = csv_delim					#: csv file delimiter
		self._csv_skip = csv_skip.lower()			#: string placeholder if no parameter is specified in csv file	
		self.__workfile = None						#: opened workbench project
		self.__DPs_imported = 0						#: Design Points imported from input file
		self.__DPs_present = 0						#: Design Points already present in project
		self.__DPs = None
		
		self.__control_srch_default = '*.control'	#: in-build key string to search for cotrol file
		self.__input_srch_default = '*.input'		#: in-build  key string to search for input file
		
		self._log_('Class version: ' + self.__version__ , 1)
		
	# ---------------------------------------------------------------		
	# Public methods
	# ---------------------------------------------------------------
	
	def read_control(self, control_file_template=None, csv_delim=None, csv_skip=None):
		"""
		Read csv file with IO parameter list.
		Example of control file format with 2 inputs and 3 outputs:
			#type 'no' to skip inputs
			p3,p2
			#type 'no' to skip ouputs
			p1,p4,p5
		Note that lines with '#' will be skipped.
		"""
		
		if control_file_template is None: control_file_template=self._control_srch
		if csv_delim is None: csv_delim=self._csv_delim
		if csv_skip is None: csv_skip=self._csv_skip
		
		if csv_delim is None or csv_skip is None:
			self._log_('Missing csv delimiter or skipper!', 1)
			raise MissingCSVParameter
			
		# if control_file_template is None:
			# self._log_('No control file defined!', 1)
			# raise FileSearchParameterNotDefined	
		
		for defiter in range(2):
			control_used = control_file_template if not defiter else self.__control_srch_default
			msg_mod = '' if not defiter else ' default'
			if control_used:
				self._log_('Searching for{} control file...'.format(msg_mod))
				try:
					file_list = [f for f in glob(control_used)]
					self._control_file = file_list[0]
				except:
					self._log_('Control file not found! No parameters will be used.', 1)
				else:
					self._log_('Control file found: ' + self._control_file)
					self._log_('Reading parameter list...')
					try:
						with open(self._control_file, 'r') as csvfile:
							spamreader = csvreader(self.decomment(csvfile), delimiter=csv_delim)
							for i, row in enumerate(spamreader):
								for rlin in row:
									skip = rlin.lower().find(csv_skip)
									if i == 0 and skip == -1:
										self._param_in.append(rlin.strip().upper())
									elif i == 1 and skip == -1:
										self._param_out.append(rlin.strip().upper())
					except Exception as err_msg:
						self._log_('An error occured wile reading control file!')
						self._log_(err_msg, 1)
						raise
					self._log_('Reading successful: ' + str(len(self._param_in)) + ' input(s), ' 
							 + str(len(self._param_out)) + ' output(s)', 1)
					break

			
		
	def read_input(self, input_file_template=None, csv_delim=None):
		"""
		Read csv file with input parameters.
		Example of format for 2 parameters and 3 Design Points:
			#p3,p2
			10,200
			30,400
			50,600
		Note that lines with '#' will be skipped.
		"""
		
		if input_file_template is None: input_file_template=self._input_srch
		if csv_delim is None: csv_delim=self._csv_delim
		
		if csv_delim is None:
			self._log_('Missing csv delimiter or skipper!', 1)
			raise MissingCSVParameter
		
		# if input_file_template is None:
			# self._log_('No input file defined!', 1)
			# raise FileSearchParameterNotDefined
		# elif self._control_srch is None:
			# self._log_('No control file defined!', 1)
			# raise FileSearchParameterNotDefined	
		if not self._param_in:
			self._log_('No parameters to input!', 1)
			return
		
		for defiter in range(2):
			input_used = input_file_template if not defiter else self.__input_srch_default
			msg_mod = '' if not defiter else ' default'
			if input_used:
		
				self._log_('Searching for{} input file...'.format(msg_mod))
				try:
					file_list = [f for f in glob(input_used)]
					self._input_file = file_list[0]
				except:
					self._log_('Input file not found!', 1)
				else:
					self._log_('Input file found: ' + self._input_file, 1) 
					self._log_('Reading input parameters...')
					try:
						with open(self._input_file, 'r') as csvfile:
							spamreader = csvreader(self.decomment(csvfile), delimiter=csv_delim)
							for row in spamreader:
								for key, elem in zip(self._param_in, row):
									self._param_in_value[key.upper()].append(elem.strip())
					except Exception as err_msg:
						self._log_('An error occured while reading input file!')
						self._log_(err_msg, 1)
						raise
					self.__DPs_imported = len(self._param_in_value[self._param_in[0]])
					self._log_('Reading successful: ' + str(self.__DPs_imported) + ' Design Point(s) found', 1)
					break
					
	def input_by_name(self, inp):
		"""
		Allows directly feed input parameters as list of dict
		Example of format for 2 parameters and 3 Design Points:
			list = [[1, 2, 3], [10, 20, 30]]
			dict = {'p1':[1, 2, 3], 'p2':[10, 20, 30]}
		"""
		self._log_('Direct data input by name issued')
		
		if not self.is_matrix(inp):
			self._log_('Incorrect input format!')
			raise ValueError('Incorrect input format!')
		
		if isinstance(inp, list):
			if not self._param_in:
				self._log_('No parameter keys found!')
				raise KeysNotFound
			self._input_list_by_name(inp)
		if isinstance(inp, dict):
			self._input_dict_by_name(inp)
		self._log_('Input successful: {} input(s) in {} Design Point(s)'.format(len(self._param_in),self.__DPs_imported), 1)	
	
	def input_by_DPs(self, inp, keys=None):
		"""
		Allows directly feed input parameters as list with list of keys.
		Example of format for 2 parameters and 3 Design Points:
			inp = [[1, 10], [2, 20], [3, 30]]
			keys = ['p1', 'p2']
		If no keys provided use already existing keys
		"""
		self._log_('Direct data input by DPs issued')
		if not self.is_matrix(inp):
			self._log_('Incorrect input format!')
			raise ValueError('Incorrect input format!')
		if not self._param_in and keys is None:
			self._log_('No parameter keys found!')
			raise KeysNotFound
		if len(inp[0]) != len(self._param_in) and keys is None:
			self._log_('Incorrect input format!')
			raise ValueError('Incorrect input format!')
		
		vkeys = keys if keys is not None else self._param_in
		self._param_in = []
		self._param_in_value = defaultdict(list)
		for row in inp:
			for key, elem in zip(vkeys, row):
				elem_s = str(elem).strip()
				self._param_in.append(key.strip().upper())
				self._param_in_value[key.strip().upper()].append(elem_s)
		self.__DPs_imported = len(inp)
		self._log_('Input successful: {} input(s) in {} Design Point(s)'.format(len(self._param_in),self.__DPs_imported), 1)		
		
	def open_archive(self, archive='*.wbpz'):
		"""Search for Workbench archive in working directory and open it."""
		self._log_('Searching for Workbench archive...')
		try:
			file_list = [f for f in glob(archive)]
			wbpz_file = file_list[0]
		except Exception as err_msg:
			self._log_('Archive not found!', 1)
			raise
		
		self.__workfile = wbpz_file.replace('.wbpz','.wbpj')
		
		self._log_('Archive found: ' + self.__workfile, 1)
		self._log_('Unpacking archive...')
		
		try: 
			workbench.Unarchive(ArchivePath=wbpz_file,
							  ProjectPath=self.__workfile,
							  Overwrite=True)
			workbench.ClearMessages()
		except Exception as err_msg:
			self._log_('Unpacking failed!')
			self._log_(err_msg, 1)
			raise
			
		self.__active = True
		self.__DPs = self._get_DPs()
		self.__DPs_present = len(self.__DPs)
		self._log_('Unpacking successful!', 1)
				
	def open_project(self, project='*.wbpj', refresh=False):		
		"""Search for Workbench project in working directory and open it."""
		self._log_('Searching for Workbench project...')
		try:
			file_list = [f for f in glob(project)]
			wbpj_file = file_list[0]
		except Exception as err_msg:
			self._log_('Project not found!', 1)
			raise
			
		self.__workfile = wbpj_file
		
		self._log_('Project found: ' + self.__workfile, 1)
		self._log_('Opening project...')
		try: 
			workbench.Open(FilePath=self.__workfile)
			workbench.ClearMessages()
			if refresh: workbench.Refresh()
		except Exception as err_msg:
			self._log_('Opening failed!')
			self._log_(err_msg, 1)
			raise
			
		self.__active = True
		self.__DPs = self._get_DPs()
		self.__DPs_present = len(self.__DPs)
		self._log_('Success!', 1)		
	
	def set_parameters(self, saveproject=True):
		"""Set imported parameters into Workbench."""
		self._log_('Setting Workbench parameters...')
		if self.__active and self.__DPs_imported <= 0:
			self._log_('No parameters to set!', 1)
			return
		elif not self.__active:
			self._log_('Cannot set parameters: No active project found!', 1)
			raise
		
		self._clear_DPs()
		while self.__DPs_imported > self.__DPs_present:
			self._add_DP(exported=True, retained=True)
		
		try:
			for i, (par, par_values) in enumerate(self._param_in_value.items()):
				for j, par_value in enumerate(par_values):
					self._set_parameter(self.__DPs[j], par, par_value)
		except Exception as err_msg:
			self._log_('An error occured while setting parameters!')
			self._log_(err_msg, 1)
			raise
		self._log_('Success!', 1)
		if saveproject: self._save_project()
		
	def update_project(self, skip_error=True, skip_uncomplete=True):
		"""Update Workbench project."""
		if not self.__active:
			self._log_('Cannot update project: No active project found!', 1)
			raise NoActiveProjectFound
			
		self._log_('Updating Workbench project...')
		workbench.Parameters.ClearDesignPointsCache()
		
		# skip_er = 'SkipDesignPoint' if skip_error else 'Stop'
		# skip_unc = 'Continue' if skip_uncomplete else 'Stop'
		
		self._param_out_value = defaultdict(list)
		
		try:
			if self.__DPs_present == 1:
				workbench.Update()
			else:
				workbench.UpdateAllDesignPoints(DesignPoints=self.__DPs,
										     ErrorBehavior='SkipDesignPoint' if skip_error else 'Stop',
										     CannotCompleteBehavior='Continue' if skip_uncomplete else 'Stop')
			# system1 = GetSystem(Name="SYS")
			# component1 = system1.GetComponent(Name="Model")
			# component1.Update(AllDependencies=True)
			if workbench.IsProjectUpToDate():
				self._log_('Update successful!', 1)
			else:
				self._log_('Project is not up-to-date, see messages below')
				for msg in workbench.GetMessages():
					self._log_(msg.MessageType + ": " + msg.Summary)   
		except Exception as err_msg:  
			self._log_('Project failed to update!')
			self._log_(err_msg, 1)
			raise
		finally:
			self._save_project()
	
	def set_output(self, out_par=None):
		"""
		Set list of output parameters by list
		Example: ['p1', 'p3', 'p4']
		"""
		self._log_('Setting output parameters...')
		if not out_par:
			self._log_('No output parameters specified!')
			raise AttributeError
			
		self._param_out_value = defaultdict(list)
		try:		
			self._param_out = [par.upper() for par in out_par]
		except Exception as err_msg:
			self._log_('Failed to set parameters!')
			self._log_(err_msg, 1)
		
		self._log_('Success: {} outputs specified'.format(len(self._param_out)), 1)
	
	def output_parameters(self, output_file_name=None, full_report_file=None, csv_delim=None, fkey='wb'):
		"""
		Output parameters in a file. Set output_file_name = '' to suppress
		output to a file.
		
		"""
		if output_file_name is None: output_file_name=self._out_file
		if full_report_file is None: full_report_file=self._full_file
		if csv_delim is None: csv_delim=self._csv_delim
		
		if csv_delim is None:
			self._log_('Missing csv delimiter or skipper!', 1)
			raise MissingCSVParameter
		
		if not self._param_out:
			self._log_('No parameters to output!', 1)
			return
		elif not self.__active:
			self._log_('Cannot output parameters: No active project found!', 1)
			raise NoActiveProjectFound
			
		self._param_out_value = defaultdict(list)
		self._log_('Retrieving output parameters... ')
	
		for key in self._param_out:
			for dp in self.__DPs:
				val = self._get_parameter_value(dp, key)
				self._param_out_value[key.upper()].append(val)	
				
		if output_file_name is None or output_file_name == '':
			self._log_('No output file defined! Results stored internally.', 1)
			return
			
		self._log_('Outputing parameters to ' + output_file_name + '...')

		try:
			# Group values by Design Point 
			out_map = self._output_group_by_DPs()
			
			with open(output_file_name, mode=fkey) as out_file:
				out_writer = csvwriter(out_file, delimiter=csv_delim, quotechar='"', quoting=QUOTE_MINIMAL)			
				for row in out_map:
					out_writer.writerow(row)
					
		except Exception as err_msg:
			self._log_('An error occured while outputting parameters!')
			self._log_(err_msg, 1)  
		else:
			self._log_('Output successful!', 1)
		finally:
			try:
				workbench.Parameters.ExportAllDesignPointsData(FileName=full_report_file)
				self._log_('Workbench parametric report written to {}'.format(full_report_file), 1)
			except:
				self._log_('Cannot generate Workbench report!', 1)
				
			self._save_project()
	# ---------------------------------------------------------------
	# Private methods
	# ---------------------------------------------------------------
	def _clear_DPs(self):
		
		dps = self._get_DPs()
		for dp in dps:
			try:
				dp.Delete()
			except:
				pass
			
		self.__DPs = self._get_DPs()
		self.__DPs_present = len(self.__DPs)
		
	def _input_list_by_name(self, inp):
		"""Read from list method"""
		self._log_('Reading input parameters values from list...')
		self._param_in_value = defaultdict(list)
		self.__DPs_imported = 0
		inp_str = [map(str, t) for t in inp]
		try:
			for key, row in zip(self._param_in, inp_str):
				for elem in self._listify(row):
					self._param_in_value[key.upper()].append(elem.strip())
		except Exception as err_msg:
			self._log_('An error occured while processing input!')
			self._log_(err_msg, 1)
			raise
		self.__DPs_imported = len(self._param_in_value[key.upper()])
		
	def _input_dict_by_name(self, inp):
		"""Read from dict method"""
		self._log_('Reading input parameters table...')
		self._param_in_value = defaultdict(list)
		self._param_in = []
		self.__DPs_imported = 0
		try:
			for key, row in inp.items():
				self._param_in.append(key.upper())
				for elem in self._listify(row):
					self._param_in_value[key.upper()].append(str(elem).strip())
		except Exception as err_msg:
			self._log_('An error occured while processing input!')
			self._log_(err_msg, 1)
			raise
		self.__DPs_imported = len(self._param_in_value[key.upper()])
	
	def _output_group_by_DPs(self):
		res = [[0]*len(self._param_out_value) for x in xrange(self.__DPs_present)]
		for i in xrange(self.__DPs_present):
			for j, key in enumerate(self._param_out):
				res[i][j] = self._param_out_value[key][i]
		return res
	
	def _save_project(self):
		workbench.Save(Overwrite=True)
		self._log_('Project Saved!', 1)
		
	def _set_parameter(self, dp, name, value):
		dp.SetParameterExpression(Parameter=self._get_parameter(name), Expression=value)			
		
	def _get_parameter_value(self, dp, name):
		return str(dp.GetParameterValue(self._get_parameter(name)).Value)
	
	def _get_DPs(self):
		"""Get Design Points list from project."""
		if self.__active:
			return workbench.Parameters.GetAllDesignPoints()
		else:
			self._log_('Cannot get Design Points: No active project found!', 1)
			raise NoActiveProjectFound
			
	def _add_DP(self, exported=True, retained=True):
		"""Design Points list."""
		if self.__active:
			dp =  workbench.Parameters.CreateDesignPoint(Exported=exported, Retained=retained)
			self.__DPs.append(dp)
			self.__DPs_present += 1
		else:
			self._log_('Cannot add Design Point: No active project found!', 1)
			raise NoActiveProjectFound
	# ---------------------------------------------------------------		
	# Static methods
	# ---------------------------------------------------------------
	
	@staticmethod
	def is_matrix(inp):
		"""Returns if list of lists is a matrix"""
		return True if len(set(WBInterface.nested_len(inp))) == 1 else False
	
	@staticmethod
	def nested_len(inp):
		"""Returns a list of lengths of nested lists"""
		iter_list = inp
		if isinstance(inp, dict):
			iter_list = inp.values()
		ns_len = []
		for row in iter_list:	
			lrow = WBInterface._listify(row)
			ns_len.append(len(lrow))
		return ns_len
		
	@staticmethod
	def decomment(csvfile):
		"""Remove comments from file."""
		for row in csvfile:
			raw = row.split('#')[0].strip()
			if raw: yield raw
			
	@staticmethod		
	def transpose(inp):
		"""Returns a list of transposed elements"""
		if isinstance(inp, list):
			return list(map(list, zip(*inp)))
		else:
			raise ValueError('Input must be a list')
			
	@staticmethod
	def make_dict(keys, values):
		"""Makes dict, duh"""
		return dict(zip(keys, values))	
	
	@staticmethod
	def _listify(inp):	
		"""Returns list of 1 items if input is not a list"""
		return inp if isinstance(inp, list) else [inp]
	
	@staticmethod
	def _get_parameter(name):
		return workbench.Parameters.GetParameter(Name=name)	
	

		
#__________________________________________________________

class NoActiveProjectFound(Exception):
	def __init__(self):
		pass
	def __str__(self):
		return 'Cannot execute method: No active project found!'
		
class FileSearchParameterNotDefined(Exception):
	def __init__(self):
		pass
	def __str__(self):
		return 'Cannot execute method: No file search parameter defined!'
		
class MissingCSVParameter(Exception):
	def __init__(self):
		pass
	def __str__(self):
		return 'Cannot execute method: Missing csv delimiter or skipper!'
		
class KeysNotFound(Exception):
	def __init__(self):
		pass
	def __str__(self):
		return 'Cannot execute method: Missing parameter keys!'		
#__________________________________________________________