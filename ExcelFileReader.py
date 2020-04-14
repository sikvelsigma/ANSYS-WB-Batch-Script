# -*- coding: utf-8 -*-
""" Script by Toybich Egor
"""
import clr
from System import Type, Activator


try:
	clr.AddReferenceByName('Microsoft.Office.Interop.Excel')
except:
	clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

import Microsoft.Office.Interop.Excel as Excel
from System.Runtime.InteropServices import Marshal

__version__ = '1.0.3'

class ExcelFileReader(object):
	"""Excel Interop adapter class for reading a file
	Methods (Excel numbers elements from 1, not 0!):
	
		activesheet_set(num:int) :change active sheet number
		g_xlr2list(cell1:str, cell2:str) --> list :gets range to list
		g_xlr2dict(cell1:str, cell2:str, d_column=1:int) --> dict :gets range to dict with d_column as keys
		g_active_xlr2list() --> list :gets used range to list
		g_active_xlr2dict(d_column=1:int)) --> dict :gets used range to dict with d_column as keys
		
	"""
	__version__ = '1.0.2'
	
	#====================================================================== 
	@property
	def filename(self):
		return self._filename
	#====================================================================== 
	@classmethod
	def kill_excel_processes(cls, times=10):
		print('ExcelFileReader| Killing excel 10 times...')
		oShell = Activator.CreateInstance(Type.GetTypeFromProgID("WScript.Shell"))
		for _ in range(10):
			oShell.Run("taskkill /im EXCEL.EXE",0 , True)
		
		print('ExcelFileReader| Done killing...')
		
	#====================================================================== 
	def __init__(self, filename):
		self._set_none()
		self._filename = filename
		self.application = Excel.ApplicationClass()
		self.workbooks = self.application.Workbooks
		try:
			self.workbook = self.workbooks.Open(filename)
		except:
			print('ExcelFileReader| Incorrect path!')
			raise
			self.__del__
		else:
			self.worksheets = self.workbook.Worksheets 
			self.activesheet = self.worksheets[1]
			self.application.Visible = False
			self.application.DisplayAlerts = False

	#====================================================================== 	
	def activesheet_set(self, num=1):
		"""Change active sheet number
		"""
		try:
			self.activesheet = self.worksheets[num]
		except:
			print('ExcelFileReader| Invalid sheet number!')
			self.activesheet = self.worksheets[1]

	#====================================================================== 
	def g_xlr2list(self, cell1, cell2):
		"""Gets excel range into a list
		"""
		xlrange = self.activesheet.Range[cell1,cell2]
		list1 = self.xlr2list(xlrange)
		
		Marshal.FinalReleaseComObject(xlrange)
		xlrange = None
		
		return list1
		
		#====================================================================== 
	def g_xlr2dict(self, cell1, cell2, d_column=1):
		"""Gets excel range into a dictionary
		"""
		return self.list2dict(self.g_xlr2list(cell1, cell2), d_column-1)
	#====================================================================== 
	def g_active_xlr2list(self):
		"""Gets used range in an active sheet into a dictionary
		"""
		xlrange = self.activesheet.UsedRange
		list1 = self.xlr2list(xlrange)
		
		Marshal.FinalReleaseComObject(xlrange)
		xlrange = None
		return list1

	#====================================================================== 	
	def g_active_xlr2dict(self, d_column=1):
		"""Gets used range in an active sheet into a dictionary
		"""
		xlrange = self.activesheet.UsedRange
		list1 = self.xlr2list(xlrange)
		
		Marshal.FinalReleaseComObject(xlrange)
		xlrange = None
		return self.list2dict(list1,d_column-1)
	

	#====================================================================== 	
	@staticmethod
	def list2dict(lrange, key_column=0):
		"""Converts list to a dictionary with some column as keys
		"""
		key_column = int(key_column)
		if key_column >= 0 or key_column < len(lrange[0]):
			rdict = {}
			for row in lrange:
				cutrow = []
				for i, elem in enumerate(row):
					if not(i == key_column):
						cutrow.append(elem)
				if row[key_column] in rdict:
					print('ExcelFileReader| Duplicate key found!')
				else:
					rdict[row[key_column]] = cutrow
			return rdict
		else:
			print('ExcelFileReader| Invalid key number!')
			return 0
			raise
			
	#====================================================================== 		
	@staticmethod
	def xlr2list(xlrange):
		"""Converts excel range to python list
		"""
		rows = xlrange.Rows
		columns = xlrange.Columns
		n = rows.Count
		m = columns.Count

		vals = [[0] * m for i in range(n)]
		for i in range(n):
			for j in range(m):
				tmp = xlrange[i+1,j+1].Value2
				try:
					if int(tmp) == tmp:
						tmp = int(tmp)
				except:
					# tmp = tmp.unicode('utf8')
					pass
				# print tmp
				vals[i][j] = tmp
					
		Marshal.FinalReleaseComObject(rows)
		Marshal.FinalReleaseComObject(columns)
		Marshal.FinalReleaseComObject(xlrange)
		rows = None
		columns = None
		xlrange = None
		
		return vals		

	#====================================================================== 	
	def __del__(self):
		self.workbook.Close(False)
		self.application.Quit()
		
		Marshal.FinalReleaseComObject(self.activesheet)
		Marshal.FinalReleaseComObject(self.worksheets)
		Marshal.FinalReleaseComObject(self.workbooks)
		Marshal.FinalReleaseComObject(self.workbook)
		Marshal.FinalReleaseComObject(self.application)
		
		self._set_none()
		# del self.activesheet, self.worksheets, self.workbook
		# del self.workbooks, self.application
	#====================================================================== 
	def _set_none(self):
		self.activesheet = None
		self.worksheets = None 
		self.workbooks = None
		self.workbook = None
		self.application = None
	
	#====================================================================== 
	def __repr__(self):
		return 'ExcelInterface on file: {}'.format(self.filename)	
	#====================================================================== 
	def __str__(self):
		return self.filename
	#====================================================================== 	
	def __exit__(self, exc_type, exc_value, traceback):
		return True
		self.__del__
	#====================================================================== 	
	def __enter__(self):
		return self
	#====================================================================== 	
	def __len__(self):
		cnt = 0
		for dummy in self.worksheets:
			cnt += 1
			
		Marshal.FinalReleaseComObject(dummy)
		dummy = None
		
		return cnt