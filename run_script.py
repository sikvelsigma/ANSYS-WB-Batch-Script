# -*- coding: utf-8 -*-
""" Script by Toybich Egor
"""
#__________________________________________________________
#__________________________________________________________
from glob import glob 
import sys
import platform
import os

def find_module(st_in):
	res = []
	stlist = st_in if isinstance(st_in, list) else [st_in]
	
	for st in stlist:
		try:
			srch = [f for f in glob('{}*.py'.format(st))]
			print('Found: {}'.format(srch))
			srch = srch[0] if srch[0] == '{}.py'.format(st) else srch[-1]
			srch = srch.replace('.py','')
		except:
			res.append(None)
		else:
			res.append(srch)
	return tuple(res) if len(stlist) > 1 else res[0]
	

modules = ['WBInterface', 'ExcelFileReader', 'Logger', 'CSVTable']
modules_files = find_module(modules)

print('Using: {}, {}, {}, {}'.format(*modules))

if modules_files[0]: exec('from {} import WBInterface'.format(modules_files[0]))
if modules_files[1]: exec('from {} import ExcelFileReader'.format(modules_files[1]))
if modules_files[2]: exec('from {} import Logger'.format(modules_files[2]))
if modules_files[3]: exec('import {} as CSVTable'.format(modules_files[3]))

#__________________________________________________________

if __name__ == '__main__':
#__________________________________________________________

	wb = WBInterface()

	try:
		wb.open_archive()
	except:
		try:
			wb.open_project()
		except:
			wb.log('Nothing to open!')
			raise

	try:
		wb.read_control()
		wb.read_input()
		wb.set_parameters()
		wb.update_project()
		wb.output_parameters()
	except Exception as err_msg:
		wb.log('CRITICAL ERROR')
		wb.log(err_msg)
	else:
		wb.log('RUN SUCCESSFUL')
	
	wb.log('END SCRIPT')


