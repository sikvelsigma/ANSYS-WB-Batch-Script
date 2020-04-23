# -*- coding: utf-8 -*-
""" Script by Toybich Egor
"""
__version__ = '1.0.1'

import csv
from glob import glob 

def read_to_list(filename, csv_delim=','):
	
	try:
		file_list = [f for f in glob(filename)]
		file_found = file_list[0]
	except:
		print('File not found!')
		raise
		
	try:
		res = []
		with open(file_found, 'r') as csvfile:
			spamreader = csv.reader(decomment(csvfile), delimiter=csv_delim)
			for i, row in enumerate(spamreader):
				res.append(row)
		return res
	except Exception as err_msg:
		print('An error occured wile reading CSV file!')
		return None
		raise
	

def read_to_dict(filename, key_column=1, csv_delim=','):
	cd = csv_delim
	return list2dict(read_to_list(filename, csv_delim=cd), key_column-1)
	
def list2dict(lrange, key_el=0):
	"""Converts list to a dictionary with some column as keys
	"""
	key_el = int(key_el)
	if key_el >= 0 or key_el < len(lrange[0]):
		rdict = {}
		for row in lrange:
			cutrow = []
			for i, elem in enumerate(row):
				if not(i == key_el):
					cutrow.append(elem)
			if row[key_el] in rdict:
				print('Duplicate key found!')
			else:
				rdict[row[key_el]] = cutrow
		return rdict
	else:
		print('Invalid key number!')
		return 0
		raise
		
def decomment(csvfile):
	"""Remove comments from file."""
	for row in csvfile:
		raw = row.split('#')[0].strip()
		if raw: yield raw