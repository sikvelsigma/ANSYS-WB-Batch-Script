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
#===========================================================================

if __name__ == '__main__':
    filepath = os.path.abspath(__file__)
    filedir = os.path.dirname(filepath)
    os.chdir(filedir)
    print('CWD: ' + os.getcwd())
    print('File: ' + filepath)
    cwdp = lambda x: os.path.join(filedir, x)
    #__________________________________________________________    

    try:
        wb = WBInterface()
        
        wb.open_any(archive_first=True)
        wb.find_and_import_parameters()
        
        """ Set number of cores to max available on this machine using system SYS
        Replace 'SYS' with any other mechanical system.
        This works only with new ribbon interface, which means ANSYS 2019R2 or higher!
        """
        # wb.set_cores_number('SYS')          
        
        """Turn on DMP solver using system SYS
        Replace 'SYS' with any other mechanical system
        This works only with new ribbon interface, which means ANSYS 2019R2 or higher!
        """
        # wb.set_distributed('SYS', True) 
        
        wb.update_project()
        
        """ Try to open SYS and save all Figures (NOT PLOTS!) from a project tree.
        This includes all connected projects which share 'Model' since they are displayed 
        in the same tree.
        This was tested only on ANSYS with new ribbon interface!
        """
        # wb.save_figures('SYS', os.path.join(filedir, 'pictures'), width=1920, height=1080, fontfact=1.35)
        
        wb.output_parameters()
        wb.export_wb_report()
    except Exception as err_msg:
        wb.fatal_error(err_msg)
    finally:
        wb.issue_end()
    


