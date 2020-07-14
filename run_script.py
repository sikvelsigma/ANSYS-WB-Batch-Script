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
        except: res.append(None)
        else: res.append(srch)
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
    
    wb = WBInterface()

    try:
        wb.open_any(archive_first=True)
        wb.find_and_import_parameters()
        
        #----------------------------------------------------------------
        # Parameters can be imported directly
        # input_p = {'p1':[1, 2, 3], 'p2':[10, 20, 30]}
        # wb.input_by_name(input_p)
        # output_p = ['p1', 'p3', 'p4']
        # wb.set_output(output_p)
        
        # wb.import_parameters()
        #============================================================================== 
        # Sets maximum number of cores
        # wb.set_cores_number('SYS')   
        
        # Activate distrubuted solver
        # wb.set_distributed('SYS', True)
        
        # Sets unit system
        # wb.set_unit_system('SYS', unit_sys='NMM')
        #============================================================================== 
        
        wb.update_project()
        
        #============================================================================== 
        # Set figure scale 
        # wb.set_figures_scale('SYS', scale='auto') 
        
        # Picture parameters
        # overview_args = dict(width=1920, height=1080, zoom_to_fit=True, view='iso')
        # mesh_args = dict(width=1920*2, height=1080*2, zoom_to_fit=True, view='iso')
        # env_args = dict(width=1920, height=1080, zoom_to_fit=True, fontfact=1.5, view='iso')
        # fig_args = dict(width=1920, height=1080, zoom_to_fit=True, fontfact=1.35, view='iso')
        # ani_args = dict(width=1920/2, height=1080/2, zoom_to_fit=True, scale='auto', frames=20, view='iso')
        
        
        # Save pictures parameters
        # wb.save_overview('SYS', cwdp('pictures'), 'model_overview.jpg', **overview_args) 
        # wb.save_mesh_view('SYS', cwdp('pictures'), 'mesh.png', **mesh_args) 
        # wb.save_setups_view('SYS', cwdp('pictures'), **env_args) 
        # wb.save_figures('SYS', cwdp('pictures'), **fig_args) 
        # wb.save_animations('SYS', cwdp('animations'), **ani_args)
        
        #----------------------------------------------------------------
        # Can also save for each Design Point
        # for i, dp in enumerate(wb.DPs):
            # wb.set_active_DP(dp)
            
            # mesh_file = 'mesh_DP{}.png'.format(i)
            # mesh_args = dict(width=1920*2, height=1080*2, zoom_to_fit=True)
            # wb.save_mesh_view('SYS', cwdp('pictures'), mesh_file, **mesh_args)

            # fig_pref = 'Result_DP{}'.format(i)
            # fig_args = dict(fpref=fig_pref, width=1920*2, height=1080*2, zoom_to_fit=True, fontfact=1.35)
            # wb.save_figures('SYS', cwdp('pictures'), **fig_args) 
        
            # env_pref = 'Setup_DP{}'.format(i)
            # env_args = dict(fpref=env_pref, width=1920, height=1080, zoom_to_fit=True, fontfact=1.5)
            # wb.save_setups_view('SYS', cwdp('pictures'), **env_args) 
        #============================================================================== 
        
        wb.output_parameters()
        wb.export_wb_report()
    except Exception as err_msg:
        wb.fatal_error(err_msg)
    finally:
        wb.archive_if_complete()
        wb.issue_end()
    


