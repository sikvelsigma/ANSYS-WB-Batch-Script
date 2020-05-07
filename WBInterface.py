# -*- coding: utf-8 -*-
""" Script by Toybich Egor
Note: Workbench uses IronPython (Python 2.7)!
"""
#__________________________________________________________
#__________________________________________________________
import os
import shutil
# import csv
from glob import glob
from functools import partial 
from collections import defaultdict

from csv import reader as csvreader
from csv import writer as csvwriter
from csv import QUOTE_MINIMAL

from glob import glob 
from time import sleep
from System.Threading import Thread, ThreadStart
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

__version__ = '3.0.0'
#__________________________________________________________
class WBInterface(object):
    """
    A class to open Workbench project/archive, input/output parameters
    and start calculations.
    """
    __version__ = '3.0.0'
    
    _macro_def_dir = '_TempScript'
    __macro_dir_path = ''
    __macros_count = 0    
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
        """Returns if there is an opened Workbench project."""
        return self.__active
        
    @property
    def parameters(self):
        """Returns IO parameters as dictionaries."""
        return (self._param_in_value, 	self._param_out_value)
        
    @property
    def failed_to_update(self):
        """Returns if project failed to update"""
        return self.__failed_to_update
       
    # ---------------------------------------------------------------		
    # Magic methods
    # ---------------------------------------------------------------
    
    def __init__(self, logger = None, out_file='output.txt', full_report_file='full_report.txt', 
                 control_file_template='*_control.csv', input_file_template='*_input.csv', csv_delim=',', 
                 csv_skip='no', loginfo=None, apdl_log='APDL_output.txt', async_timer=0.5):
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
        try: self._log_ = partial(self._logger.log, info=log_prefix)
        except: self._log_ = self._logger.log
        
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
        self.__failed_to_update = False
        
        self.__control_srch_default = '*.control'	#: in-build key string to search for cotrol file
        self.__input_srch_default = '*.input'		#: in-build  key string to search for input file
        
        self.__apdl_log = apdl_log
        
        dir = os.path.join(os.getcwd(),'_ProjectScratch')
        args = dict(outfile=self.__apdl_log, watch_dir=dir, watchfile='solve.out', sec=async_timer)
        self.__async_log = AsyncLogChecker(**args)
        
        self._log_('Class version: ' + self.__version__ , 1)
        
        try: self.runtime = self._logger.runtime
        except: self.runtime = None
        
    # --------------------------------------------------------------------        
    def __del__(self):   
        if self.__macro_dir_path and os.path.exists(__macro_dir_path):
            shutil.rmtree(self.__macro_dir_path, ignore_errors=True)
    
    def __bool__(self):
        return self.__active
        
    def __str__(self):
        return self.__workfile
    # ---------------------------------------------------------------		
    # Public methods
    # ---------------------------------------------------------------
    
    def read_control(self, control_file_template=None, csv_delim=None, csv_skip=None):
        """
        Read csv file with IO parameter list.
        Example of control file format with 2 inputs and 3 outputs:
            #type 'no' to skip inputs
            p3,p2
            #type 'no' to skip outputs
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
                    self._log_('Control file not found! No parameters will be used', 1)
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

            
    # --------------------------------------------------------------------   
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
    # --------------------------------------------------------------------                 
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
    # -------------------------------------------------------------------- 
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
    # --------------------------------------------------------------------     
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
    # --------------------------------------------------------------------             
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
        self._log_('Success', 1)		
    # -------------------------------------------------------------------- 
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
        self._log_('Success', 1)
        if saveproject: self._save_project()
    # --------------------------------------------------------------------     
    def update_project(self, skip_error=True, skip_uncomplete=True):
        """Update Workbench project."""
        if not self.__active:
            self._log_('Cannot update project: No active project found!', 1)
            raise NoActiveProjectFound
            
        self._log_('Updating Workbench project...')
        self._log_('Check _ProjectScratch directory for solver logs')
        workbench.Parameters.ClearDesignPointsCache()
        self.__failed_to_update = False
        # skip_er = 'SkipDesignPoint' if skip_error else 'Stop'
        # skip_unc = 'Continue' if skip_uncomplete else 'Stop'
        
        self._param_out_value = defaultdict(list)
        
        try:
            self.__async_log.is_solving = True
            newthread = Thread(ThreadStart(self.__async_log.main))
            newthread.Start()
            if self.__DPs_present == 1:
                workbench.Update()
            else:
                args = dict(DesignPoints=self.__DPs, ErrorBehavior='SkipDesignPoint' if skip_error else 'Stop',
                            CannotCompleteBehavior='Continue' if skip_uncomplete else 'Stop')                          
                workbench.UpdateAllDesignPoints(**args) 
            self.__async_log.is_solving = False
            
            if workbench.IsProjectUpToDate():
                self._log_('Update successful', 1)
            else:
                self._log_('Project is not up-to-date, see messages below')
                for msg in workbench.GetMessages():
                    self._log_(msg.MessageType + ": " + msg.Summary)   
                self._logger.blank()
        except Exception as err_msg:  
            self._log_('Project failed to update!')
            self._log_(err_msg, 1)
            self.__failed_to_update = True
        finally:
            self._save_project()
    # -------------------------------------------------------------------- 
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
    # -------------------------------------------------------------------- 
    def save_project(self):
        workbench.Save(Overwrite=True)
        self._log_('Project Saved', 1)
    # --------------------------------------------------------------------
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
    
        try:
            for key in self._param_out:
                for dp in self.__DPs:
                    val = self._get_parameter_value(dp, key)
                    self._param_out_value[key.upper()].append(val)	
        except Exception as err_msg:
            self._log_('Failed to retriev parameters!')
            self._log_(err_msg, 1)
            return
            
        if output_file_name is None or output_file_name == '':
            self._log_('No output file defined! Results were stored internally', 1)
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
            self._log_('Output successful', 1)
        finally:
            try:
                workbench.Parameters.ExportAllDesignPointsData(FileName=full_report_file)
                self._log_('Workbench parametric report written to {}'.format(full_report_file), 1)
            except:
                self._log_('Cannot generate Workbench report!', 1)
                
            self._save_project()
    
    def copy_from_userfiles(self, template_str, target):
        self.copy_files(template_str, workbench.GetUserFilesDirectory() , target)
        
    def move_from_userfiles(self, template_str, target):
        self.move_files(template_str, workbench.GetUserFilesDirectory() , target)
    # ---------------------------------------------------------------
    # JScript Methods
    # --------------------------------------------------------------- 
    def set_cores_number(self, container, value, module='Model', ignore_js_err=True):
        """
        Tested only on ANSYS2019R3! ANSYS version with old interface won't work!
        Sets number of cores in Mechanical. 
        
        Warning: number of cores for Mechanical is a global setting and will be
        saved as default. Next project run on this machine will use this setting!
        
        Arg:
            container: str; specify a Mechanical system, e.g. 'SYS'
            value: int
            module: str; module to open
            ignore_js_err: bool; wraps js main cammands in a try block
        """
        if not isinstance(value, int) or value < 0:
            self._log_('Error: Attempted to set incorrect number of cores: {}'.format(value))
            self._log_('Number of cores used unchanged!', 1)
            return
            
        self._log_('Setting number of cores to {}'.format(value))
        self._log_('Do not set this number higher than a number of physical cores on a machine!'.format(value))
        
        # jscode = 'DS.Script.Configure_setNumberOfCores("{}")'.format(value)
        
        jsfun = '''
             function setNumberOfCores(value)
             {
                    var jobHandlerManager = DS.Script.getJobHandlerManager();
                    if (jobHandlerManager == null)
                    {
                        return;
                    }
                    var defaultHandler = DS.Script.getDefaultHandler(jobHandlerManager);
                    if (defaultHandler == null)
                    {
                        return;
                    }
                    var numProcessors = parseInt(value);
                    if (!isNaN(numProcessors) && numProcessors >= 1 && defaultHandler.MaxNumberProcessors != numProcessors)
                    {
                        defaultHandler.MaxNumberProcessors = numProcessors;
                        jobHandlerManager.Save();
                        return;
                    }
             }
        '''
        
        jsmain = 'setNumberOfCores("{}");'.format(value)
        
        if ignore_js_err: jscode = jsfun + self._try_wrapper_js(jsmain)
        else: jscode = jsfun + jsmain
        
        try:
            self._send_js_macro(container, jscode, module)
        except Exception as err_msg:
            self._log_('An error occured!')
            self._log_('Your ANSYS might use old interface')
            self._log_(err_msg, 1)
        
    # -------------------------------------------------------------------- 
    def set_distributed(self, container, value, module='Model', ignore_js_err=True):
        """
        Tested only on ANSYS2019R3! ANSYS version with old interface won't work!
        Activates/deactivates DMP in Mechanical. 
        
        Warning: this property is a global setting and will be
        saved as default. Next project run on this machine will use this setting!
        
        Arg:
            container: str; specify a Mechanical system, e.g. 'SYS'
            value: boolean
            module: str; module to open
            ignore_js_err: bool; wraps js main cammands in a try block
        """
        value_js = self._bool_js(value)
        if value_js == 'true':
            self._log_('Distributed solver: Enabled')
        else:
            self._log_('Distributed solver: Disabled')
            
        jsfun = '''
            function setDMP(value)
            {
                var jobHandlerManager = DS.Script.getJobHandlerManager();
                if (jobHandlerManager == null)
                {
                    return;
                }
                var defaultHandler = DS.Script.getDefaultHandler(jobHandlerManager);
                if (defaultHandler == null)
                {
                    return;
                }
                defaultHandler.DistributeAnsysSolution = value;
                jobHandlerManager.Save();
            }
        '''
        jsmain = 'setDMP({});'.format(value_js)
        
        
        if ignore_js_err: jscode = jsfun + self._try_wrapper_js(jsmain)
        else: jscode = jsfun + jsmain
        
        try:
            self._send_js_macro(container, jscode, module)
        except Exception as err_msg:
            self._log_('An error occured!')
            self._log_('Your ANSYS might use old interface')
            self._log_(err_msg, 1)
    
    # -------------------------------------------------------------------- 
    def save_overview(self, container, fpath, filename, width=0, height=0, fontfact=1, zoom_to_fit=False, module='Model', ignore_js_err=True):
        """
        Tested only on ANSYS2019R3!
        Saves model overview.
        This is a modified ripped function from ANSYS to dump all figures in cwd.
        
        Arg:
            container: str; specify a Mechanical system, e.g. 'SYS'
            fpath: str; save directory
            filename: str
            module: str; module to open
            ignore_js_err: bool; wraps js main cammands in a try block
        """  
        try:
            basename, ext = filename.split('.')
        except:
            self._log_('Error: Could not determine a file extention!', 1)
            return None
        
        if ext == 'png': mode = 0
        elif ext == 'jpg' or ext == 'jpeg': mode = 1
        else:
            self._log_('Unsupported file extention!', 1)
            return None
        
        self._log_('Saving model overview in {}'.format(os.path.join(fpath, filename)))      
        if not os.path.exists(fpath): os.makedirs(fpath)
                    
        jsfun = self.__jsfun_savepics() + '''
            function DumpOverview(pdir, pHeight, pWidth, pFontFactor, pFit, pName, pMode) {                                          
                var clsidModel = 104; // model
               
                var activeObjs = DS.Tree.AllObjects;
                if (!activeObjs)
                    return;
                var numObjs = activeObjs.Count;
                
                var prevColor1 = DS.Graphics.Scene.Color(1); 
                var prevColor2 = DS.Graphics.Scene.Color(2);
                var prevColor5 = DS.Graphics.Scene.Color(5);
                var prevColor6 = DS.Graphics.Scene.Color(6);
                var prevLegend = DS.Graphics.LegendVisibility;
                var prevRuler = DS.Graphics.RulerVisibility;
                var prevTriad = DS.Graphics.TriadOn;
                var prevRandom = DS.Graphics.RandomColors; 
                
                prepocPicOutput(pHeight, pWidth, pFontFactor, prevColor5, prevColor6);                                                                                                                        
                
                //debugger;
                // ====Make model overview====
                DS.Graphics.TriadOn = false;
                DS.Graphics.LegendVisibility = false; 
                DS.Graphics.RulerVisibility = false;  
                DS.Graphics.RandomColors = false;
                saveObjectsPictures(clsidModel, activeObjs, pdir, pName, "", pFit, pMode);                             
                
                // ====Restore settings====              
                postPicOutput(prevColor1, prevColor2, prevColor5, prevColor6, prevLegend, prevRuler, prevTriad, prevRandom)                         
            }
        '''
        jsmain = 'DumpOverview("{}", {}, {}, {}, {}, "{}", {});'.format(self._winpath_js(fpath), height, width, fontfact, 
                                                              self._bool_js(zoom_to_fit), basename, mode)
        
        if ignore_js_err: jscode = jsfun + self._try_wrapper_js(jsmain)
        else: jscode = jsfun + jsmain
         
        try:
            self._send_js_macro(container, jscode, module, visible=True)
        except Exception as err_msg:
            self._log_('An error occured!')
            self._log_(err_msg, 1) 
    # -------------------------------------------------------------------- 
    def save_mesh_view(self, container, fpath, filename, width=0, height=0, fontfact=1, zoom_to_fit=False, module='Model', ignore_js_err=False):
        """
        Tested only on ANSYS2019R3!
        Saves mesh view.
        
        Arg:
            container: str; specify a Mechanical system, e.g. 'SYS'
            fpath: str; save directory
            filename: str
            module: str; module to open
            ignore_js_err: bool; wraps js main cammands in a try block
        """  
        try:
            basename, ext = filename.split('.')
        except:
            self._log_('Error: Could not determine a file extention!', 1)
            return None
        
        if ext == 'png': mode = 0
        elif ext == 'jpg' or ext == 'jpeg': mode = 1
        else:
            self._log_('Unsupported file extention!', 1)
            return None
        
        self._log_('Saving model overview in {}'.format(os.path.join(fpath, filename)))      
        if not os.path.exists(fpath): os.makedirs(fpath)
                    
        jsfun = self.__jsfun_savepics() + '''
            function DumpMesh(pdir, pHeight, pWidth, pFontFactor, pFit, pName, pMode) {                                          
                var clsidMesh = 127; // mesh
               
                var activeObjs = DS.Tree.AllObjects;
                if (!activeObjs)
                    return;
                var numObjs = activeObjs.Count;
                
                var prevColor1 = DS.Graphics.Scene.Color(1); 
                var prevColor2 = DS.Graphics.Scene.Color(2);
                var prevColor5 = DS.Graphics.Scene.Color(5);
                var prevColor6 = DS.Graphics.Scene.Color(6);
                var prevLegend = DS.Graphics.LegendVisibility;
                var prevRuler = DS.Graphics.RulerVisibility;
                var prevTriad = DS.Graphics.TriadOn;
                var prevRandom = DS.Graphics.RandomColors; 
                
                prepocPicOutput(pHeight, pWidth, pFontFactor, prevColor5, prevColor6);                                                                                                                        
                
                //debugger;
                // ====Make model overview====
                DS.Graphics.TriadOn = false;
                DS.Graphics.LegendVisibility = false; 
                DS.Graphics.RulerVisibility = false;        
                DS.Graphics.RandomColors = false; 
                saveObjectsPictures(clsidMesh, activeObjs, pdir, pName, "", pFit, pMode);                             
                
                // ====Restore settings====              
                postPicOutput(prevColor1, prevColor2, prevColor5, prevColor6, prevLegend, prevRuler, prevTriad, prevRandom)                         
            }
        '''
        jsmain = 'DumpMesh("{}", {}, {}, {}, {}, "{}", {});'.format(self._winpath_js(fpath), height, width, fontfact, 
                                                              self._bool_js(zoom_to_fit), basename, mode)
        
        if ignore_js_err: jscode = jsfun + self._try_wrapper_js(jsmain)
        else: jscode = jsfun + jsmain
         
        try:
            self._send_js_macro(container, jscode, module, visible=True)
        except Exception as err_msg:
            self._log_('An error occured!')
            self._log_(err_msg, 1) 
       
    def save_setups_view(self, container, fpath, width=0, height=0, fontfact=1, zoom_to_fit=False, module='Model', ignore_js_err=False):
        """
        Tested only on ANSYS2019R3!
        Saves all environments setups in png.
        
        Arg:
            container: str; specify a Mechanical system, e.g. 'SYS'
            fpath: str; save directory
            module: str; module to open
            ignore_js_err: bool; wraps js main cammands in a try block
        """  
        self._log_('Saving all environment setups in {}'.format(fpath))      
        if not os.path.exists(fpath): os.makedirs(fpath)
        
        jsfun = self.__jsfun_savepics() + '''
            function DumpSetups(pdir, pHeight, pWidth, pFontFactor, pFit) {                                          
                var clsidEnv = 105; // load cases
               
                var activeObjs = DS.Tree.AllObjects;
                if (!activeObjs)
                    return;
                var numObjs = activeObjs.Count;
                
                var prevColor1 = DS.Graphics.Scene.Color(1); 
                var prevColor2 = DS.Graphics.Scene.Color(2);
                var prevColor5 = DS.Graphics.Scene.Color(5);
                var prevColor6 = DS.Graphics.Scene.Color(6);
                var prevLegend = DS.Graphics.LegendVisibility;
                var prevRuler = DS.Graphics.RulerVisibility;
                var prevTriad = DS.Graphics.TriadOn;
                var prevRandom = DS.Graphics.RandomColors; 
                
                prepocPicOutput(pHeight, pWidth, pFontFactor, prevColor5, prevColor6);                                                                                                                        
                                                                         
                // ====Dump all enviroments====
                DS.Graphics.TriadOn = true;
                DS.Graphics.LegendVisibility = true; 
                DS.Graphics.RulerVisibility = false;
                DS.Graphics.RandomColors = true;
                saveObjectsPictures(clsidEnv, activeObjs, pdir, "", "Setup", pFit, 0);
              
                // ====Restore settings====              
                postPicOutput(prevColor1, prevColor2, prevColor5, prevColor6, prevLegend, prevRuler, prevTriad, prevRandom)                         
            }
        '''                        
        jsmain = 'DumpSetups("{}", {}, {}, {}, {});'.format(self._winpath_js(fpath), height, width, fontfact, self._bool_js(zoom_to_fit))
        
        if ignore_js_err: jscode = jsfun + self._try_wrapper_js(jsmain)
        else: jscode = jsfun + jsmain
         
        try:
            self._send_js_macro(container, jscode, module, visible=True)
        except Exception as err_msg:
            self._log_('An error occured!')
            self._log_(err_msg, 1) 

        
    # -------------------------------------------------------------------- 
    def save_figures(self, container, fpath, width=0, height=0, fontfact=1, zoom_to_fit=False, module='Model', ignore_js_err=False):
        """
        Tested only on ANSYS2019R3!
        Saves all figures (not plot!) in png. 
        
        Arg:
            container: str; specify a Mechanical system, e.g. 'SYS'
            fpath: str; save directory
            module: str; module to open
            ignore_js_err: bool; wraps js main cammands in a try block
        """
        self._log_('Saving all figures in {}'.format(fpath))
        
        if not os.path.exists(fpath): os.makedirs(fpath)

        
        jsfun = self.__jsfun_savepics() + '''
            function DumpAllFigures(pdir, pHeight, pWidth, pFontFactor, pFit) {                                          
                var clsidFigure = 147; // figures
                
                var activeObjs = DS.Tree.AllObjects;
                if (!activeObjs)
                    return;
                var numObjs = activeObjs.Count;
                
                var prevColor1 = DS.Graphics.Scene.Color(1); 
                var prevColor2 = DS.Graphics.Scene.Color(2);
                var prevColor5 = DS.Graphics.Scene.Color(5);
                var prevColor6 = DS.Graphics.Scene.Color(6);
                var prevLegend = DS.Graphics.LegendVisibility;
                var prevRuler = DS.Graphics.RulerVisibility;
                var prevTriad = DS.Graphics.TriadOn;
                var prevRandom = DS.Graphics.RandomColors; 
                
                prepocPicOutput(pHeight, pWidth, pFontFactor, prevColor5, prevColor6);                                                                                                                        
                                                                                                                     
                // ====Dump all figures====
                DS.Graphics.TriadOn = true;
                DS.Graphics.LegendVisibility = true; 
                DS.Graphics.RulerVisibility = false;
                DS.Graphics.RandomColors = false;
                saveObjectsPictures(clsidFigure, activeObjs, pdir, "", "Result", pFit, 0);
                
                // ====Restore settings====              
                postPicOutput(prevColor1, prevColor2, prevColor5, prevColor6, prevLegend, prevRuler, prevTriad, prevRandom)                         
            }
        '''
        jsmain = 'DumpAllFigures("{}", {}, {}, {}, {});'.format(self._winpath_js(fpath), height, width, fontfact, self._bool_js(zoom_to_fit))
        
        if ignore_js_err: jscode = jsfun + self._try_wrapper_js(jsmain)
        else: jscode = jsfun + jsmain
         
        try:
            self._send_js_macro(container, jscode, module, visible=True)
        except Exception as err_msg:
            self._log_('An error occured!')
            self._log_(err_msg, 1) 
     
   
    # -------------------------------------------------------------------- 
    def set_unit_system(self, container, unit_sys, module='Model', ignore_js_err=True):
        """
        Tested only on ANSYS2019R3!
        Changes unit system is Mechanical.
        
        Arg:
            container: str; specify a Mechanical system, e.g. 'SYS'
            unit_sys_id: str; unit system; 
                'MKS', 'CGS', 'NMM', 'BFT', 'BIN', 'UMKS', 'NMMton', 'NMMdat'          
            module: str; module to open
            ignore_js_err: bool; wraps js main cammands in a try block
        """       
        if unit_sys == 'MKS': 
            unit_sys_id = 0
            unit_msg = 'Metric (m, kg, N, s, V, A)'
        elif unit_sys == 'CGS': 
            unit_sys_id = 1
            unit_msg = 'Metric (cm, g, dyne, s, V, A)'
        elif unit_sys == 'NMM': 
            unit_sys_id = 2
            unit_msg = 'Metric (mm, kg, N, s, mV, mA)'
        elif unit_sys == 'BFT': 
            unit_sys_id = 3
            unit_msg = 'U.S. Customary (ft, lbm, lbf, °F, s, V, A)'
        elif unit_sys == 'BIN': 
            unit_sys_id = 4
            unit_msg = 'U.S. Customary (in, lbm, lbf, °F, s, V, A)'
        elif unit_sys == 'UMKS': 
            unit_sys_id = 9
            unit_msg = 'Metric (um, kg, uN, s, V, mA)'
        elif unit_sys == 'NMMton': 
            unit_sys_id = 13
            unit_msg = 'Metric (mm, t, N, s, mV, mA)'
        elif unit_sys == 'NMMdat': 
            unit_sys_id = 14
            unit_msg = 'Metric (mm, dat, N, s, mV, mA)'
        else: 
            self._log_('Cannot set unit system: unknown system ID = {}'.format(unit_sys), 1)
            return
        
        self._log_('Setting units to {}: {}'.format(unit_sys, unit_msg))
        
        jsfun = '''
            function setUnits(sysId) {               
                DS.UnitSystemID = sysId;
                DS.Graphics.Redraw(1); 
                DS.Script.FireFinished();
            }
        '''
    
        jsmain = 'setUnits({});'.format(unit_sys_id)
        
        if ignore_js_err: jscode = jsfun + self._try_wrapper_js(jsmain)
        else: jscode = jsfun + jsmain
         
        try:
            self._send_js_macro(container, jscode, module, visible=True)
        except Exception as err_msg:
            self._log_('An error occured!')
            self._log_(err_msg, 1) 
        # -------------------------------------------------------------------- 
    def set_figures_scale(self, container, scale, module='Model', ignore_js_err=True):
        """
        Tested only on ANSYS2019R3!
        Sets scale of figures. 
        
        Arg:
            container: str; specify a Mechanical system, e.g. 'SYS'         
            module: str; module to open
            scale: num or str; valid strings are
                '0.5auto', 'auto', '2auto', '5auto', 'undef', 'actual'
            ignore_js_err: bool; wraps js main cammands in a try block
        """   
        strwrap = lambda x: '"{}"'.format(x)
        undf_st = ('undef','undeformed')
        
        if isinstance(scale, str):
            arg = strwrap(scale)
            if scale == 'auto'       : msg = 'auto'               
            elif scale == '2auto'    : msg = 'double auto'               
            elif scale == '5auto'    : msg = 'five auto'                
            elif scale in undf_st    : msg = 'undeformed'               
            elif scale == '0.5auto'  : msg = 'half auto'               
            elif scale == 'actual'   : msg = 'true scale'
                
            else:
                self._log_('Incorrect scale!')
                return
        else:
            arg = scale
            msg = scale
            if scale < 0: 
                self._log_('Incorrect scale!')
                return 
                     
            
        self._log_('Setting figures scale to {}'.format(msg))
        
        jsfun = '''
            function setScale(pScale) {    
                var clsidFigure = 147; // figures
                
                var activeObjs = DS.Tree.AllObjects;
                if (!activeObjs)
                    return;

                var numObjs = activeObjs.Count;
                DS.Graphics.StreamMode = 1; 
                
                for (var i = 1; i <= numObjs; i++) {
                    var objActive = activeObjs.Item(i);

                    if (objActive && objActive.ID && (objActive.Class == clsidFigure))
                    {
                        DS.Graphics.Draw2(objActive.ID);

                        if (!isNaN(pScale)) {       		   
                            DS.Script.setTextResultScale(pScale);    			
                        } else {
                            var mode = 0;
                            switch(pScale)
                            {
                                case "undef"      :
                                    mode = 0;   // Undeformed
                                    break;
                                    
                                case "undeformed" :
                                    mode = 0;   // Undeformed
                                    break;

                                case "actual"     :
                                    mode = 1;   // Actual
                                    break;

                                case "0.5auto"    :
                                    mode = 2;   // HalfAuto
                                    break;

                                case "auto"       :
                                    mode = 3;   // Automatic
                                    break;

                                case "2auto"      :
                                    mode = 4;   // TwiceAuto
                                    break;

                                case "5auto"      :
                                    mode = 5;   // FiveAuto
                                    break;

                                default :
                                    return;
                            }
                            DS.Script.setResultScale(mode);
                        }                       
                    }
                }
                DS.Graphics.StreamMode = 0; 
            }
        '''
    
        jsmain = 'setScale({});'.format(arg)
        
        if ignore_js_err: jscode = jsfun + self._try_wrapper_js(jsmain)
        else: jscode = jsfun + jsmain
         
        try:
            self._send_js_macro(container, jscode, module, visible=True)
        except Exception as err_msg:
            self._log_('An error occured!')
            self._log_(err_msg, 1) 
    # -------------------------------------------------------------------- 
    def send_act_macro(self, sys, code, ext='py', comp='Model'): 
        """
        !!!!THIS IS WiP AND NOT FUNCTIONAL!!!!
        Sends Python/JScript code to Mechanical in-build macro executor.
        Arg:
            sys: str; specify a Mechanical system, e.g. 'SYS'
            code: str; code to execute
            ext: str; macro extention
            comp: str; module to open          
        """
        
        if ext == 'py': logext = 'Python'
        elif ext == 'js': logext = 'JScript'
        else: logext = ""
        logext = logext if logext else ext
                    
        self.__macros_count += 1
        tempdir = os.path.join(os.getcwd(), self._macro_def_dir)
        if not os.path.exists(tempdir): 
            os.makedirs(tempdir)
            self.__macro_dir_path = tempdir
        
        tempfile = '{}_Script.{}'.format(self.__macros_count, ext)       
        tempfile = os.path.join(tempdir, tempfile)
               
        with open(tempfile, 'w') as f:
            self._log_('"send_act_macro()" METHOD IS A PLACEHOLDER')
            return None              
               
        self.send_act_macfile(sys, tempfile, comp, ignore_js_err=False)       
        
    # --------------------------------------------------------------------     
    def send_act_macfile(self, mech_sys, filename, mech_comp='Model', ignore_js_err=False): 
        """
        Executes a macro file using Mechanical in-build macro executor.
        
        Arg:
            mech_sys: str; specify a Mechanical system, e.g. 'SYS'
            filename: str; macro file
            mech_comp: str; module to open
            ignore_js_err: bool; wraps js main cammands in a try block
        """
        try:
            ext = os.path.basename(filename).split('.')[1]
        except:
            self._log_('Error: Could not determine a file extention!', 1)
            return None
        else:
        
            if ext == 'py': logext = 'Python'
            elif ext == 'js': logext = 'JScript'
            else: logext = ""
            logext = logext if logext else ext
                
            jsfilename =  self._winpath_js(os.path.dirname(filename)) + os.path.basename(filename)
             
            self._log_('Sending ACT {} macro...'.format(logext)) 
            self._log_('File: {}'.format(filename)) 
            
            jsfun = ''       
            jsmain = 'DS.Script.doToolsRunMacro("{}");'.format(jsfilename) 
            
            if ignore_js_err: jscode = jsfun + self._try_wrapper_js(jsmain)
            else: jscode = jsfun + jsmain
            
            try:
                self._send_js_macro(mech_sys, jscode, mech_comp, visible=True)
            except Exception as err_msg:
                self._log_('An error occured!')
                self._log_(err_msg, 1) 
            else:
                self._log_('Macro execution finished', 1)
    
    @staticmethod
    def __jsfun_savepics():
        return '''
            function saveObjectsPictures(clsidObj, activeObjs, pdir, pName, pPref, pFit, imode){                              
                var numObjs = activeObjs.Count;
                var image = DS.Graphics.ImageCaptureControl;
                var clsidEnv = 105; // load cases
                
                switch (imode) 
                {
                    case 0:
                        pExt = '.png';
                        break;
                    case 1:
                        pExt = '.jpeg';
                        break;
                }
                
                var cntFigures = 0;
                
                for (var i = 1; i <= numObjs; i++) {
                    var objActive = activeObjs.Item(i);

                    if (objActive && objActive.ID && (objActive.Class == clsidObj))
                    {
                        DS.Graphics.Draw2(objActive.ID);
                        
                        if (pFit) {
                            DS.Graphics.setisoview(7);
                            DS.Script.doGraphicsFit();                           
                        }
                        
                        if (!pName) {
                            var picEnum = ("00" + cntFigures).slice (-3);
                            var nameParent = (objActive.Parent.Name).replace(/ |_/g, '-');
                            var nameFigure = (objActive.Name).replace(/ |_/g, '-');
                            
                            try {
                                var objSearch = objActive.Parent;
                                while (objSearch.Class != clsidEnv) objSearch = objSearch.Parent;                         
                                var nameSolution = (objSearch.Name).replace(/ |_/g, '-');
                                nameSolution = nameSolution + "_";
                            } catch (err) {
                                var nameSolution = "";
                            }
                            nameFull = (pPref + "_" + picEnum + "_" + nameSolution + nameParent + "_" + nameFigure);
                        } else {   
                            nameFull = pName;
                        } 
                        // DS.Graphics.Redraw(1);
                        image.Write(imode, pdir + nameFull + pExt);                      
                        cntFigures++;
                        if (pName) break;
                    }
                }
            }
            function prepocPicOutput(pHeight, pWidth, pFontFactor, col5, col6){
                var gr_IMAGE2FILE = 0x16;
                var gr_ImgResEnhanced = 0x1C;
                var gr_FONTMAGFACTOR = 0x0;
                
                DS.Graphics.Info(gr_IMAGE2FILE) = -1;
                DS.Graphics.GfxUtility.Legend.IsFontSizeCustomized = -1;
                DS.Graphics.GfxUtility.Legend.iSImgResEnhanced = -1;
                DS.Graphics.Info(gr_ImgResEnhanced) = -1;
                DS.Graphics.InfoDouble(gr_FONTMAGFACTOR) = pFontFactor; 
                if ((pHeight > 0) && (pWidth > 0)) {
                    DS.Graphics.MemStreamHeight = pHeight;
                    DS.Graphics.MemStreamWidth = pWidth;
                }
                
                DS.Graphics.StreamMode = 1; 
                
                DS.Graphics.Scene.Color(1) = 0x00ffffff; //hex value for white - this will blend in with the ANSYS logo,
                DS.Graphics.Scene.Color(2) = 0x00ffffff; //making the logo impossible to see.
                DS.Graphics.Scene.Color(5) = col5; // HACK (restore color 5)
                DS.Graphics.Scene.Color(6) = col6; // HACK (restore color 6)
                
            }
            function postPicOutput(prevColor1, prevColor2, prevColor5, prevColor6, prevLegend, prevRuler, prevTriad, prevRandom) 
            {
                var gr_IMAGE2FILE = 0x16;
                var gr_ImgResEnhanced = 0x1C;
                var gr_FONTMAGFACTOR = 0x0;
                
                DS.Graphics.Scene.Color(1) = prevColor1;
                DS.Graphics.Scene.Color(2) = prevColor2;
                DS.Graphics.Scene.Color(5) = prevColor5;
                DS.Graphics.Scene.Color(6) = prevColor6;
                DS.Graphics.LegendVisibility = prevLegend;
                DS.Graphics.RulerVisibility = prevRuler;
                DS.Graphics.TriadOn = prevTriad;
                DS.Graphics.RandomColors = prevRandom;
                             
                DS.Graphics.Info(gr_IMAGE2FILE) = 0;
                DS.Graphics.GfxUtility.Legend.IsFontSizeCustomized = 0;
                DS.Graphics.GfxUtility.Legend.IsImgResEnhanced = 0;
                DS.Graphics.Info(gr_ImgResEnhanced) = 0;
                DS.Graphics.StreamMode = 0;  //so the geometry view will become visible again                                  
            }           
        '''
    # ---------------------------------------------------------------
    # Private methods
    # ---------------------------------------------------------------
    def _send_js_macro(self, sys, code, comp='Model', visible=False):
        """
        Executes JS macro in Mechanical. This method is used for all interactions with Mechanical.
        
        Note: when executing JS via SendCommand() 'DS.' namespcae is not availables.
        This method replaces all references to 'DS' with 'WB.AppletList.Applet("DSApplet").App'
        which will allow JS macro execution. This method does NOT support macro from file!
        Try to pass 'DS.Script.doToolsRunMacro()' with appropriate filename. 
        
        Arg:
            sys: str; specify a Mechanical system, e.g. 'SYS'
            code: str; JS macro string
            comp: str; module to open
        """
    
        # filepath = workbench.GetUserFilesDirectory()
        # filename = os.path.join(filepath,'wb_mac.js')
        # with open(filename, 'w') as f:
            # f.write(code)
        self._log_('Running JScript at -> System: "{}", Component: "{}"'.format(sys, comp))
        
        ds_space = 'WB.AppletList.Applet("DSApplet").App.'
        code = code.replace('DS.', ds_space) 
        
        try:                      
            system = workbench.GetSystem(Name=sys)
            model = system.GetContainer(ComponentName=comp)           
            model.Edit(Interactive=visible)
            model.SendCommand(Command=code)           
            model.Exit()       
        except Exception as err_msg:
            self._log_('An error occured!')
            self._log_(err_msg, 1) 
        else:
            self._log_('Finished', 1)
    
    def _clear_DPs(self):
        
        dps = self._get_DPs()
        for dp in dps:
            try:
                dp.Delete()
            except:
                pass
            
        self.__DPs = self._get_DPs()
        self.__DPs_present = len(self.__DPs)
    # --------------------------------------------------------------------     
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
    # --------------------------------------------------------------------     
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
    # -------------------------------------------------------------------- 
    def _output_group_by_DPs(self):
        res = [[0]*len(self._param_out_value) for x in xrange(self.__DPs_present)]
        for i in xrange(self.__DPs_present):
            for j, key in enumerate(self._param_out):
                res[i][j] = self._param_out_value[key][i]
        return res
    # -------------------------------------------------------------------- 
    def _save_project(self):
        workbench.Save(Overwrite=True)
        self._log_('Project Saved', 1)
    # --------------------------------------------------------------------     
    def _set_parameter(self, dp, name, value):
        dp.SetParameterExpression(Parameter=self._get_parameter(name), Expression=value)			
        
    def _get_parameter_value(self, dp, name):
        return str(dp.GetParameterValue(self._get_parameter(name)).Value)
    # -------------------------------------------------------------------- 
    def _get_DPs(self):
        """Get Design Points list from project."""
        if self.__active:
            return workbench.Parameters.GetAllDesignPoints()
        else:
            self._log_('Cannot get Design Points: No active project found!', 1)
            raise NoActiveProjectFound
    # --------------------------------------------------------------------         
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
    def copy_files(template, source_dir, target_dir):
        srch_template = os.path.join(source_dir, template)
        try:
            srch = [f for f in glob(srch_template)]
            f = srch[0]
        except:
            pass
        else:
            if not os.path.exists(target_dir): os.makedirs(target_dir)
            for f in srch:
                filename = os.path.basename(f)
                shutil.copyfile(f, os.path.join(target_dir, filename))
                
    @staticmethod            
    def move_files(template, source_dir, target_dir):
        srch_template = os.path.join(source_dir, template)
        try:
            srch = [f for f in glob(srch_template)]
            f = srch[0]
        except:
            pass
        else:
            if not os.path.exists(target_dir): os.makedirs(target_dir)
            for f in srch:
                filename = os.path.basename(f)
                os.rename(f, os.path.join(target_dir, filename))
                
    @staticmethod
    def _listify(inp):	
        """Returns list of 1 items if input is not a list"""
        return inp if isinstance(inp, list) else [inp]
    
    @staticmethod
    def _get_parameter(name):
        return workbench.Parameters.GetParameter(Name=name)	
    # ---------------------------------------------------------------   
    @staticmethod
    def _try_wrapper_js(code):
        return '''
            try {
                %s
            } catch (err) {
            
            }       
        ''' % code
    # ---------------------------------------------------------------
    @staticmethod
    def _winpath_js(dirpath):
        """
        Make all back slashes into double
        """
        return os.path.join(dirpath, '').replace('\\', '\\\\')
    # ---------------------------------------------------------------    
    @staticmethod
    def _bool_js(value):
        return 'true' if value else 'false'

#__________________________________________________________
class AsyncLogChecker(object):
    """
    .NET
    Class used for pulling info from APDL solver log files
    Does not support concurrent APDL solvers!
    """
    __version__ = '0.0.1'
    
    def __init__(self, outfile, watch_dir, watchfile='solve.out', sec=1):
        self._logger = Logger('async.txt')    
        log_prefix = '{}||'.format(str(self.__class__.__name__))
        self._log_ = partial(self._logger.log, info=log_prefix) 
        
        self.outfile = outfile
        self.dir = watch_dir
        self.wait = sec*1000       
        self.watchfile = watchfile
        self.current_file = ''
        self.current_position = 0
             
        self.is_solving = False     
        with open(self.outfile, 'w') as f:
            pass
         
    def main(self):
        def doevents():
            self._log_('Invoke log update watcher', 1)
            print('Invoke log update watcher')
            while self.is_solving: 
                Thread.Sleep(self.wait)
                if not os.path.exists(self.current_file): self.current_file = ''
                if not self.current_file:
                    if self.current_position:
                        self._log_('Searching for new log file...')
                        print('Searching for new log file...')
                    self.current_position = 0
                    
                    file_list = self.reglob(self.dir, self.watchfile) 
                    try: file = file_list[0]
                    except: continue
                    
                    self.current_file = file
                    self._log_('Found {}'.format(self.current_file))  
                    print('Found {}'.format(self.current_file))  
                    
                with open(self.current_file, 'r') as f, open(self.outfile, 'a') as g:
                    f.seek(self.current_position)
                    newdata = f.read()
                    self.current_position = f.tell()
                    if newdata and newdata != '\n': 
                        g.write(newdata)
                        self._log_('File {} updated ({})'.format(self.outfile, len(newdata)))
                        print('File {} updated ({})'.format(self.outfile, len(newdata)))
                        
        self._log_('Start watching {} for {} every {} sec'.format(self.dir, self.watchfile, self.wait/1000), 1)
        print('Start watching {} for {} every {} sec'.format(self.dir, self.watchfile, self.wait/1000))
        try: doevents()
        finally: self._log_('Finished watching', 1); print('Finished watching')
        
    @staticmethod
    def reglob(dir, srch):
        file_list = []
        for i in os.walk(dir):
            c_dir = i[0]
            c_template = os.path.join(c_dir, srch)
            s = [f for f in glob(c_template)]
            try: s = s[0]
            except: continue
            if s: file_list.append(s)
        return file_list
            
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