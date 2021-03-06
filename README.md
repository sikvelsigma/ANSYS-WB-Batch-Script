# ANSYS Workbench - Batch Mode - Python Scripts

**Platform** - Tested on Windows 7, 10

**ANSYS Version** - Tested on 19.5 (2019R3); will most likely work on other versions that are not too old

This is a collection of python scripts I made to make my life easier when dealing with ANSYS Workbench Batch mode. It requres no additional python packages whatsoever. Right now there are 4 modules:

1. *WBInterface.py*
2. *Logger.py*
3. *ExcelFileReader.py*
4. *CSVTable.py*

Ansys Workbench comes with IronPython 2.7 so to run it from batch mode we only need to write a python script, which will control the flow of the project (*run_script.py* as an example here) and a *.bat* file (*run.bat* as an example here).

- Module *WBInterface.py* is the main module which contains a class with all the useful workbench commands. I tried to document it as much as I could. This module is absolutely essential to have.

- Module *Logger.py* contains a class which will create a log file in the project directory and write the flow of the project to it. This module is absolutely essential to have.

- Module *ExcelFileReader.py* is an adapter class to the COM Excel API. Can be used to read tables from excel files. This module can also be used as a stand-alone with IronPython to read data from Excel. You have to have Excel installed on your machine. Not essential

- Module *CSVTable.py* lets us easily import csv file to a list/dict. Not essential.

## How to use 
First of all, this was all made with running it on a remote machine in mind, when you don't have an ability to install additional software or open apps, but have access to a file system. Second of all, in my work I use a specialised hierarchical software which does not support making changes to already existing files, so there was a need to be able to import modules by their modified names (like *WBInterface_100.py* for example). That's why there's this weird system with **exac()** commands implemented to input everything correctly. 
So, the simplest way to use all of this would be:

1. Drop modules *WBInterface.py* and *Logger.py* into the same folder where your ANSYS WB archive/project currently resides
2. Drop *run_script.py* and *run.bat* there also
3. Configure *.bat* file for your machine and run it; format of *.bat* file:

        "<ansysdir>\v<ver>\Framework\bin\Win64\RunWB2.exe" -B -R "run_script.py"


That's it! Easy and simple as I intended. The script will automatically try to find *.wbpz* file and open it or, failing that, it will try to find a *.wbpj* file. After the project was opened, the script will try to issue a global Update command for all DPs (Design Points) present in the project. 
Note: You don't need to delete anything from *run_script.py* since methods won't raise errors if they don't find anything, but it will be logged to a log file.

## Additional functionality
You can also input/output WB parameters with multiple DPs which will managed automatically. All you need to do is to create 2 csv files:

- *.control* or *_control.csv*
- *.input* or *_input.csv*

The first one contain a list of parameters to input and a list of parameters to output. The second one contains values of input parameters with each row being a new DP. Both files support comments (with #).
Example of a *.control* file:
```
  #Input parameters (type 'no' if no inputs)
  P1,P2,P3
  #Output parameters (type 'no' if no outputs)
  P4,P5
```
Example of an *.input* file:
```
  # p1,p2,p3
  501,50,25
  499,49,24
  500,52,26
```
This will generate 3 DPs in our project with 3 input and 2 output parameters.

By default in the project directory a *log.txt* file will be created. Output is written to an *output.txt* file csv-style and Workbench parametric report is saved to a *full_report.txt* file. Of course this is all customizable.

You can also use *CSVTable.py* module or just regular **open()** to read parameters into a list or dict and use **input_by_name()** or **input_by_DPs()** methods of *WBInterface.py* to set them directly.

Note that I'm not a programmer and I apologize in advance for any inconsistencies/bad practises in my code.
