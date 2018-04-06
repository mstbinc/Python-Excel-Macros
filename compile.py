"""

Purpose:

Write a script that makes VBA macros to call
RunPython() on the relevant script files.
Assign these macros shortcuts that are defined in
a custom JSON config file.

Output a PERSONAL.xlsm file.
[Optional] place this along with source in the XLSTART Folder


"""

import os
import json
import time
import win32com
from win32com.client import Dispatch
from shutil import copyfile


path = os.getcwd()
config = json.load(open('config.json'))

# Start Excel
xl = Dispatch("Excel.Application")
xl.Visible = True
wb = xl.Workbooks.Add()

src = config['source-path']


sub_template = """' {4}
Sub {3}()
Attribute {3}.VB_ProcData.VB_Invoke_Func = "{1}\\n14"
    RunPython ("import {2}; {2}.{3}()")
End Sub"""

vba_module = ''

for macro in config['macros']:

	string_import = src.replace('/', '.') + macro
	description = config['macros'][macro]['description']
	hotkey = config['macros'][macro]['hotkey']
	entry = config['macros'][macro]['entry']

	vba_sub = sub_template.format(macro, hotkey, string_import, entry, description)
	vba_module += vba_sub

bas_file = open("output.bas", "w")
bas_file.write(vba_module)
bas_file.close()

xlmodule = wb.VBProject.VBComponents.Add(1)
xlmodule.CodeModule.AddFromFile(path + "/output.bas")

wings = wb.VBProject.VBComponents.Add(1)
wings.CodeModule.AddFromFile(path + "/xlwings.txt")

# Save file
wb.SaveAs(path + '/' + config['excel-filename'], FileFormat=52)


# Quit
os.remove("output.bas")
xl.Quit()

appdata = os.getenv('APPDATA') + '\Microsoft\Excel\XLSTART'
os.startfile(appdata)

copyfile('PERSONAL.xlsm', appdata + '/PERSONAL.xlsm')
os.remove("PERSONAL.xlsm")