import xlwings as xw
import win32ui
import win32con
import time


wb = xw.books.active
sheet = xw.sheets.active

def select_totals():
	val = win32ui.MessageBox('Select design column?', 'Select Box Totals', win32con.MB_YESNO)
	if val == win32con.IDYES:
		column = 'E'
	else:
		column = 'H'
	test = None
	for x in sheet.api.UsedRange.Columns(column).cells:
		if x.Font.Bold == True and x.Mergecells == False:
			if test == None:
				test = x
			else:
				test = xw.apps.active.api.Union(test, x)
	if test != None:
		test.select