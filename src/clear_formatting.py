import xlwings as xw

def clear():
	xw.apps.active.selection.api.FormatConditions.Delete()