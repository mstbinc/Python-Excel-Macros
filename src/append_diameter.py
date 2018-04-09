import xlwings as xw

def diameter():

    wb = xw.books.active
    data = xw.apps.active.selection

    for r in data:
    	if r[0].value is not None:
    		r[0].value = unicode('\x80')