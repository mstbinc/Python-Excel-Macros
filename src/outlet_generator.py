import xlwings as xw

# The number row where the sheet ends.
# Used for calculating how many lines need to be inserted.
sheet_end_row = 43


def generate_outlets():
    wb = xw.books.active
    data = xw.apps.active.selection
    sheet = xw.sheets.active

    sum_total = 0

    i = 16
    for r in data.rows:

    	tag = r.value[0]
    	amount = r.value[1]

    	if amount > 1:
    		sum_total += (amount + 2)
    	elif amount == 1:
    		sum_total += 2

    	xw.Range("Q" + str(i)).value = [tag, amount]
    	i += 1

    # Take off last additional line because nothing is after it
    sum_total -= 1

    # Save size formula in temporary location
    xw.Range("D15").api.Copy()
    xw.Range("N1").select()
    sheet.api.Paste()

    xw.Range("A42:I43").api.Copy()
    xw.Range("A15").select()
    sheet.api.Paste()

    xw.Range("M30").value = sum_total

    #xw.Range("N1").clear()