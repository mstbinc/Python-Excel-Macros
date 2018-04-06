import alphabet
import xlwings as xw
from xlwings import constants



sheet_start_row = 15
sheet_end_row = 43



def generate_outlets():
    wb = xw.books.active
    data = xw.apps.active.selection
    sheet = xw.sheets.active

    sum_total = 0
    equipment = []

    i = 16
    for r in data.rows:

    	tag = r.value[0]
    	amount = r.value[1]

    	if amount > 1:
    		sum_total += (amount + 2)
    	elif amount == 1:
    		sum_total += 2

        equipment.append([tag, amount])

    	i += 1

    # Take off last additional line because nothing is after it
    sum_total -= 1

    xw.Range("E15").value = ""
    # Save size formula in temporary location
    xw.Range("D15:I15").api.Copy()
    xw.Range("N1").select()
    sheet.api.Paste()

    xw.Range("A42:I43").api.Copy()
    xw.Range("A15").select()
    sheet.api.Paste()

    xw.Range("M30").value = sum_total

    xw.Range("A43:I43").api.Copy()

    # insert the necessary number of rows for all systems
    while sum_total > (sheet_end_row - sheet_start_row + 1):
        xw.Range("A43:I43").api.Insert(constants.InsertShiftDirection.xlShiftDown)
        sum_total -= 1

    xw.Range("Q20").value = alphabet.get_alphabet()


    # Draw each piece of equipment
    counts = {}
    current_row = sheet_start_row

    num = 1

    for e in equipment:
        tag = e[0]
        amount = e[1]

        # label equipment
        xw.Range("A" + str(current_row)).value = tag

        start_row = current_row

        # create entries
        for row in xrange(0, int(amount)):

            #number
            xw.Range("N1:S1").api.Copy()
            xw.Range("D{}".format(current_row)).select()
            sheet.api.Paste()

            xw.Range("B{}".format(current_row)).value = num

            num += 1
            current_row += 1

        # TOTAL row
        if amount > 1:
            xw.Range("N1:S1").api.Copy()
            xw.Range("D{}".format(current_row)).select()
            sheet.api.Paste()
            xw.Range("C{0}:D{0}".format(current_row)).api.Merge()
            xw.Range("C{0}:D{0}".format(current_row)).value = "TOTAL:"
            xw.Range("C{0}:D{0}".format(current_row)).api.Font.Bold = True
            # cfm sums
            for col in ['E', 'F', 'H']:
                xw.Range("{}{}".format(col, current_row)).formula = '=SUM({0}{1}:{0}{2})'.format(col, start_row, current_row - 1)

        else:
            current_row -= 1
            xw.Range("E{}".format(current_row)).value = 0
            xw.Range("F{}".format(current_row)).value = 0
            xw.Range("H{}".format(current_row)).value = 0

        # bold total row
        xw.Range("E{}".format(current_row)).api.Font.Bold = True
        xw.Range("G{0}:I{0}".format(current_row)).api.Font.Bold = True

        current_row += 2


    xw.Range("N1:S1").clear()
    xw.apps.active.api.ActiveWindow.ScrollRow = 1
    xw.Range("C15").select()