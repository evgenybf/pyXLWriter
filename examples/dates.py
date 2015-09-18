#!/usr/bin/env python

__revision__ = """$Id: dates.py,v 1.2 2004/07/01 17:14:18 fufff Exp $"""

###############################################################################
#
# Example of how to using the datetime and mx.DateTime modules.
#

import pyXLWriter as xl

try:
    import datetime as dt       # Needs python 2.3
except ImportError:
    dt = None    

try:
    import mx.DateTime as mxdt
except ImportError:
    mxdt = None        

# Create a new workbook and add a worksheet
wb = xl.Writer("dates.xls")
ws = wb.add_worksheet()

# Expand the first column so that the date is visible.
ws.set_column("B:B", 25)

# Add a format for the date
date_format =  wb.add_format()
date_format.assign(wb.get_default_datetime_format())
date_format.set_color("red")

section_format = wb.add_format()
section_format.set_bold(True)

############################
# Write some dates and times

row = 0

if dt:
    ws.write((row, 0), "datetime", format=section_format)
    row += 1
    
    ws.write((row, 0), "datetime:")
    ws.write_date((row, 1), dt.datetime(1979, 8, 30, 18, 30, 33), date_format)
    row += 1

    ws.write((row, 0), "date")
    ws.write_date((row, 1), dt.datetime.now().date())
    row += 1

    # You can use write method
    ws.write((row, 0), "time:")
    ws.write((row, 1), dt.datetime.now().time())
    row += 1

if mxdt:
    if row > 0:
        row += 1
    ws.write((row, 0), "mx.DateTime", format=section_format)
    row += 1
    
    ws.write((row, 0), "datetime:")
    ws.write((row, 1), mxdt.DateTime(1979, 8, 30, 18, 30, 33))
    row += 1

    
# Workbooke must be closed!
wb.close()


