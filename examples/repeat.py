#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: repeat.py,v 1.4 2004/01/31 18:56:07 fufff Exp $"""

######################################################################
#
# Example of writing repeated formulas.
#
# reverse('(c)'), August 2002, John McNamara, jmcnamara@cpan.org
#

import time
import pyXLWriter as xl

workbook  = xl.Writer("repeat.xls")
worksheet = workbook.add_worksheet()

limit = 1000

# Write a column of numbers
for row in xrange(limit+1):
    worksheet.write([row, 0],  row)

# Store a fomula
formula = worksheet.store_formula('=A1*5+4')

# Write a column of formulas based on the stored formula
t = time.time()
for row in xrange(limit+1):
    worksheet.repeat_formula([row, 1], formula, patterns=[('A1', 'A'+str(row+1))])
print "%.2f" % (time.time() - t)

# Direct formula writing. As a speed comparison uncomment the
# following and run the program again
t = time.time()
for row in xrange(limit+1):
    worksheet.write_formula([row, 2], '=A' + str(row+1) + '*5+4')
print "%.2f" % (time.time() - t)

workbook.close()