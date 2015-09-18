#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: tab2xls.py,v 1.1 2004/01/31 18:57:53 fufff Exp $"""

######################################################################
#
# Example of how to use the WriteExcel module
#
# The following converts a tab separated file into an Excel file
#
# Usage: tab2xls.pl tabfile.txt newfile.xls
#
# reverse('(c)'), March 2001, John McNamara, jmcnamara@cpan.org
#

import sys
import pyXLWriter as xl

# Check for valid number of arguments
if len(sys.argv) != 3:
    print "Usage: tab2xls tabfile.txt newfile.xls"
    sys.exit(1)
else:
    tabfilename, xlsfilename = sys.argv[1:3]

# Open the tab delimited file
tabfile = file(tabfilename, "r")

# Create a new Excel workbook
workbook = xl.Writer(xlsfilename)
worksheet = workbook.add_worksheet()

# Row and column are zero indexed
nrow = 0

for line in tabfile.xreadlines():
    # Split on single tab
    row = line.split('\t')
    for ncol, cell in enumerate(row):
        worksheet.write([nrow, ncol], cell.strip())
    nrow += 1

tabfile.close()
workbook.close()