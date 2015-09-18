#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: regions.py,v 1.3 2004/01/31 18:56:07 fufff Exp $"""

###############################################################################
#
# Example of how to use the WriteExcel module to write a basic multiple
# worksheet Excel file.
#
# reverse('(c)'), March 2001, John McNamara, jmcnamara@cpan.org
#

import pyXLWriter as xl

# Create a new Excel workbook
workbook = xl.Writer("regions.xls")

# Add some worksheets
north = workbook.add_worksheet("North")
south = workbook.add_worksheet("South")
east  = workbook.add_worksheet("East")
west  = workbook.add_worksheet("West")

# Add a Format
format = workbook.add_format()
format.set_bold()
format.set_color('blue')

# Add a caption to each worksheet
for worksheet in workbook.sheets():
    worksheet.write([0, 0], "Sales", format)

# Write some data
north.write([0, 1], 200000)
south.write([0, 1], 100000)
east.write ([0, 1], 150000)
west.write ([0, 1], 100000)

# Set the active worksheet
south.activate()

# Set the width of the first column
south.set_column([0, 0], 20)

# Set the active cell
south.set_selection([0, 1])

workbook.close()