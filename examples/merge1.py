#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: merge1.py,v 1.3 2004/01/31 18:56:07 fufff Exp $"""

###############################################################################
#
# Simple example of merging cells using the Spreadsheet::WriteExcel module.
#
# This example shows how to merge two or more cells.
#
# reverse('(c)'), August 2002, John McNamara, jmcnamara@cpan.org
#

import pyXLWriter as xl

# Create a new workbook and add a worksheet
workbook = xl.Writer("merge1.xls")
worksheet = workbook.add_worksheet()

# Increase the cell size of the merged cells to highlight the formatting.
worksheet.set_column('B:D', 20)
worksheet.set_row(2, 30)

# Create a merge format
format = workbook.add_format(merge=1)

# Only one cell should contain text, the others should be blank.
worksheet.write([2, 1], "Merged Cells", format)
worksheet.write_blank([2, 2], format)
worksheet.write_blank([2, 3], format)

workbook.close()