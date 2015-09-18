#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: merge2.py,v 1.3 2004/01/31 18:56:07 fufff Exp $"""

###############################################################################
#
# Simple example of merging cells using the Spreadsheet::WriteExcel module
#
# This example shows how to merge two or more cells with formatting.
#
# reverse('(c)'), August 2002, John McNamara, jmcnamara@cpan.org
#

import pyXLWriter as xl

# Create a new workbook and add a worksheet
workbook  = xl.Writer("merge2.xls")
worksheet = workbook.add_worksheet()

# Increase the cell size of the merged cells to highlight the formatting.
worksheet.set_column([1, 2], 20)
worksheet.set_row(2, 30)

# Create a merged format
format = workbook.add_format(merge = 1,
                                bold = 1,
                                size = 15,
                                pattern = 1,
                                border = 6,
                                color = 'white',
                                fg_color = 'green',
                                border_color = 'yellow',
                                align = 'vcenter'
                               )

# Only one cell should contain text, the others should be blank.
worksheet.write([2, 1], "Merged Cells", format)
worksheet.write_blank([2, 2], format)

workbook.close()