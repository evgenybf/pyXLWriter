#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: merge5.py,v 1.3 2004/01/31 18:56:07 fufff Exp $"""

###############################################################################
#
# Example of how to use the Spreadsheet::WriteExcel merge_cells() workbook
# method with with complex formatting and rotation.
#
# reverse('(c)'), September 2002, John McNamara, jmcnamara@cpan.org
#

import pyXLWriter as xl

# Create a new workbook and add a worksheet
workbook  = xl.Writer('merge5.xls')
worksheet = workbook.add_worksheet()

# Increase the cell size of the merged cells to highlight the formatting.
for i in xrange(3, 8+1):
    worksheet.set_row(i, 60)
for i in (1, 3, 5):
    worksheet.set_column(i, 15)

###############################################################################
#
# Rotation 1, letters run from top to bottom
#
format1 = workbook.add_format(border = 6,
                              bold = 1,
                              color = 'red',
                              valign = 'vcentre',
                              align = 'centre',
                              rotation = 1
                             )

worksheet.merge_range('B4:B9', 'Rotation 1: Top to bottom', format1)

###############################################################################
#
# Rotation 2, 90 grad anticlockwise
#
format2 = workbook.add_format(border = 6,
                              bold = 1,
                              color = 'red',
                              valign = 'vcentre',
                              align = 'centre',
                              rotation = 2
                             )

worksheet.merge_range('D4:D9', 'Rotation 2: 90\xb0 anticlockwise', format2)

###############################################################################
#
# Rotation 3, 90 grad clockwise
#
format3 = workbook.add_format(border = 6,
                              bold = 1,
                              color = 'red',
                              valign = 'vcentre',
                              align = 'centre',
                              rotation = 3
                             )

worksheet.merge_range('F4:F9', 'Rotation 3: 90\xb0 clockwise', format3)

workbook.close()