#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: merge4.py,v 1.3 2004/01/31 18:56:07 fufff Exp $"""

###############################################################################
#
# Example of how to use the Spreadsheet::WriteExcel merge_range() workbook
# method with with complex formatting.
#
# reverse('(c)'), September 2002, John McNamara, jmcnamara@cpan.org
#

import pyXLWriter as xl

# Create a new workbook and add a worksheet
workbook  = xl.Writer('merge4.xls')
worksheet = workbook.add_worksheet()

# Increase the cell size of the merged cells to highlight the formatting.
for i in xrange(1, 11+1):
    worksheet.set_row(i, 30) 
worksheet.set_column('B:D', 20)

###############################################################################
#
# Example 1: Text centered vertically and horizontally
#
format1 = workbook.add_format(border = 6,
                              bold = 1,
                              color = 'red',
                              valign = 'vcenter',
                              align = 'center'
                             )

worksheet.merge_range('B2:D3', 'Vertical and horizontal', format1)

###############################################################################
#
# Example 2: Text aligned to the top and left
#
format2 = workbook.add_format(border = 6,
                              bold = 1,
                              color = 'red',
                              valign = 'top',
                              align = 'left'
                             )

worksheet.merge_range('B5:D6', 'Aligned to the top and left', format2)

###############################################################################
#
# Example 3:  Text aligned to the bottom and right
#
format3 = workbook.add_format(border = 6,
                              bold = 1,
                              color = 'red',
                              valign = 'bottom',
                              align = 'right'
                             )

worksheet.merge_range('B8:D9', 'Aligned to the bottom and right', format3)

###############################################################################
#
# Example 4:  Text justified (i.e. wrapped) in the cell
#
format4 = workbook.add_format(border = 6,
                              bold = 1,
                              color = 'red',
                              valign = 'top',
                              align = 'justify'
                             )

worksheet.merge_range('B11:D12', 'Justified: '+ 'so on and '*18, format4)

workbook.close()