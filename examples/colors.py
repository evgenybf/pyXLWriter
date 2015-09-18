#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: colors.py,v 1.3 2004/01/31 18:56:07 fufff Exp $"""

######################################################################
#
# Demonstrates Spreadsheet::WriteExcel's named colors and the Excel
# color palette.
#
# reverse('(c)'), March 2002, John McNamara, jmcnamara@cpan.org
#

import pyXLWriter as xl

workbook = xl.Writer("colors.xls")

# Some common formats
center  = workbook.add_format(align = 'center')
heading = workbook.add_format(align = 'center', bold = 1)

# Try this to see the default Excel 5 palette
# workbook.set_palette_xl5()

######################################################################
#
# Demonstrate the named colors.
#
colors = {0x08: 'black',
          0x0C: 'blue',
          0x10: 'brown',
          0x0F: 'cyan',
          0x17: 'gray',
          0x11: 'green',
          0x0B: 'lime',
          0x0E: 'magenta',
          0x12: 'navy',
          0x35: 'orange',
          0x14: 'purple',
          0x0A: 'red',
          0x16: 'silver',
          0x09: 'white',
          0x0D: 'yellow',
         }

worksheet1 = workbook.add_worksheet('Named colors')

worksheet1.set_column([0, 3], 15)

worksheet1.write([0, 0], "Index", heading)
worksheet1.write([0, 1], "Index", heading)
worksheet1.write([0, 2], "Name",  heading)
worksheet1.write([0, 3], "Color", heading)

i = 1

for index, color in colors.items():
    format = workbook.add_format(fg_color = color,
                                 pattern = 1,
                                 border = 1
                                )
    worksheet1.write([i+1, 0], index, center)
    worksheet1.write([i+1, 1], "0x%02X" %index, center)
    worksheet1.write([i+1, 2], color, center)
    worksheet1.write([i+1, 3], '', format)
    i += 1

######################################################################
#
# Demonstrate the standard Excel colors in the range 8..63.
#

worksheet2 = workbook.add_worksheet('Standard colors')

worksheet2.set_column([0, 3], 15)

worksheet2.write([0, 0], "Index", heading)
worksheet2.write([0, 1], "Index", heading)
worksheet2.write([0, 2], "Color", heading)
worksheet2.write([0, 3], "Name",  heading)

for i in xrange(8, 63+1):
    format = workbook.add_format(fg_color = i,
                                 pattern = 1,
                                 border = 1
                                )
    worksheet2.write([i-7, 0], i, center)
    worksheet2.write([i-7, 1], "0x%02X" % i, center)
    worksheet2.write([i-7, 2], '', format)
    # Add the  color names
    if colors.has_key(i):
        worksheet2.write([i-7, 3], colors[i], center)
        
workbook.close()        