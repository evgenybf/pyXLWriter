#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: long_string.py,v 1.3 2004/01/31 18:56:07 fufff Exp $"""

######################################################################
#
# This is a example of how to work around Spreadsheet::WriteExcel and
# Excel5's 255 character string limitation using a formula to create
# a long string from shorter strings.
#
# For genuine long strings see the following link for information 
# about the Excel97 pre-release version of this module:
# http://freshmeat.net/projects/writeexcel/#comment-24916
#
# reverse('(c)'), April 2002, John McNamara, jmcnamara@cpan.org
#

import string
import pyXLWriter as xl

workbook  = xl.Writer("long_string.xls")
worksheet = workbook.add_worksheet()

######################################################################
#
# long_string(str)
#
# Converts long strings into an Excel string concatenation formula.
# The concatenation is inserted between words to improve legibility.
#
# returns: An Excel formula if string is longer than 255 chars.
#          The unmodified string otherwise.
#
def long_string(str):
    limit = 255
    # Return short strings
    if len(str) <= limit:
        return str 
    # Split the line at word boundaries where possible
    
    segments = [str[0:limit]]
    i_prev = 0
    for i in xrange(limit, len(str), limit):
        segments.append(str[i_prev:i])
        i_prev=i
    # Join the string back together with quotes and Excel concatenation
    str = string.join(segments, '"&"')
    # Add formatting to convert the string to a formula string
    return '="' + str + '"'

#
# The following formatting is optional.
#
wrap = workbook.add_format(text_wrap=1, valign='top')
worksheet.set_column('B:B', 50)
worksheet.set_row(1, 170)
worksheet.set_row(2, 170)

# Example 1
#
# Create a long string using the Excel concatenation operator "&".
# The formula has the format '="String1" & "String1" & ...'
#
str1 = 'these are the days and '*10
str2 = '="' + str1 + '"&"' + str1 + '"&"' + str1 + 'stop."'

worksheet.write('B2', str2, wrap)

# Example 2
#
# Create a long string using the Excel concatenation operator "&".
# The methodology is the same as the previous example except that
# we use a function to insert the formatting.
#
str3 = ('these are the days and ' * 30) + 'stop."'

worksheet.write('B3', long_string(str3), wrap)

# Leaves shorter strings unmodified
worksheet.write('B4', long_string("hello, world"))

workbook.close()