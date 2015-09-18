#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: win32ole.py,v 1.2 2004/02/12 06:11:17 fufff Exp $"""

###############################################################################
#
# This is a simple example of how to create an Excel file using the
# Win32::OLE module.
#
# reverse('(c)'), March 2001, John McNamara, jmcnamara@cpan.org
#

import os, os.path
import win32com.client
from win32com.client import gencache, Dispatch

# For help see testMSOffice.py in pywin32 distribution
# Excel 97
mod = gencache.EnsureModule("{00020813-0000-0000-C000-000000000046}", 0, 1, 2, bForDemand=1)
xlc = mod.constants

xl = Dispatch("Excel.Application.8")    # "Excel.Application"

#xl.Visible = 1

wb = xl.Workbooks.Add()

try:
    ws = wb.Worksheets(1)
    
    ws.Cells(1,1).Value = "Hello World"
    ws.Cells(2,1).Value = "One"
    ws.Cells(3,1).Value = "Two"
    ws.Cells(4,1).Value = 3
    ws.Cells(5,1).Value = 4.0000001
    
    # Add some formatting
    ws.Cells(1,1).Font.Bold = 1
    ws.Cells(1,1).Font.Size = 16
    ws.Cells(1,1).Font.ColorIndex = 3
    ws.Columns("A:A").ColumnWidth = 25
    
    # Write a hyperlink
    range = ws.Range("A7:A7")
    ws.Hyperlinks.Add(Anchor=range, Address="http://www.perl.com/")
    
    # Get current directory 
    dir = os.getcwd()
    
    # Generate Excel file name for saving
    xlsfilename = os.path.join(dir, 'win32ole.xls')
    
    # Save it
    wb.SaveAs(Filename=xlsfilename, FileFormat=xlc.xlNormal) 
    
finally:
    # Close workbook and Excel application
    wb.Close()
