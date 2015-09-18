#!/usr/bin/env python
__revision__ = """$Id: test_Workbook.py,v 1.10 2004/04/14 11:06:37 fufff Exp $"""

import os, os.path
import unittest

import testsupport
from pyXLWriter.Workbook import Workbook
from testsupport import is_equal_files


class WorkbookTest(unittest.TestCase):

    filename = "test.xls"
    perldir = "perl_sample/"
    
    def setUp(self):
        self.wb = Workbook(self.filename)

    def tearDown(self):
        if self.wb:
            self.wb.close()
            self.wb = None
        if os.path.exists(self.filename):
            os.remove(self.filename)

    def _compare_files(self, filename1, filename2):
        if not is_equal_files(filename1, filename2):
            self.fail("Files are not equals")

    def test_add_worksheet(self):
        ws = self.wb.add_worksheet()

    def test_calc_sheet_offsets(self):
        ws = self.wb.add_worksheet()
        self.wb._calc_sheet_offsets()

    def test_add_format(self):
        self.wb.add_format()
        self.assertEqual(5, len(self.wb._formats), "Bad number of formats")

    def test_etalon_worksheet(self):
        wb = self.wb
        perl_filename = os.path.join(self.perldir, "worksheet.xls")
        ws = wb.add_worksheet()
        self.wb.close()
        self._compare_files(self.filename, perl_filename)

    def test_etalon_write(self):
        wb = self.wb
        perl_filename = os.path.join(self.perldir, "write.xls")
        ws1 = wb.add_worksheet("List1")
        ws2 = wb.add_worksheet("List2")
        ws1.write("AA4", "AA4!")
        ws2.write("C1", "C1!")
        ws1.write("A3", "!A3!")
        self.wb.close()
        self._compare_files(self.filename, perl_filename)

    def test_etalon_selection(self):
        wb = self.wb
        perl_filename = os.path.join(self.perldir, "selection.xls")
        ws = wb.add_worksheet("Sheet 1")
        ws.set_selection([4, 6, 10, 100])
        wb.close()
        self._compare_files(self.filename, perl_filename)


class WorkbookBigTest(unittest.TestCase):
    
    filename = "test.xls"
    perldir = "perl_sample/"
    
    def setUp(self):
        self.wb = Workbook(self.filename, big=True)

    def tearDown(self):
        if self.wb:
            self.wb.close()
            self.wb = None
        if os.path.exists(self.filename):
            os.remove(self.filename)

    def _compare_files(self, filename1, filename2):
        if not is_equal_files(filename1, filename2):
            self.fail("Files are not equals")

    def test_addworksheet(self):
        self.wb.add_worksheet()
        self.wb.close()
        self.wb = None

    def test_write(self):
        ws = self.wb.add_worksheet()
        ws.write((0, 0), "(0, 0)")
        self.wb.close()
        self.wb = None

    def test_etalon_write_small_big(self):
        ws = self.wb.add_worksheet()
        ws.write("A1", "Used OLEWriterBig!")
        self.wb.close()
        self.wb = None
        perl_filename = os.path.join(self.perldir, "small_big.xls")
        self._compare_files(self.filename, perl_filename)

    def test_etalon_write_big(self):
        ws = self.wb.add_worksheet()
        ws.set_column((0, 50), 18)
        for col in xrange(50+1):
            for row in xrange(60+1):
                ws.write((row, col), "Row: %d Col: %d" % (row, col))
        self.wb.close()
        self.wb = None
        perl_filename = os.path.join(self.perldir, "big.xls")
        self._compare_files(self.filename, perl_filename)


if __name__ == "__main__":
    unittest.main()