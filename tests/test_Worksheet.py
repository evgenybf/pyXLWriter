#!/usr/bin/env python
__revision__ = """$Id: test_Worksheet.py,v 1.12 2004/08/17 07:22:38 fufff Exp $"""

import os, os.path
import unittest

import testsupport
from testsupport import read_file
from pyXLWriter.Worksheet import Worksheet
from pyXLWriter.Format import Format


class WorksheetTest(unittest.TestCase):
    
    def setUp(self):
        self.ws = Worksheet("test", None, 0)
        self.format = Format(color="green")

    def tearDown(self):
        self.ws = None

    def test_methods_no_error(self):
        self.ws.write([0, 1], None)
        self.ws.write((0, 2), "Hello")
        self.ws.write((0, 3), 888)
        self.ws.write([0, 4], 888L)
        self.ws.write_row((0, 0), [])
        self.ws.write_row((0, 0), ["one", "two", "three"])
        self.ws.write_row((0, 0), [1, 2, 3])
        self.ws.write_col((0, 0), [])
        self.ws.write_col((0, 0), ["one", "two", "three"])
        self.ws.write_col((0, 0), [1, 2, 3])
        self.ws.write_blank((0, 0), [])

    def test_store_dimensions(self):
        self.ws._store_dimensions()
        datasize = self.ws._datasize
        self.assertEqual(14, datasize)

    def test_store_window2(self):
        self.ws._store_window2()
        datasize = self.ws._datasize
        self.assertEqual(14, datasize) 

    def test_store_selection(self):
        self.ws._store_selection(0, 0, 0, 0)
        datasize = self.ws._datasize
        self.assertEqual(19, datasize)

    def test_store_colinfo_output(self):
        self.ws._store_colinfo()
        datasize = self.ws._datasize
        self.assertEqual(15, datasize) 

    def test_format_row(self):
        self.ws.set_row(3, None)
        self.ws.set_row(2, 10)
        self.ws.set_row(1, 25, self.format)
        self.ws.set_row(4, None, self.format) # [4, 6]

    def test_format_column(self):
        self.ws.set_column(3, None)
        self.ws.set_column(2, 10)
        self.ws.set_column(1, 25, self.format)
        self.ws.set_column([4, 6], None, self.format)

    def test_process_cell(self):
        rc = self.ws._process_cell("C80")
        self.assertEqual((79, 2), rc)
        rc = self.ws._process_cell([63, 84])
        self.assertEqual((63, 84), rc)

    def test_process_cellrange(self):
        rc = self.ws._process_cellrange("C80")
        self.assertEqual((79, 2, 79, 2), rc)
        rc = self.ws._process_cellrange((7, 6))
        self.assertEqual((7, 6, 7, 6), rc)

    def test_process_rowrange(self):
        rc = self.ws._process_rowrange(6)
        self.assertEqual((6, 0, 6, -1), rc)
        rc = self.ws._process_rowrange((5, 6))
        self.assertEqual((5, 0, 6, -1), rc)
        rc = self.ws._process_rowrange("4:6")
        self.assertEqual((3, 0, 5, -1), rc)

    def test_process_colrange(self):
        rc = self.ws._process_colrange(5)
        self.assertEqual((0, 5, -1, 5), rc)
        rc = self.ws._process_colrange((5, 6))
        self.assertEqual((0, 5, -1, 6), rc)
        rc = self.ws._process_colrange("D:E")
        self.assertEqual((0, 3, -1, 4), rc)


if __name__ == "__main__":
    unittest.main()
