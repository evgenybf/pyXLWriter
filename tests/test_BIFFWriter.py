#!/usr/bin/env python
__revision__ = """$Id: test_BIFFWriter.py,v 1.9 2004/02/12 07:08:21 fufff Exp $"""

import unittest

import testsupport
from pyXLWriter.BIFFWriter import BIFFWriter


class BIFFWriterTest(unittest.TestCase):

    def setUp(self):
        self.b = BIFFWriter()

    def tearDown(self):
        self.b = None

    def test_append_no_error(self):
        self.b._append("World")

    def test_prepend_no_error(self):
        self.b._prepend("Hello")

    def test_data_added(self):
        self.b._append("Hello", "World")
        self.assertEqual("HelloWorld", self.b._data)
        self.assertEqual(10, self.b._datasize)

    def test_data_prepended(self):
        self.b._append("Hello")
        self.b._prepend("World")
        self.assertEqual("WorldHello", self.b._data)

    def test_store_bof_length(self):
        self.b._store_bof()
        self.assertEqual(12, self.b._datasize)

    def test_store_eof_length(self):
        self.b._store_eof()
        self.assertEqual(4, self.b._datasize)

    def test_datasize_mixed(self):
        self.b._append("Hello")
        self.b._prepend("World")
        self.b._store_bof()
        self.b._store_eof()
        self.assertEqual(26, self.b._datasize)


if __name__ == "__main__":
    unittest.main()
