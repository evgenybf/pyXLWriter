#!/usr/bin/env python
__revision__ = """$Id: test_OLEWriter.py,v 1.9 2004/02/08 14:47:03 fufff Exp $"""

import os, os.path
import unittest

import testsupport
from pyXLWriter.OLEWriter import OLEWriter


class OLEWriterTest(unittest.TestCase):

    def setUp(self):
        self.f = "test.ole"
        self.o = OLEWriter(self.f)

    def tearDown(self):
        self.o.close()
        self.o = None  # todo!!!
        if os.path.exists(self.f):
            os.remove(self.f)

    def test_set_size_too_big(self):
         err = "Should have raised MaxSizeError"
         self.assertRaises(Exception, lambda:self.o.set_size(999999999))

    def test_booksize_large(self):
        self.o.set_size(8192)
        self.assertEqual(8192, self.o._booksize)

    def test_booksize_small(self):
        self.o.set_size(2048)
        self.assertEqual(4096, self.o._booksize)

    def test_biffsize(self):
        self.o.set_size(2048)
        self.assertEqual(2048, self.o._biffsize)

    def test_size_allowed(self):
        self.o.set_size(4096)
        self.assertEqual(True, self.o._size_allowed)

    def test_big_block_size_default(self):
        self.o.set_size(4096)
        self.o._calculate_sizes()
        self.assertEqual(8, self.o._big_blocks, "Bad big block size")

    def test_big_block_size_rounded_up(self):
        self.o.set_size(4099)
        self.o._calculate_sizes()
        self.assertEqual(9, self.o._big_blocks, "Bad big block size")

    def test_list_block_size(self):
        self.o.set_size(4096)
        self.o._calculate_sizes()
        self.assertEqual(1, self.o._list_blocks, "Bad list block size")

    def test_root_start_size_default(self):
        self.o.set_size(4096)
        self.o._calculate_sizes()
        self.assertEqual(8, self.o._big_blocks, "Bad root start size")

    def test_root_start_size_rounded_up(self):
        self.o.set_size(4099)
        self.o._calculate_sizes()
        self.assertEqual(9, self.o._big_blocks, "Bad root start size")


if __name__ == "__main__":
    unittest.main()
