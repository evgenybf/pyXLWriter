#!/usr/bin/env python
__revision__ = """$Id: test_Format.py,v 1.10 2004/04/14 16:50:08 fufff Exp $"""

import unittest
import os, os.path

import testsupport
from testsupport import read_file
from pyXLWriter.Format import Format


class FormatTest(unittest.TestCase):
    
    res_file = "xf_test"
    
    def setUp(self):
        self.format = Format()

    def tearDown(self):
        if os.path.exists(self.res_file):
            os.remove(self.res_file)
        self.format = None

    def test_xf_biff_size(self):
        perl_file = "perl_output/f_xf_biff"
        size = os.path.getsize(perl_file)
        f = file(self.res_file, "wb")
        try:
            f.write(self.format.get_xf())
        finally:
            f.close()
        rsize = os.path.getsize(self.res_file)
        self.assertEqual(size, rsize, "File sizes not the same")
        
        contents = read_file(perl_file)
        rcontents = read_file(self.res_file)
        self.assertEqual(contents, rcontents, "Contents not the same")

    def test_font_biff_size(self):
        perl_file = "perl_output/f_font_biff"
        f = file(self.res_file, "wb")
        try:
            f.write(self.format.get_font())
        finally:
            f.close()
        contents = read_file(perl_file)
        rcontents = read_file(self.res_file)
        self.assertEqual(contents, rcontents, "Contents not the same")

        contents = read_file(perl_file)
        rcontents = read_file(self.res_file)
        self.assertEqual(contents, rcontents, "Contents not the same")

    def test_get_font_key_size(self):
        perl_file = "perl_output/f_font_key"
        f = file(self.res_file, "wb")
        try:
            f.write(self.format.get_font_key())
        finally:
            f.close()
        self.assertEqual(os.path.getsize(perl_file), os.path.getsize(self.res_file), \
                            "Bad file size")

        contents = read_file(perl_file)
        rcontents = read_file(self.res_file)
        self.assertEqual(contents, rcontents, "Contents not the same")

    def test_color(self):
        self.format.set_color("blue")
        self.assertEqual(0x0C, self.format._color)

    def test_color_bogus(self):
        self.format.set_color("blah")
        self.assertEqual(0x7FFF, self.format._color)

    def test_align(self):
        self.format.set_align()
        self.assertEqual(0, self.format._text_h_align)
        self.assertEqual(2, self.format._text_v_align)

    def test_align_center(self):
        self.format.set_align("center")
        self.assertEqual(2, self.format._text_h_align)
        self.assertEqual(2, self.format._text_v_align)

    def test_bold(self):
        self.format.set_bold(False)
        self.assertEqual(400, self.format._bold)
        self.format.set_bold()
        self.assertEqual(700, self.format._bold)
        self.format.set_bold(101)
        self.assertEqual(101, self.format._bold)
        self.format.set_bold(1001)
        self.assertEqual(1000, self.format._bold)
        self.format.set_bold(550.1)
        self.assertEqual(550, self.format._bold)


if __name__ == "__main__":
    unittest.main()