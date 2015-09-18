pyWriteExcel 0.4a2
==================

pyXLWriter is a Python library for generating Excel compatible spreadsheets.
It's a port of John McNamara's Perl Spreadsheet::WriteExcel module
version 1.01 to Python. It allows writing of Excel compatible spreadsheets
without the need for COM objects.

Now ported all features, but work in progress.

Requires Python 2.2+


Installation
============

For install run 'python setup.py install' in the root of the distribution.


Documentation
=============

Sorry, there is no documentation yet.

You'll have to use the documentation for Spreadsheet::WriteExcel available
at
http://search.cpan.org/doc/JMCNAMARA/Spreadsheet-WriteExcel-0.42/WriteExcel/doc/WriteExcel.html

I've included several example python scripts (see "examples/" directory).
Try to translate the Perl code into Python with the help of the examples.


TODOs and changes
=================

See TODO.txt and CHANGES.txt :-]


Authors, copyright, and license
===============================

pyXLWriter was ported by Evgeny Filatov from Perl Spreadsheet::WriteExcel
module. The author of the Spreadsheet::WriteExcel module is John McNamara
<jmcnamara@cpan.org>. You can find this module on http://www.cpan.org

pyXLWriter use PLY package (by David M. Beazley <beazley@cs.uchicago.edu>)
for parsing formulas (see http://systems.cs.uchicago.edu/ply/)

Copyright (c) 2004 Evgeny Filatov <fufff@users.sourceforge.net>
Copyright (c) 2002-2004 John McNamara (PERL Spreadsheet::WriteExcel)

Distribution's license is GNU LGPL. The included LICENSE.txt file describes
this in detail.


Useful links
============

Perl Spreadsheet WriteExcel
http://search.cpan.org/search?dist=Spreadsheet-WriteExcel

Perl Spreadsheet::ParseExcel
http://search.cpan.org/search?dist=Spreadsheet-ParseExcel

OpenOffice's Spreadsheet Project
http://sc.openoffice.org/
http://sc.openoffice.org/excelfileformat.pdf

POIFS (OLE2) format documentation
http://jakarta.apache.org/poi/poifs/fileformat.html
http://jakarta.apache.org/poi/poifs/fileformat.pdf

PHP Pear Spreadsheet-Excel-Writer
http://pear.php.net/package/Spreadsheet_Excel_Writer/

php_writeexcel
http://www.bettina-attack.de/jonny/projects/php_writeexcel/

PHP SimpleXlsGen
http://psxlsgen.sourceforge.net

Ruby Spreadsheet::Excel
http://www.freshports.org/textproc/ruby-spreadsheet-excel/

Java API for reading, writing and modifying the contents of Excel spreadsheets
http://www.andykhan.com/jexcelapi/index.html

PHP Spreadsheet Excel Reader
http://sourceforge.net/projects/phpexcelreader/
