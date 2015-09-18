#!/bin/perl -w
# $Id: gen_f_xf_biff.pl,v 1.7 2004/01/31 12:31:37 fufff Exp $

use strict;
use Spreadsheet::WriteExcel::Format;

my $format = Spreadsheet::WriteExcel::Format->new();

open FO, "> f_xf_biff";
binmode FO;

print FO $format->get_xf();