#!/usr/bin/perl

use strict;
use warnings;

use Test::More tests => 6;

use Spreadsheet::Read;

{   my $ref;
    $ref = ReadData ("no_such_file.sc");
    ok (!defined $ref, "Nonexistent file");
    $ref = ReadData ("empty.sc");
    ok (!defined $ref, "Empty file");
    }

my $sc;
ok ($sc = ReadData ("files/test.sc"), "Read/Parse sc file");

is ($sc->[1]{maxcol},		10,		"Columns");
is ($sc->[1]{maxrow},		26,		"Rows");
is ($sc->[1]{cell}[1][22],	"  Workspace",	"Just checking one cell");
