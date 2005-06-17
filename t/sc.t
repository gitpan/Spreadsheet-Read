#!/usr/bin/perl

use strict;
use warnings;

use Test::More tests => 13;

use Spreadsheet::Read;

{   my $ref;
    $ref = ReadData ("no_such_file.sc");
    ok (!defined $ref, "Nonexistent file");
    $ref = ReadData ("empty.sc");
    ok (!defined $ref, "Empty file");
    }

my $sc;
ok ($sc = ReadData ("files/test.sc"), "Read/Parse sc file");

ok (1, "Base values");
is (ref $sc,			"ARRAY",	"Return type");
is ($sc->[0]{type},		"sc",		"Spreadsheet type");
is ($sc->[0]{sheets},		1,		"Sheet count");
is (ref $sc->[0]{sheet},	"HASH",		"Sheet list");
is (scalar keys %{$sc->[0]{sheet}},
				1,		"Sheet list count");
cmp_ok ($sc->[0]{version},	"==",	0,	"Parser version");

is ($sc->[1]{maxcol},		10,		"Columns");
is ($sc->[1]{maxrow},		26,		"Rows");
is ($sc->[1]{cell}[1][22],	"  Workspace",	"Just checking one cell");
