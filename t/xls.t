#!/usr/bin/perl

use strict;
use warnings;

use Test::More tests => 195;

use Spreadsheet::Read;

{   my $ref;
    $ref = ReadData ("no_such_file.xls");
    ok (!defined $ref, "Nonexistent file");
    $ref = ReadData ("empty.xls");
    ok (!defined $ref, "Empty file");
    }

my $content;
{   local $/;
    open my $xls, "< files/test.xls";
    $content = <$xls>;
    }

my $xls;
foreach my $base ( [ "files/test.xls",	"Read/Parse xls file"	],
		   [ $content,		"Parse xls data"	],
		   ) {
    my ($txt, $msg) = @$base;
    ok ($xls = ReadData ($txt),	$msg);

    ok (1, "Base values");
    is (ref $xls,		"ARRAY",	"Return type");
    is ($xls->[0]{type},	"xls",		"Spreadsheet type");
    is ($xls->[0]{sheets},	2,		"Sheet count");
    is (ref $xls->[0]{sheet},	"HASH",		"Sheet list");
    is (scalar keys %{$xls->[0]{sheet}},
				2,		"Sheet list count");
    cmp_ok ($xls->[0]{version}, ">=",	0.26,	"Parser version");

    ok (1, "Defined fields");
    foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($xls->[1]{cell}[$c][$r],	$cell,	"Unformatted cell $cell");
	is ($xls->[1]{$cell},		$cell,	"Formatted   cell $cell");
	}

    ok (1, "Undefined fields");
    foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($xls->[1]{cell}[$c][$r],	undef,	"Unformatted cell $cell");
	is ($xls->[1]{$cell},		undef,	"Formatted   cell $cell");
	}
    }

# This files is generated under Mac OS/X Tiger
ok (1, "XLS File fom Mac OS X");
ok ($xls = ReadData ("files/macosx.xls"),	"Read/Parse Mac OS X xls file");

ok (1, "Base values");
is ($xls->[0]{sheets},		3,		"Sheet count");
is ($xls->[0]{sheet}{Sheet3},	3,		"Sheet labels");
is ($xls->[1]{maxrow},		25,		"MaxRow");
is ($xls->[1]{maxcol},		3,		"MaxCol");
is ($xls->[2]{label},		"Sheet2",	"Sheet label");
is ($xls->[2]{maxrow},		0,		"Empty sheet maxrow");
is ($xls->[2]{maxcol},		0,		"Empty sheet maxcol");

ok (1, "Content");
is ($#{$xls->[1]{cell}[3]}, $xls->[1]{maxrow}, "cell structure");
ok (defined $xls->[1]{cell}[$xls->[1]{maxcol}][$xls->[1]{maxrow}], "last cell");

foreach my $x (1 .. 17) {
    my $cell = cr2cell (1, $x);
    is ($xls->[1]{$cell},		$x,	"Cell $cell");
    is ($xls->[1]{cell}[1][$x],		$x,	"Cell 1, $x");
    }
foreach my $x (1 .. 25) {
    my $cell = cr2cell (3, $x);
    is ($xls->[1]{$cell},		$x,	"Cell $cell");
    is ($xls->[1]{cell}[3][$x],		$x,	"Cell 3, $x");
    }
foreach my $cell (qw( A18 B1 B6 B20 C26 D14 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($xls->[1]{cell}[$c][$r],	undef,	"Cell $cell");
    is ($xls->[1]{$cell},		undef,	"Cell $c, $r");
    }
