#!/usr/bin/perl

use strict;
use warnings;

use Test::More qw( no_plan );

use Spreadsheet::Read;

my $xls;
ok ($xls = ReadData ("files/test.xls"),	"Read/Parse xls file");

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
