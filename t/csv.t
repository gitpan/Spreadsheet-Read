#!/usr/bin/perl

use strict;
use warnings;

use Test::More tests => 107;

use Spreadsheet::Read;

{   my $ref;
    $ref = ReadData ("no_such_file.csv");
    ok (!defined $ref, "Nonexistent file");
    $ref = ReadData ("empty.csv");
    ok (!defined $ref, "Empty file");
    }

my $csv;
ok ($csv = ReadData ("files/test.csv"),	"Read/Parse csv file");

ok (1, "Defined fields");
foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	$cell,	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		$cell,	"Formatted   cell $cell");
    }

ok (1, "Undefined fields");
foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	"",   	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		"",   	"Formatted   cell $cell");
    }

ok ($csv = ReadData ("files/test_m.csv"),	"Read/Parse M\$ csv file");

ok (1, "Defined fields");
foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	$cell,	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		$cell,	"Formatted   cell $cell");
    }

ok (1, "Undefined fields");
foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	"",   	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		"",   	"Formatted   cell $cell");
    }

ok ($csv = ReadData ("files/test_x.csv", sep => "=", quote => "_"),
					    "Read/Parse strange csv file");

ok (1, "Defined fields");
foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	$cell,	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		$cell,	"Formatted   cell $cell");
    }

ok (1, "Undefined fields");
foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
    my ($c, $r) = cell2cr ($cell);
    is ($csv->[1]{cell}[$c][$r],	"",   	"Unformatted cell $cell");
    is ($csv->[1]{$cell},		"",   	"Formatted   cell $cell");
    }
