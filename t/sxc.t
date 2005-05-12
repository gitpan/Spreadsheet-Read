#!/usr/bin/perl

use strict;
use warnings;

use Test::More tests => 273;

use Spreadsheet::Read;

my $content;
{   local $/;
    open my $xml, "< files/content.xml";
    $content = <$xml>;
    }

foreach my $base ( [ "files/test.sxc",		"Read/Parse sxc file" ],
		   [ "files/content.xml",	"Read/Parse xml file" ],
		   [ $content,			"Parse xml data" ],
		   ) {
    my ($txt, $msg) = @$base;
    my $sxc;
    ok ($sxc = ReadData ($content), $msg);

    ok (1, "Sheet 1");
    # Simple sheet with cells filled with the cell label:
    # -- -- -- --
    # A1 B1    D1
    # A2 B2
    # A3    C3 D3
    # A4 B4 C4

    ok (1, "Defined fields");
    foreach my $cell (qw( A1 A2 A3 A4 B1 B2 B4 C3 C4 D1 D3 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($sxc->[1]{cell}[$c][$r],	$cell,	"Unformatted cell $cell");
	is ($sxc->[1]{$cell},		$cell,	"Formatted   cell $cell");
	}

    ok (1, "Undefined fields");
    foreach my $cell (qw( B3 C1 C2 D2 D4 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($sxc->[1]{cell}[$c][$r],	undef,	"Unformatted cell $cell");
	is ($sxc->[1]{$cell},		undef,	"Formatted   cell $cell");
	}

    ok (1, "Nonexistent fields");
    foreach my $cell (qw( A9 X6 B17 AB4 BE33 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($sxc->[1]{cell}[$c][$r],	undef,	"Unformatted cell $cell");
	is ($sxc->[1]{$cell},		undef,	"Formatted   cell $cell");
	}

    ok (1, "Sheet 2");
    # Sheet with merged cells and notes/annotations
    # x   x   x
    #   x   x 
    # x   x   x

    ok (1, "Defined fields");
    foreach my $cell (qw( A1 C1 E1 B2 D2 A3 C3 E3 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($sxc->[2]{cell}[$c][$r],	"x",	"Unformatted cell $cell");
	is ($sxc->[2]{$cell},		"x",	"Formatted   cell $cell");
	}

    ok (1, "Undefined fields");
    foreach my $cell (qw( B1 D1 A2 C2 E2 B3 D3 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($sxc->[2]{cell}[$c][$r],	undef,	"Unformatted cell $cell");
	is ($sxc->[2]{$cell},		undef,	"Formatted   cell $cell");
	}

    ok (1, "Nonexistent fields");
    foreach my $cell (qw( A9 X6 B17 AB4 BE33 )) {
	my ($c, $r) = cell2cr ($cell);
	is ($sxc->[2]{cell}[$c][$r],	undef,	"Unformatted cell $cell");
	is ($sxc->[2]{$cell},		undef,	"Formatted   cell $cell");
	}
    }
