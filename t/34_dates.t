#!/usr/bin/perl

use strict;
use warnings;

use Test::More;

BEGIN { delete @ENV{qw( LANG LC_ALL LC_DATE )}; }

use Spreadsheet::Read;
if (Spreadsheet::Read::parses ("xls")) {
    plan tests => 69;
    }
else {
    plan skip_all => "No M\$-Excel parser found";
    }

my $xls;
ok ($xls = ReadData ("files/Dates.xls", attr => 1, dtfmt => "yyyy-mm-dd"), "Excel Date testcase");

my $ss   = $xls->[1];
my $attr = $ss->{attr};

my @date = (undef, 39668, 39672,      39790,        39673);
my @fmt  = (undef, undef, "yyyymmdd", "yyyy-mm-dd", "mm/dd/yyyy");
foreach my $r (1 .. 4) {
    is ($ss->{cell}[$_][$r], $date[$r],	"Date value  row $r col $_") for 1 .. 4;

    is ($attr->[$_][$r]{type},   "date",   "Date type   row $r col $_")  for 1 .. 4;
    is ($attr->[$_][$r]{format}, $fmt[$_], "Date format row $r col $_")  for 1 .. 4;
    }

is ($ss->{A1},	 "8-Aug",	"Cell content A1");
is ($ss->{A2},	"12-Aug",	"Cell content A2");
is ($ss->{A3},	 "8-Dec",	"Cell content A3");
is ($ss->{A4},	"13-Aug",	"Cell content A4");

is ($ss->{B1},	20080808,	"Cell content B1");
is ($ss->{B2},	20080812,	"Cell content B2");
is ($ss->{B3},	20081208,	"Cell content B3");
is ($ss->{B4},	20080813,	"Cell content B4");

is ($ss->{C1},	"2008-08-08",	"Cell content C1");
is ($ss->{C2},	"2008-08-12",	"Cell content C2");
is ($ss->{C3},	"2008-12-08",	"Cell content C3");
is ($ss->{C4},	"2008-08-13",	"Cell content C4");

is ($ss->{D1},	"08/08/2008",	"Cell content D1");
is ($ss->{D2},	"08/12/2008",	"Cell content D2");
is ($ss->{D3},	"12/08/2008",	"Cell content D3");
is ($ss->{D4},	"08/13/2008",	"Cell content D4");

is ($ss->{E1},	"08 Aug 2008",	"Cell content E1");
is ($ss->{E2},	"12 Aug 2008",	"Cell content E2");
is ($ss->{E3},	"08 Dec 2008",	"Cell content E3");
is ($ss->{E4},	"13 Aug 2008",	"Cell content E4");

# Below can only be checked when SS::PE 0.34 is out
#use DDumper;
#foreach my $r (1..4,6..7) {
#    foreach my $c (1..5) {
#	my $cell = cr2cell ($c, $r);
#	my $fmt  = $ss->{attr}[$c][$r]{format};
#	defined $ss->{$cell} or next;
#	printf STDERR "# attr %s: %-22s %s\n",
#	    $cell, $ss->{$cell}, defined $fmt ? "'$fmt'" : "<undef>";
#	}
#    }
