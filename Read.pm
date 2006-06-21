#!/pro/bin/perl

package Spreadsheet::Read;

=head1 NAME

Spreadsheet::Read - Read the data from a spreadsheet

=head1 SYNOPSYS

 use Spreadsheet::Read;
 my $ref = ReadData ("test.csv", sep => ";");
 my $ref = ReadData ("test.sxc");
 my $ref = ReadData ("test.ods");
 my $ref = ReadData ("test.xls");

 my $a3 = $ref->[1]{A3}, "\n"; # content of field A3 of sheet 1

=cut

use strict;
use warnings;

our $VERSION = "0.15";
sub  Version { $VERSION }

use Exporter;
our @ISA       = qw( Exporter );
our @EXPORT    = qw( ReadData cell2cr cr2cell );
our @EXPORT_OK = qw( parses rows );

use File::Temp           qw( );
use Data::Dumper;

my %can = map { $_ => 0 } qw( csv sxc xls prl );
for (	[ csv	=> "Text::CSV_XS"		],
#	[ csv	=> "Text::CSV"			],	# NYI
	[ sxc	=> "Spreadsheet::ReadSXC"	],
	[ xls	=> "Spreadsheet::ParseExcel"	],
	[ prl	=> "Spreadsheet::Perl"		],
	) {
    my ($flag, $mod) = @$_;
    $can{$flag} and next;
    eval "require $mod; \$can{$flag} = '$mod'";
    }
$can{sc} = 1;	# SquirelCalc is built-in

my $debug = 0;
my @def_attr = (
    type    => "text",
    fgcolor => undef,
    bgcolor => undef,
    font    => undef,
    size    => undef,
    format  => undef,
    halign  => "left",
    valign  => "top",
    uline   => 0,
    bold    => 0,
    italic  => 0,
    wrap    => 0,
    enc     => "utf-8", # $ENV{LC_ALL} // $ENV{LANG} // ...
    );

# Helper functions

# Spreadsheet::Read::parses ("csv") or die "Cannot parse CSV"
sub parses ($)
{
    my $type = shift		or  return 0;
    $type = lc $type;
    # Aliases and fullnames
    $type eq "excel"		and return $can{xls};
    $type eq "oo"		and return $can{sxc};
    $type eq "ods"		and return $can{sxc};
    $type eq "openoffice"	and return $can{sxc};
    $type eq "perl"		and return $can{prl};

    # $can{$type} // 0;
    exists $can{$type} ? $can{$type} : 0;
    } # parses

# cr2cell (4, 18) => "D18"
# No prototype to allow 'cr2cell (@rowcol)'
sub cr2cell
{
    my ($c, $r) = @_;
    defined $c && defined $r && $c > 0 && $r > 0 or return "";
    my $cell = "";
    while ($c) {
	use integer;

	substr ($cell, 0, 0) = chr (--$c % 26 + ord "A");
	$c /= 26;
	}
    "$cell$r";
    } # cr2cell

# cell2cr ("D18") => (4, 18)
sub cell2cr ($)
{
    my ($cc, $r) = ((uc $_[0]) =~ m/^([A-Z]+)(\d+)$/) or return (0, 0);
    my $c = 0;
    while ($cc =~ s/^([A-Z])//) {
	$c = 26 * $c + 1 + ord ($1) - ord ("A");
	}
    ($c, $r);
    } # cell2cr

# Convert {cell}'s [column][row] to a [row][column] list
# my @rows = rows ($ss->[1]);
sub rows ($)
{
    my $sheet = shift or return;
    ref    $sheet eq "HASH" && exists $sheet->{cell}   or return;
    exists $sheet->{maxcol} && exists $sheet->{maxrow} or return;
    my $s = $sheet->{cell};

    map {
	my $r = $_;
	[ map { $s->[$_][$r] } 1..$sheet->{maxcol} ];
	} 1..$sheet->{maxrow};
    } # rows

# If option "clip" is set, remove the trailing lines and
# columns in each sheet that contain no visible data
sub _clipsheets
{
    my ($clip, $ref) = @_;
    $clip or return $ref;

    foreach my $sheet (1 .. $ref->[0]{sheets}) {
	my $ss = $ref->[$sheet];

	# Remove trailing empty columns
	while ($ss->{maxcol} and not (
		grep { defined && m/\S/ } @{$ss->{cell}[$ss->{maxcol}]})
		) {
	    (my $col = cr2cell ($ss->{maxcol}, 1)) =~ s/1$//; 
	    my $recol = qr{^$col(?=\d+)$};
	    delete $ss->{$_} for grep m/$recol/, keys %{$ss};
	    $ss->{maxcol}--;
	    }
	$ss->{maxcol} or $ss->{maxrow} = 0;

	# Remove trailing empty lines
	while ($ss->{maxrow} and not (
		grep { defined && m/\S/ }
		map  { $ss->{cell}[$_][$ss->{maxrow}] }
		1 .. $ss->{maxcol}
		)) {
	    my $rerow = qr{^[A-Z]+$ss->{maxrow}$};
	    delete $ss->{$_} for grep m/$rerow/, keys %{$ss};
	    $ss->{maxrow}--;
	    }
	$ss->{maxrow} or $ss->{maxcol} = 0;
	}
    $ref;
    } # _clipsheets

sub ReadData ($;@)
{
    my $txt = shift	or  return;
    ref $txt		and return; # TODO: support IO stream (ref $txt eq "IO")

    my $tmpfile;

    my %opt;
    if (@_) {
	   if (ref $_[0] eq "HASH")  { %opt = %{shift @_} }
	elsif (@_ % 2 == 0)          { %opt = @_          }
	}
    defined $opt{rc}	or $opt{rc}	= 1;
    defined $opt{cells}	or $opt{cells}	= 1;
    defined $opt{attr}	or $opt{attr}	= 0;
    defined $opt{clip}	or $opt{clip}	= $opt{cell};

    # $debug = $opt{debug} // 0;
    $debug = defined $opt{debug} ? $opt{debug} : 0;
    $debug > 4 and print STDERR Data::Dumper->Dump ([\%opt],["Options"]);

    # CSV not supported from streams
    if ($txt =~ m/\.(csv)$/i and -f $txt) {
	$can{csv} or die "CSV parser not installed";

	open my $in, "< $txt" or return;
	my $csv;
	my @data = (
	    {	type	=> "csv",
		version	=> $Text::CSV_XS::VERSION,
		sheets	=> 1,
		sheet	=> { $txt => 1 },
		},
	    {	label	=> $txt,
		maxrow	=> 0,
		maxcol	=> 0,
		cell	=> [],
		attr	=> [],
		},
	    );
	while (<$in>) {
	    unless ($csv) {
		my $quo = defined $opt{quote} ? $opt{quote} : '"';
		my $sep = # If explicitely set, use it
		   defined $opt{sep} ? $opt{sep} :
		       # otherwise start auto-detect with quoted strings
		       m/["\d];["\d;]/  ? ";"  :
		       m/["\d],["\d,]/  ? ","  :
		       m/["\d]\t["\d,]/ ? "\t" :
		       # If neither, then for unquoted strings
		       m/\w;[\w;]/      ? ";"  :
		       m/\w,[\w,]/      ? ","  :
		       m/\w\t[\w,]/     ? "\t" :
					  ","  ;
		$csv = Text::CSV_XS->new ({
		    sep_char   => $sep,
		    quote_char => $quo,
		    binary     => 1,
		    });
		}
	    $csv->parse ($_);
	    my @row = $csv->fields () or next;
	    my $r = ++$data[1]{maxrow};
	    @row > $data[1]{maxcol} and $data[1]{maxcol} = @row;
	    foreach my $c (0 .. $#row) {
		my $val = $row[$c];
		my $cell = cr2cell ($c + 1, $r);
		$opt{rc}    and $data[1]{cell}[$c + 1][$r] = $val;
		$opt{cells} and $data[1]{$cell} = $val;
		$opt{attr}  and $data[1]{attr}[$c + 1][$r] = { @def_attr };
		}
	    }
	for (@{$data[1]{cell}}) {
	    defined $_ or $_ = [];
	    }
	close $in;
	return _clipsheets $opt{clip}, [ @data ];
	}

    # From /etc/magic: Microsoft Office Document
    if ($txt =~ m/^(\376\067\0\043
		   |\320\317\021\340\241\261\032\341
		   |\333\245-\0\0\0)/x) {
	$can{xls} or die "Spreadsheet::ParseExcel not installed";
	$tmpfile = File::Temp->new (SUFFIX => ".xls", UNLINK => 1);
	print $tmpfile $txt;
	$txt = "$tmpfile";
	}
    if ($txt =~ m/\.xls$/i and -f $txt) {
	$can{xls} or die "Spreadsheet::ParseExcel not installed";
	$debug and print STDERR "Opening XLS $txt\n";
	my $oBook = Spreadsheet::ParseExcel::Workbook->Parse ($txt);
	$debug > 8 and print STDERR Data::Dumper->Dump ([$oBook],["oBook"]);
	my @data = ( {
	    type	=> "xls",
	    version	=> $Spreadsheet::ParseExcel::VERSION,
	    sheets	=> $oBook->{SheetCount},
	    sheet	=> {},
	    } );
	$debug and print STDERR "\t$data[0]{sheets} sheets\n";
	foreach my $oWkS (@{$oBook->{Worksheet}}) {
	    $opt{clip} and !defined $oWkS->{Cells} and next; # Skip empty sheets
	    my %sheet = (
		label	=> $oWkS->{Name},
		maxrow	=> 0,
		maxcol	=> 0,
		cell	=> [],
		attr	=> [],
		);
	    defined $sheet{label}  or  $sheet{label}  = "-- unlabeled --";
	    exists $oWkS->{MaxRow} and $sheet{maxrow} = $oWkS->{MaxRow} + 1;
	    exists $oWkS->{MaxCol} and $sheet{maxcol} = $oWkS->{MaxCol} + 1;
	    my $sheet_idx = 1 + @data;
	    $debug and print STDERR "\tSheet $sheet_idx '$sheet{label}' $sheet{maxrow} x $sheet{maxcol}\n";
	    if (exists $oWkS->{MinRow}) {
		foreach my $r ($oWkS->{MinRow} .. $sheet{maxrow}) { 
		    foreach my $c ($oWkS->{MinCol} .. $sheet{maxcol}) { 
			my $oWkC = $oWkS->{Cells}[$r][$c] or next;
			defined (my $val = $oWkC->{Val})  or next;
			my $cell = cr2cell ($c + 1, $r + 1);
			$opt{rc}    and $sheet{cell}[$c + 1][$r + 1] = $val;	# Original
			$opt{cells} and $sheet{$cell} = $oWkC->Value;	# Formatted
			if ($opt{attr}) {
			    my $FmT = $oWkC->Format;
			    my $FnT = $FmT->Font;
			    $sheet{attr}[$c + 1][$r] = {
				@def_attr,

				type    => lc $oWkC->Type,
				enc     => $oWkC->Code,
				format  => "...",
				halign  => [ undef, qw( left center right
					    fill justify ), undef,
					    "equal_space" ]->[$FmT->AlignH],
				valign  => [ qw( top center bottom justify
					    equal_space )]->[$FmT->AlignV],
				wrap    => $FmT->Wrap,
				font    => $FnT->Name,
				fgcolor => $FnT->Color, # Color INDEX :(
				bgcolor => $FmT->Fill->[2], # Color INDEX
				};
			    }
			}
		    }
		}
	    for (@{$sheet{cell}}) {
		defined $_ or $_ = [];
		}
	    push @data, { %sheet };
#	    $data[0]{sheets}++;
	    if ($sheet{label} eq "-- unlabeled --") {
		$sheet{label} = "";
		}
	    else {
		$data[0]{sheet}{$sheet{label}} = $#data;
		}
	    }
	return _clipsheets $opt{clip}, [ @data ];
	}

    if ($txt =~ m/^# .*SquirrelCalc/ or $txt =~ m/\.sc$/ && -f $txt) {
	if ($txt !~ m/\n/ && -f $txt) {
	    local $/;
	    open my $sc, "< $txt" or return;
	    $txt = <$sc>;
	    $txt =~ m/\S/ or return;
	    }
	my @data = (
	    {	type	=> "sc",
		version	=> undef,
		sheets	=> 1,
		sheet	=> { sheet => 1 },
		},
	    {	label	=> "sheet",
		maxrow	=> 0,
		maxcol	=> 0,
		cell	=> [],
		attr	=> [],
		},
	    );

	for (split m/\s*[\r\n]\s*/, $txt) {
	    if (m/^dimension.*of (\d+) rows.*of (\d+) columns/i) {
		@{$data[1]}{qw(maxrow maxcol)} = ($1, $2);
		next;
		}
	    s/^r(\d+)c(\d+)\s*=\s*// or next;
	    my ($c, $r) = map { $_ + 1 } $2, $1;
	    if (m/.* {(.*)}$/ or m/"(.*)"/) {
		my $cell = cr2cell ($c, $r);
		$opt{rc}    and $data[1]{cell}[$c][$r] = $1;
		$opt{cells} and $data[1]{$cell} = $1;
		$opt{attr}  and $data[1]{attr}[$c + 1][$r] = { @def_attr };
		next;
		}
	    # Now only formula's remain. Ignore for now
	    # r67c7 = [P2L] 2*(1000*r67c5-60)
	    }
	for (@{$data[1]{cell}}) {
	    defined $_ or $_ = [];
	    }
	return _clipsheets $opt{clip}, [ @data ];
	}

    if ($txt =~ m/^<\?xml/ or -f $txt) {
	$can{sxc} or die "Spreadsheet::ReadSXC not installed";
	my $sxc_options = { OrderBySheet => 1 }; # New interface 0.20 and up
	my $sxc;
	   if ($txt =~ m/\.(sxc|ods)$/i) {
	    $debug and print STDERR "Opening \U$1\E $txt\n";
	    $sxc = Spreadsheet::ReadSXC::read_sxc      ($txt, $sxc_options)	or  return;
	    }
	elsif ($txt =~ m/\.xml$/i) {
	    $debug and print STDERR "Opening XML $txt\n";
	    $sxc = Spreadsheet::ReadSXC::read_xml_file ($txt, $sxc_options)	or  return;
	    }
	# need to test on pattern to prevent stat warning
	# on filename with newline
	elsif ($txt !~ m/^<\?xml/i and -f $txt) {
	    $debug and print STDERR "Opening XML $txt\n";
	    open my $f, "< $txt"	or  return;
	    local $/;
	    $txt = <$f>;
	    }
	!$sxc && $txt =~ m/^<\?xml/i and
	    $sxc = Spreadsheet::ReadSXC::read_xml_string ($txt, $sxc_options);
	$debug > 8 and print STDERR Data::Dumper->Dump ([$sxc],["sxc"]);
	if ($sxc) {
	    my @data = ( {
		type	=> "sxc",
		version	=> $Spreadsheet::ReadSXC::VERSION,
		sheets	=> 0,
		sheet	=> {},
		} );
	    my @sheets = ref $sxc eq "HASH"	# < 0.20
		? map {
		    {   label => $_,
			data  => $sxc->{$_},
			}
		    } keys %$sxc
		: @{$sxc};
	    foreach my $sheet (@sheets) {
		my @sheet = @{$sheet->{data}};
		my %sheet = (
		    label	=> $sheet->{label},
		    maxrow	=> scalar @sheet,
		    maxcol	=> 0,
		    cell	=> [],
		    attr	=> [],
		    );
		my $sheet_idx = 1 + @data;
		$debug and print STDERR "\tSheet $sheet_idx '$sheet{label}' $sheet{maxrow} rows\n";
		foreach my $r (0 .. $#sheet) {
		    my @row = @{$sheet[$r]} or next;
		    foreach my $c (0 .. $#row) {
			my $val = $row[$c] or next;
			my $C = $c + 1;
			$C > $sheet{maxcol} and $sheet{maxcol} = $C;
			my $cell = cr2cell ($C, $r + 1);
			$opt{rc}    and $sheet{cell}[$C][$r + 1] = $val;
			$opt{cells} and $sheet{$cell} = $val;
			$opt{attr}  and $sheet{attr}[$C][$r + 1] = { @def_attr };
			}
		    }
		for (@{$sheet{cell}}) {
		    defined $_ or $_ = [];
		    }
		$debug and print STDERR "\tSheet $sheet_idx '$sheet{label}' $sheet{maxrow} x $sheet{maxcol}\n";
		push @data, { %sheet };
		$data[0]{sheets}++;
		$data[0]{sheet}{$sheet->{label}} = $#data;
		}
	    return _clipsheets $opt{clip}, [ @data ];
	    }
	}

    return;
    } # ReadData

1;

=head1 DESCRIPTION

Spreadsheet::Read tries to transparantly read *any* spreadsheet and
return its content in a universal manner independent of the parsing
module that does the actual spreadsheet scanning.

For OpenOffice this module uses Spreadsheet::ReadSXC

For Excel this module uses Spreadsheet::ParseExcel

For CSV this module uses Text::CSV_XS

For SquirrelCalc there is a very simplistic built-in parser

=head2 Data structure

The data is returned as an array reference:

  $ref = [
 	# Entry 0 is the overall control hash
 	{ sheets  => 2,
	  sheet   => {
	    "Sheet 1"	=> 1,
	    "Sheet 2"	=> 2,
	    },
	  type    => "xls",
	  version => 0.26,
	  },
 	# Entry 1 is the first sheet
 	{ label  => "Sheet 1",
 	  maxrow => 2,
 	  maxcol => 4,
 	  cell   => [ undef,
	    [ undef, 1 ],
	    [ undef, undef, undef, undef, undef, "Nugget" ],
	    ],
 	  A1     => 1,
 	  B5     => "Nugget",
 	  },
 	# Entry 2 is the second sheet
 	{ label => "Sheet 2",
 	  :
 	:

To keep as close contact to spreadsheet users, row and column 1 have
index 1 too in the C<cell> element of the sheet hash, so cell "A1" is
the same as C<cell> [1, 1] (column first). To switch between the two,
there are two helper functions available: C<cell2cr ()> and C<cr2cell ()>.

The C<cell> hash entry contains unformatted data, while the hash entries
with the traditional labels contain the formatted values (if applicable).

The control hash (the first entry in the returned array ref), contains
some spreadsheet metadata. The entry C<sheet> is there to be able to find
the sheets when accessing them by name:

  my %sheet2 = %{$ref->[$ref->[0]{sheet}{"Sheet 2"}]};

=head2 Functions

=over 2

=item C<my $ref = ReadData ($source [, option => value [, ... ]]);>

=item C<my $ref = ReadData ("file.csv", sep =&gt; ',', quote => '"');>

=item C<my $ref = ReadData ("file.xls");>

=item C<my $ref = ReadData ("file.ods");>

=item C<my $ref = ReadData ("file.sxc");>

=item C<my $ref = ReadData ("content.xml");>

=item C<my $ref = ReadData ($content);>

Tries to convert the given file, string, or stream to the data
structure described above.

Precessing data from a stream or content is supported for Excel (through a
File::Temp temporary file), or for XML (OpenOffice), but not for CSV.

ReadSXC does preserve sheet order as of version 0.20.

Currently supported options are:

=over 2

=item cells

Control the generation of named cells ("A1" etc). Default is true.

=item rc

Control the generation of the {cell}[c][r] entries. Default is true.

=item attr

Control the generation of the {attr}[c][r] entries. Default is false.

=item clip

If set, C<ReadData ()> will remove all trailing lines and columns per
sheet that have no visual data.
This option is only valid if C<cells> is true. The default value is
true if C<cells> is true, and false otherwise.

=item sep

Set separator for CSV. Default is comma C<,>.

=item quote

Set quote character for CSV. Default is C<">.

=item debug

Enable some diagnostic messages to STDERR.

The value determines how much diagnostics are dumped (using Data::Dumper).
A value of 9 and higher will dump the entire structure from the backend
parser.

=back

=item C<my $cell = cr2cell (col, row)>

C<cr2cell ()> converts a C<(column, row)> pair (1 based) to the
traditional cell notation:

  my $cell = cr2cell ( 4, 14); # $cell now "D14"
  my $cell = cr2cell (28,  4); # $cell now "AB4"

=item C<my ($col, $row) = cell2cr ($cell)>

C<cell2cr ()> converts traditional cell notation to a C<(column, row)>
pair (1 based):

  my ($col, $row) = cell2cr ("D14"); # returns ( 4, 14)
  my ($col, $row) = cell2cr ("AB4"); # returns (28,  4)

=item my @rows = rows ($ref)

=item C<my @rows = Spreadsheet::Read::rows ($ss-&gt;[1])>

Convert C<{cell}>'s C<[column][row]> to a C<[row][column]> list.

Note that the indexes in the returned list are 0-based, where the
index in the C<{cell}> entry is 1-based.

C<rows ()> is not imported by default, so either specify it in the
use argument list, or call it fully qualified.

=item parses ($format)

=item C<Spreadsheet::Read::parses ("CSV")>

C<parses ()> returns Spreadsheet::Read's capability to parse the
required format.

C<parses ()> is not imported by default, so either specify it in the
use argument list, or call it fully qualified.

=item my $rs_version = Version ()

=item C<my $v = Spreadsheet::Read::Version ()>

Returns the current version of Spreadsheet::Read.

C<Version ()> is not imported by default, so either specify it in the
use argument list, or call it fully qualified.

=back

=head1 TODO

=over 4

=item Cell attributes

Future plans include cell attributes, available as for example:

 	{ label  => "Sheet 1",
 	  maxrow => 2,
 	  maxcol => 4,
 	  cell   => [ undef,
	    [ undef, 1 ],
	    [ undef, undef, undef, undef, undef, "Nugget" ],
	    ],
 	  attr   => [ undef,
 	    [ undef, {
 	      type   => "numeric",
 	      color  => "Red",
 	      font   => "Arial",
 	      enc    => "iso10646-1",
 	      size   => "12",
 	      format => "## ###.##",
 	      halign => "right",
 	      }, ]
	    [ undef, undef, undef, undef, undef, {
 	      type   => "text",
 	      color  => "#e2e2e2",
 	      font   => "LetterGothic",
 	      enc    => "iso8859-1",
 	      size   => "15",
 	      format => undef,
 	      halign => "left",
 	      }, ]
 	  A1     => 1,
 	  B5     => "Nugget",
 	  },

=item Options

=over 2

=item Module Options

New Spreadsheet::Read options are bound to happen. I'm thinking of an
option that disables the reading of the data entirely to speed up an
index request (how many sheets/fields/columns). See C<xlscat -i>.

=item Parser options

Try to transparently support as many options as the encapsulated modules
support regarding (un)formatted values, (date) formats, hidden columns
rows or fields etc. These could be implemented like C<attr> above but
names C<meta>, or just be new values in the C<attr> hashes.

=back

=item Other spreadsheet formats

I consider adding any spreadsheet interface that offers a usable API.

=item Safety / flexibility

Now that the different parsers are only activated if the module can be
loaded, we need more flexibility is switching from Text::CSV_XS to
Text::CSV in the parser part.

=item OO-ify

Consider making the ref an object, though I currently don't see the big
advantage (yet). Maybe I'll make it so that it is a hybrid functional /
OO interface.

=back

=head1 SEE ALSO

=over 2

=item Text::CSV_XS

http://search.cpan.org/~jwied/

A pure perl version is available on http://search.cpan.org/~makamaka/

=item Spreadsheet::ParseExcel

http://search.cpan.org/~kwitknr/

=item Spreadsheet::ReadSXC

http://search.cpan.org/~terhechte/

=item Text::CSV_XS, Text::CSV

http://search.cpan.org/~jwied/
http://search.cpan.org/~alancitt/

=item Spreadsheet::BasicRead

http://search.cpan.org/~gng/ for xlscat likewise functionality (Excel only)

=item Spreadsheet::ConvertAA

http://search.cpan.org/~nkh/ for an alternative set of cell2cr () /
cr2cell () pair

=item Spreadsheet::Perl

http://search.cpan.org/~nkh/ offers a Pure Perl implementation of a
spreadsheet engine. Users that want this format to be supported in
Spreadsheet::Read are hereby motivated to offer patches. It's not high
on my todo-list.

=item xls2csv

http://search.cpan.org/~ken/ offers an alternative for my C<xlscat -c>,
in the xls2csv tool, but this tool focusses on character encoding
transparency, and requires some other modules.

=back

=head1 AUTHOR

H.Merijn Brand, <h.m.brand@xs4all.nl>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2005-2006 H.Merijn Brand

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself. 

=cut