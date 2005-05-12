#!/pro/bin/perl

package Spreadsheet::Read;

=head1 NAME

Spreadsheet::Read - Read the data from a spreadsheet

=head1 SYNOPSYS

use Spreadsheet::Read;
my $xls = ReadData ("test.xls");
my $sxc = ReadData ("test.sxc");

=cut

use strict;
use warnings;

our $VERSION = "0.01";
sub  Version { $VERSION }

use Exporter;
our @ISA     = qw( Exporter );
our @EXPORT  = qw( ReadData cell2cr cr2cell );

use Spreadsheet::ReadSXC qw( read_sxc read_xml_file read_xml_string );
use Spreadsheet::ParseExcel;

my  $debug = 0;

# Helper functions

# cr2cell (4, 18) => "D18"
sub cr2cell ($$)
{
    my ($c, $r) = @_;
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

sub ReadData ($)
{
    my $txt = shift				or  return;
    ref $txt					and return;

    if ($txt =~ m/\.(xls)$/i and -f $txt) {
	$debug and print STDERR "Opening XLS $txt\n";
	my $oBook = Spreadsheet::ParseExcel::Workbook->Parse ($txt);
	my @data = ( {
	    type   => "xls",
	    sheets => $oBook->{SheetCount},
	    sheet  => {},
	    } );
	$debug and print STDERR "\t$data[0]{sheets} sheets\n";
	foreach my $oWkS (@{$oBook->{Worksheet}}) {
	    my %sheet = (
		label	=> $oWkS->{Name},
		maxrow	=> $oWkS->{MaxRow} + 1,
		maxcol	=> $oWkS->{MaxCol} + 1,
		cell	=> [],
		);
	    my $sheet_idx = 1 + @data;
	    $debug and print STDERR "\tSheet $sheet_idx '$sheet{label}' $sheet{maxrow} x $sheet{maxcol}\n";
	    foreach my $r ($oWkS->{MinRow} .. $sheet{maxrow}) { 
		foreach my $c ($oWkS->{MinCol} .. $sheet{maxcol}) { 
		    my $oWkC = $oWkS->{Cells}[$r][$c] or next;
		    my $val = $oWkC->{Val} or next;
		    my $cell = cr2cell ($c + 1, $r + 1);
		    $sheet{cell}[$c + 1][$r + 1] = $val;	# Original
		    $sheet{$cell} = $oWkC->Value;		# Formatted
		    }
		}
	    push @data, { %sheet };
	    $data[0]{sheets}++;
	    $data[0]{sheet}{$sheet{label}} = $#data;
	    }
	return [ @data ];
	}

    if ($txt =~ m/^<\?xml/ or -f $txt) {
	my $sxc;
	   if ($txt =~ m/\.sxc$/i) {
	    $debug and print STDERR "Opening SXC $txt\n";
	    $sxc = read_sxc ($txt)		or  return;
	    }
	elsif ($txt =~ m/\.xml$/i) {
	    $debug and print STDERR "Opening XML $txt\n";
	    $sxc = read_xml_file ($txt)	or  return;
	    }
	# need to test on pattern to prevent stat warning
	# on filename with newline
	elsif ($txt !~ m/^<\?xml/i and -f $txt) {
	    $debug and print STDERR "Opening XML $txt\n";
	    open my $f, "<$txt"		or  return;
	    local $/;
	    $txt = <$f>;
	    }
	!$sxc && $txt =~ m/^<\?xml/i and $sxc = read_xml_string ($txt);
	if ($sxc) {
	    my @data = ( {
		type   => "sxc",
		sheets => 0,
		sheet  => {},
		} );
	    foreach my $sheet (keys %$sxc) {
		my @sheet = @{$sxc->{$sheet}};
		my %sheet = (
		    label	=> $sheet,
		    maxrow	=> scalar @sheet,
		    maxcol	=> 0,
		    cell	=> [],
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
			$sheet{cell}[$C][$r + 1] = $sheet{$cell} = $val;
			}
		    }
		$debug and print STDERR "\tSheet $sheet_idx '$sheet{label}' $sheet{maxrow} x $sheet{maxcol}\n";
		push @data, { %sheet };
		$data[0]{sheets}++;
		$data[0]{sheet}{$sheet} = $#data;
		}
	    return [ @data ];
	    }
	}

    return;
    } # ReadData

1;

=head1 DESCRIPTION

Spreadsheet::Read tries to transparantly read *any* spreadsheet and
return it's content in a universal manner independant of the parsing
module that does the actual spreadsheet scanning.

For OpenOffice this module uses Spreadsheet::ReadSXC

For Excel this module uses Spreadsheet::ParseExcel

=head2 Data structure

The data is returned as an array reference:

  $ref = [
 	# Entry 0 is the overall control hash
 	{ sheets => 2,
	  sheet  => {
	    "Sheet 1"	=> 1,
	    "Sheet 2"	=> 2,
	    },
	  type   => "xls",
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
 	  B4     => "Nugget",
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

=item C<my $ref = ReadData ("file.xls");

=item C<my $ref = ReadData ("file.sxc");

=item C<my $ref = ReadData ("content.xml");

=item C<my $ref = ReadData ($content);

Tries to convert the given file or string to the data structure
described above.

Currently ReadSXC does not preserve sheet order.

=item C<my $cell = cr2cell (col, row)>

C<cr2cell ()> converts a C<(column, row)> pair (1 based) to the
traditional cell notation:

  my $cell = cr2cell ( 4, 14); # $cell now "D14"
  my $cell = cr2cell (28,  4); # $cell now "AB4"

=iten C<my ($col, $row) = cell2cr ($cell)>

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
 	      color  => "Red",
 	      font   => "Arial",
 	      size   => "12",
 	      format => "## ###.##",
 	      align  => "right",
 	      }, ]
	    [ undef, undef, undef, undef, undef, {
 	      color  => "#e2e2e2",
 	      font   => "LetterGothic",
 	      size   => "15",
 	      format => undef,
 	      align  => "left",
 	      }, ]
 	  A1     => 1,
 	  B4     => "Nugget",
 	  },

=item Options

Try to transparently support as many options as the encapsulated modules
support regarding (un)formatted values, (date) formats, hidden columns
rows or fields etc. These could be implemented like C<attr> above but
names C<meta>, or just be new values in the C<attr> hashes.

=item Other spreadsheet formats

I consider adding CSV

=item Safety / flexibility

Make the different formats/modules just load if available and ignore if
not available.

=item OO-ify

Consider making the ref an object, though I currently don't see the big
advantage (yet). Maybe I'll make it so that it is a hybrid functional /
OO interface.

=back

=head1 SEE ALSO

=over 2

=item Spreadsheet::ParseExcel

http://search.cpan.org/~kwitknr/

=item Spreadsheet::ReadSXC

http://search.cpan.org/~terhechte/

=item Text::CSV_XS, Text::CSV

http://search.cpan.org/~jwied/
http://search.cpan.org/~alancitt/

=back

=head1 AUTHOR

H.Merijn Brand, <h.m.brand@xs4all.nl>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2005-2005 H.Merijn Brand

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself. 

=cut
