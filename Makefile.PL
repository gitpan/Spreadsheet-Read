use ExtUtils::MakeMaker;

my @exe;
for ( [ "xlscat",	"Convert XLS/SXC to plain text or CSV"	],
      ) {
    prompt ("Do you want to install '$_->[0]' ($_->[1]) ? ", "n") =~ m/[Yy]/ and
	push @exe, "examples/$_->[0]";
    }
my %wm = (
    NAME         => "Spreadsheet::Read",
    ABSTRACT     => "Read the data from a spreadsheet",
    AUTHOR       => "H.Merijn Brand <merijn\@procura.nl>",
    VERSION_FROM => "Read.pm",
    EXE_FILES    => [ @exe ],
    PREREQ_FATAL => 0,
    PREREQ_PM    => {
	Text::CSV_XS		=> 0.00,
	Spreadsheet::ReadSXC	=> 0.12,
	Spreadsheet::ParseExcel	=> 0.26,
#	Spreadsheet::Perl	=> 0.00,	# Not yet

	# Core modules
	File::Temp		=> 0.14,
	},
    );

my $rv = WriteMakefile (%wm);
