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
    PREREQ_PM    => {
	Spreadsheet::ReadSXC	=> 0.12,
	Spreadsheet::ParseExcel	=> 0.26,
	},
    );

my $rv = WriteMakefile (%wm);
