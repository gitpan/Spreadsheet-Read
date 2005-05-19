use ExtUtils::MakeMaker;

my %wm = (
    NAME         => "Spreadsheet::Read",
    VERSION_FROM => "Read.pm",
    PREREQ_PM    => {
	Spreadsheet::ReadSXC	=> 0.12,
	Spreadsheet::ParseExcel	=> 0.26,
	},
    );

my $rv = WriteMakefile (%wm);
