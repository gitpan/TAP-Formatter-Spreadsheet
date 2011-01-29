use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    NAME              => 'TAP::Formatter::Spreadsheet',
    VERSION_FROM      => 'lib/TAP/Formatter/Spreadsheet.pm', # finds $VERSION
    PREREQ_PM         => {			'TAP::Parser' => 0,
						'Spreadsheet::WriteExcel' => 0,
								  'Test::Harness' => 0,}, 
    AUTHOR        => 'Murugesan Kandasamy (murugesan.kandasamy@gmail.com)',
    ABSTRACT      => 'TAP Formatter which writes to a cross platform Excel file',
);
