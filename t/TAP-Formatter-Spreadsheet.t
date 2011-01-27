use strict;
use warnings;

use lib 'lib';
use Test::More tests=>27;
use TAP::Harness;
BEGIN { use_ok('TAP::Formatter::Spreadsheet') };

my @tests = glob("t/data/*.pl");

is(scalar @tests, 11, "Number of sub test scripts");

my $fmt = new_ok( "TAP::Formatter::Spreadsheet" );

is($fmt->verbose,0, "Default Verbose Value");

is($fmt->filename, '__example.xls', "Default excel sheet name");

is($fmt->number_of_sheets, 1, "Default number of sheets");

is($fmt->sheetname_format, 'range', "Default sheetname format");

is($fmt->chart_name, "Pie_Chart", "Default Chart Sheet Name");

is($fmt->chart_type, "pie",  "Default Chart type");

is($fmt->summary_sheet_name, "Summary", "Default Summary Sheet Name");

$fmt->verbose(1);
is($fmt->verbose,1, "Changed Verbose Value");
$fmt->verbose(0);

$fmt->filename('abc.xls');
is($fmt->filename, 'abc.xls', "Changed excel sheet name");

$fmt->number_of_sheets(4);
is($fmt->number_of_sheets, 4, "Changed number of sheets");

$fmt->sheetname_format('first');
is($fmt->sheetname_format, 'first', "Changed sheetname format");

$fmt->chart_name("Pie");
is($fmt->chart_name, "Pie", "Changed Chart Sheet Name");

$fmt->chart_type("bar");
is($fmt->chart_type, "bar",  "Changed Chart type");

$fmt->summary_sheet_name("Tests_Summary");
is($fmt->summary_sheet_name, "Tests_Summary", "Changed Summary Sheet Name");

my $harness = TAP::Harness->new({ formatter => $fmt, merge => 1 });
my $a = $harness->runtests(@tests);

# Test Stats Verification

is(scalar $a->failed, 50, "Number of failed tests");
is(scalar $a->parse_errors, 2, "Number of Parse errors");
is(scalar $a->passed, 486, "Number of Passed Tests");
is(scalar $a->total, 536, "Number of Total Tests");
is(scalar  $a->skipped, 4, "Number of Skipped Tests");
is(scalar  $a->todo, 10, "Number of todo Tests");
is(scalar  $a->wait, 273664, "Overall wait status");
is(scalar  $a->exit, 1069, "overall exit status");

# Sheet creation & Deletion
ok(-e "abc.xls", "Excel Sheet is present");
unlink("abc.xls");
ok(not (-e "abc.xls"),"Succesfully deleted abc.xls after testing");




    
    