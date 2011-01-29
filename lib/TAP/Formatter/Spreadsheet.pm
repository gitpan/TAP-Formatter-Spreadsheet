package TAP::Formatter::Spreadsheet;
use strict;
use warnings;
use Carp;
use TAP::Base;    #base class for all TAP::*
use Data::Dumper;
use Spreadsheet::WriteExcel;    # Obviously for writing excel
use Spreadsheet::WriteExcel::Utility qw( xl_range_formula xl_rowcol_to_cell );

our $VERSION = '0.02';

BEGIN {
    our @ISA = qw(TAP::Base);

    #To Pacify Test Harness module and to create accessors
    my @ACCESSOR =
      qw( file parser show_count results test verbose verbosity number_of_sheets sheetname_format filename test_header_format test_todo_ok_format test_todo_not_ok_format test_skip_ok_format test_ok_format test_not_ok_format test_plan_format test_comment_format chart_name summary_sheet_name chart_type summary_header_format summary_format);

    for my $method (@ACCESSOR) {
        no strict 'refs';
        *$method = sub {
            my $class = shift;
            if (@_) {
                $class->{$method} = shift;
            }
            return $class->{$method};
        };
    }
}

##### Initialize default format values

sub _initialize {
    my ( $self, $arg_for ) = @_;
    $arg_for ||= {};
    $self->SUPER::_initialize($arg_for);
	$self->verbosity(0);
	$self->verbose(0);
    $self->filename('__example.xls');
    $self->number_of_sheets(1);
    $self->sheetname_format('range');
    $self->chart_name("Pie_Chart");
    $self->chart_type("pie");
    $self->summary_sheet_name("Summary");
    $self->test_todo_ok_format(
        { 'bg_color' => 'silver', 'color' => 'pink', 'border' => 1 } );
    $self->test_todo_not_ok_format(
        { 'bg_color' => 'silver', 'color' => 'orange', 'border' => 1 } );
    $self->test_skip_ok_format(
        { 'bg_color' => 'silver', 'color' => 'yellow', 'border' => 1 } );
    $self->test_ok_format(
        { 'bg_color' => 'silver', 'color' => 'green', 'border' => 1 } );
    $self->test_not_ok_format(
        { 'bg_color' => 'silver', 'color' => 'red', 'border' => 1 } );
    $self->test_plan_format(
        { 'bg_color' => 'silver', 'color' => 'magenta', 'border' => 1 } );
    $self->test_comment_format(
        { 'bg_color' => 'silver', 'color' => 'purple', 'border' => 1 } );
    $self->summary_header_format(
        {
            'bg_color' => 'green',
            'color'    => 'black',
            'border'   => 2,
            'bold'     => 1,
            'align'    => 'center'
        }
    );
    $self->summary_format(
        { 'bg_color' => 'silver', 'color' => 'purple', 'border' => 1 } );

    foreach my $key ( keys %$arg_for ) {
        $self->$key( $arg_for->{$key} ) if ( $self->can($key) );
    }
    return $self;
}

##### Called by Test::Harness before any test output is generated.

sub prepare {
    my ( $self, @tests ) = @_;
    $self->log("The tests which are executed are @tests");
}

#Called to create a new test session.

sub open_test {
    my ( $self, $test, $parser ) = @_;
    my $class = 'TAP::Formatter::Spreadsheet::Session';
    eval "require $class";
    $self->_croak($@) if $@;
    my $session = $class->new(
        {
            test      => $test,
            parser    => $parser,
            formatter => $self,
            verbose   => $self->verbose
        }
    );

    push @{ $self->{sessions} }, $session;
    return $session;
}

#C<summary> prints the summary report after all tests are run.  The argument is an aggregate.

sub summary {
    my ( $self, $aggregate ) = @_;
    my @t       = $aggregate->descriptions;
    my $tests   = [@t];
    my $runtime = $aggregate->elapsed_timestr;
    my $total   = $aggregate->total;
    my $passed  = $aggregate->passed;
    my $r       = $self->report($aggregate);
    $self->process($r);
}

# This subroutine is used to create a report using aggregate object

sub report {
    my ( $self, $a ) = @_;
    my $r = {
        tests        => [],
        start_time   => '?',
        end_time     => '?',
        elapsed_time => $a->elapsed_timestr,
    };

    for my $key (
        qw(total has_errors has_problems failed parse_errors passed skipped todo todo_passed wait exit)
      )
    {
        $r->{$key} = $a->$key;
    }
    $r->{actual_passed} = $r->{passed} + $r->{todo_passed};
    if ( $r->{total} ) {
        $r->{percent_passed} =
          sprintf( '%.1f', $r->{actual_passed} / $r->{total} * 100 );
    }
    else {
        $r->{percent_passed} = 0;
    }
    $r->{num_files} = scalar @{ $self->{sessions} };
    my $total_time = 0;
    foreach my $s ( @{ $self->{sessions} } ) {
        my $sr = $s->process_session_report;
        push @{ $r->{tests} }, $sr;
        $total_time += $sr->{elapsed_time} || 0;
    }
    $r->{total_time} = $total_time;
    return $r;
}

#### subroutine process consist of three stages:
####   i) Create Test summary page with Chart of type bar or pie along with overall stats.
####  ii) Prepare Individual Test reports with paging formats and coloring formats along with links to the charts.
#### iii) Prepare Charts for individual tests in the final sheet

sub process {
    my $self         = shift;
    my $r            = shift;
    my $file_name    = $self->filename || croak "Excel File name is empty\n";
    my $workbook     = Spreadsheet::WriteExcel->new($file_name);
    my $total_sheets = my $ts = $self->number_of_sheets
      || croak "Number of sheets is empty\n";
    my $sheetname_format = $self->sheetname_format
      || croak "Sheet name is empty\n";

    my $total_tests          = scalar @{ $r->{tests} };
    my $num_of_test_per_page = int( $total_tests / $total_sheets );

    $self->log("Test summary is not generated because of no tests being ran!!!")
      if ( ( $r->{percent_passed} + 0 ) == 0 );

    $self->process_test_summary_page( $workbook, $r );

    $self->log("Total Number of tests are: $total_tests");
    $self->log("Tests per page is $num_of_test_per_page");
    $self->log("Sheets: $total_sheets");
    my $plotter;
    my $i          = 0;
    my $item_no    = 0;
    my $item_track = 0;
    while ($total_sheets) {
        my $max = 0;    # Maximum column width
        my $l_of_test_per_page =
            ( $total_sheets == 1 )
          ? ( $num_of_test_per_page + ( $total_tests % $ts ) )
          : $num_of_test_per_page;

        #Get Sheet name
        my $sheetname =
          $self->get_sheetname( $sheetname_format, $r, $i + 1,
            $i + $l_of_test_per_page );

        my $worksheet = $workbook->add_worksheet($sheetname);
        my $row       = 0;
        foreach
          my $tests ( @{ $r->{tests} }[ $i .. $i + $l_of_test_per_page - 1 ] )
        {
            my ( $pass, $fail );
            $row += 5 if ($row);    #Spacing between seq tests
                  # Populate Pass/Fail/Test values for future Chart usage.
            if ( not $tests->{tests_run} ) {
                $pass = $fail = "NA";
            }
            else {
                $pass = $tests->{percent_actual_passed};
                $fail = 100 - $pass;
                $fail = 0 if ( $fail < 0 );
                $item_track++;
            }
            $plotter->{ $item_no++ } = {
                'passed' => $pass,
                'failed' => $fail,
                'tests'  => $tests->{test},
                'item'   => $item_track
            };

            ### Process Individual Test Header
            my $hbg = $tests->{test_status} eq 'passed' ? "green" : "red";
            my $hf = $self->test_header_format
              || {
                'bg_color' => $hbg,
                'color'    => 'black',
                'border'   => 2
              };
            my $htext = "Test Summary for $tests->{test}";
            $max = length($htext) if $max < length($htext);
            my $header_format = $workbook->add_format(%$hf);
            $worksheet->write( $row, 0, $htext, $header_format );

            my $link =
                "internal:"
              . $self->chart_name . "!"
              . xl_rowcol_to_cell( 25 * ( $item_track - 1 ), 0 );
            if ( $pass ne 'NA' ) {
                $worksheet->write( $row + 3, 2, $link, 'Chart Link' );
            }
            else {
                $worksheet->write( $row + 3, 2, 'No Chart' );
                $worksheet->write_comment(
                    $row + 3,
                    2,
"No Chart due to No of tests executed is $tests->{tests_run}"
                );
            }

            ### Process Individual Test Results body part
            my $res_count = scalar @{ $tests->{results} };
            $row += 1;
            foreach my $res ( @{ $tests->{results} } ) {
                no strict 'refs';
                next if ( $res->{type} eq 'unknown' && $res->{raw} eq '' );

                ## Format for sheet but sustained only for result rows
				unless ($res->{font_help}) {
					$res->{font_help} = 'comment';
				}
				my $lsub = "test_" . $res->{font_help} . '_format';
                my $bformat =
                  $self->$lsub;    ### Ugly hack to get the proper format
                my $format = $workbook->add_format(%$bformat);
                my $text   = $res->{raw};
                $max = length($text) if $max < length($text);
                $worksheet->write( $row++, 0, $res->{raw}, $format );
            }

            ### Process Statistics for individual tests
            my $stat_format = $workbook->add_format(
                color    => 'lime',
                bg_color => 'navy',
                border   => 1
            );
            $worksheet->write( $row++, 0, "Tests Run			: $tests->{tests_run}",
                $stat_format );
            $worksheet->write( $row++, 0,
                "Actual Passed		: $tests->{actual_passed}", $stat_format );
            $worksheet->write( $row++, 0, "Overall Passed		: $tests->{passed}",
                $stat_format );
            my $passed =
              ( $tests->{tests_run} ) ? $tests->{percent_actual_passed} : "NA";
            $worksheet->write_string( $row++, 0, "Passed \%			: $passed",
                $stat_format );
            $worksheet->write( $row++, 0, "Exit Status			: $tests->{exit}",
                $stat_format );
            $worksheet->write( $row++, 0, "Wait Status			: $tests->{wait}",
                $stat_format );

            # Fix the max space for the column based on max text length
            $worksheet->set_column( 0, 0, $max );
        }
        $i += $num_of_test_per_page;
        $total_sheets--;
    }

    # Plot the charts
    $self->plot_chart( $workbook, $plotter, 0, 4, 0, '' )
      ; #Workbook, plothash, Category start column, Chart insert column, Hiddenflag, Worksheet, If worksheet null get value for sheet from sheetname
    $workbook->close();
}

sub get_sheetname {
    my $self = shift;
    my ( $sheetname_format, $r, $first, $last ) = @_;
    my $sheetname;
    if ( $sheetname_format =~ /range/i ) {
        $sheetname = $first . "-" . $last;
        $sheetname = $first if ( $first == $last );
    }
    elsif ( $sheetname_format =~ /first/i ) {
        $sheetname = $r->{tests}->[ $first - 1 ]->{test};
        $sheetname =~ s/\/|\\|:/-/g;    #Remove special characters i.e / \ :
    }
    elsif ( $sheetname_format =~ /last/i ) {
        $sheetname = $r->{tests}->[ $last - 1 ]->{test};
        $sheetname =~ s/\/|\\|:/-/g;
    }
    return $sheetname;
}

sub process_test_summary_page {
    my ( $self, $workbook, $r ) = @_;
    my $plotter;
    my $sheetname = $self->summary_sheet_name
      || croak "Summary sheet name is empty\n";
    my $worksheet = $workbook->add_worksheet($sheetname);
    my $row       = 3;
    my $hf        = $self->summary_header_format;
    my $sf        = $self->summary_format;
    my $pass      = sprintf( "%.1f", $r->{percent_passed} );
    $pass += 0;
    my $fail = 100 - $pass;
    $plotter->{0} = {
        'passed' => $pass,
        'failed' => $fail,
        'tests'  => "Summary Report",
        'item'   => 1
    };
    $self->plot_chart( $workbook, $plotter, 3, 6, 1, $worksheet )
      ; #Workbook, plothash, Category start column, Chart insert column, Hiddenflag, worksheet
    my $header_format  = $workbook->add_format(%$hf);
    my $summary_format = $workbook->add_format(%$sf);
    my $time_format    = $workbook->add_format( %$sf, 'align' => 'right' );
    $worksheet->merge_range( $row, 0, $row, 1, "Test Summary", $header_format );
    $worksheet->write( ++$row, 0, "Total Files:\t",        $summary_format );
    $worksheet->write( $row,   1, $r->{num_files},         $summary_format );
    $worksheet->write( ++$row, 0, "Total Tests:\t",        $summary_format );
    $worksheet->write( $row,   1, $r->{total},             $summary_format );
    $worksheet->write( ++$row, 0, "Total Tests Passed:\t", $summary_format );
    $worksheet->write( $row,   1, $r->{passed},            $summary_format );
    $worksheet->write( ++$row, 0, "Total Tests failed:\t", $summary_format );
    $worksheet->write( $row,   1, $r->{failed},            $summary_format );
    $worksheet->write( ++$row, 0, "Total To-Dos:\t",       $summary_format );
    $worksheet->write( $row,   1, $r->{todo},              $summary_format );
    $worksheet->write( ++$row, 0, "Total Skipped:\t",      $summary_format );
    $worksheet->write( $row,   1, $r->{skipped},           $summary_format );
    $worksheet->write( ++$row, 0, "Total Parse Errors:\t", $summary_format );
    $worksheet->write( $row,   1, $r->{parse_errors},      $summary_format );
    $worksheet->write( ++$row, 0, "Exit Status:\t",        $summary_format );
    $worksheet->write( $row,   1, $r->{exit},              $summary_format );
    $worksheet->write( ++$row, 0, "Wait Status:\t",        $summary_format );
    $worksheet->write( $row,   1, $r->{wait},              $summary_format );
    $worksheet->write( ++$row, 0, "Elapsed time:\t",       $summary_format );
    my $ws = $r->{elapsed_time};
    $ws =~ s/wallclock sec(s)?/ws/g;
    $ws =~ s/\(.*\)//g;
    $ws =~ s/\s*$//g;
    $worksheet->write( $row,   1, $ws,                      $time_format );
    $worksheet->write( ++$row, 0, "Percentage Passed %:\t", $summary_format );
    $worksheet->write( $row,   1, $pass,                    $summary_format );
    $worksheet->set_column( 0, 0, 20 );
}

sub plot_chart {
    my ( $self, $workbook, $plotter, $cat_column, $chart_column, $hide,
        $chartsheet )
      = @_;
    my ( $chartsheet_name, $chart_type );
    if ($chartsheet) {
        $chartsheet_name = $chartsheet->get_name;
    }
    else {
        $chartsheet_name = $self->chart_name;
        $chartsheet      = $workbook->add_worksheet($chartsheet_name);
    }

    $chart_type = lc $self->chart_type;
    croak "Chart type must be either pie or bar\n"
      unless ( $chart_type =~ /^(pie|bar)$/i );
    foreach my $in ( sort { $a <=> $b } keys %$plotter ) {
        next if ( $plotter->{$in}->{passed} eq 'NA' );
        my $format = $workbook->add_format( 'hidden' => $hide );
        my $k = 25 * ( $plotter->{$in}->{item} - 1 );
        my $data = [
            [ 'Passed',                  'Failed' ],
            [ $plotter->{$in}->{passed}, $plotter->{$in}->{failed} ],
        ];

        $chartsheet->write( $k, $cat_column, $data, $format );
        my $chart = $workbook->add_chart( type => $chart_type, embedded => 1 );
        $chart->add_series(
            categories => xl_range_formula(
                $chartsheet_name, $k, $k + 1, $cat_column, $cat_column
            ),
            values => xl_range_formula(
                $chartsheet_name,
                $k,
                $k + 1,
                $cat_column + 1,
                $cat_column + 1
            ),
            name => $plotter->{$in}->{tests},
        );
        $chart->set_title( name => $plotter->{$in}->{tests} );
        $chart->set_x_axis( name => 'Status' );
        $chart->set_y_axis( name => 'Percentage (%)' );
        $chartsheet->insert_chart( $k, $chart_column, $chart );
    }
}

sub log {
    my $self   = shift;
    my $string = shift;
    print $string, "\n" if ( $self->verbose );
}


1;

=head1 NAME

TAP::Formatter::Spreadsheet - Perl extension for formatting TAP Test Harness output to Spreadsheet excel output

=head1 SYNOPSIS

	use TAP::Formatter::Spreadsheet;
	my $fmt = TAP::Formatter::Spreadsheet->new;
	$fmt->filename('abc.xls');
	$fmt->number_of_sheets(4);
	$fmt->sheetname_format('first');
	$fmt->chart_name("Pie");
	$fmt->chart_type("bar");
	$fmt->summary_sheet_name("Tests_Summary");

	my @tests = glob( 't/*.t' );
	my $harness = TAP::Harness->new({ formatter => $fmt, merge => 1 });
	$harness->runtests( @tests );

=head1 DESCRIPTION

This Module provides Excel output with customizable formats for the Test::Harness output. This Module can be 
use for formatting the test results along with creating bar charts or pie charts based on the output results
both for overall test summary and for individual tests.

This Module is based on TAP::Formatter::Console and TAP::Formatter::Base modules which are part of Core Perl
distribution. TAP::Base is used as base class for this implementation. 

=head1 METHODS

=head2 CONSTRUCTOR

=head3 new

  my $fmt = $class->new({ %args });

=head2 ACCESSORS

=head3 verbose

     1   verbose        Print individual test results (and more) to STDOUT.
     0   normal

=head3 filename

C<filename> is used as accessor for the excel filename where the results are written. The default value is
C<__example.xls>

$fmt->filename("hi.xls")

=head3 number_of_sheets

C<filename> is used as accessor for the number of sheets to be used to distribute the individual test results
in the excel file. The default value is C<1>

$fmt->number_of_sheets(4)

=head3 sheetname_format

C<sheetname_format> is used as accessor for the naming the sheet names which are used to distribute the 
individual test results in the excel file. The default value is C<range>

List of possible values are: 'first', 'last', 'range'

first -> First testname in the sheet is used as sheet name
last  -> Last testname in the sheet is used as sheet name
range - Numeric range of the tests which are present in that sheet is used as sheet name

$fmt->sheetname_format('first');

=head3 chart_name

C<chart_name> is used as accessor for the naming the sheet where the charts are drawn for  
individual test results in the excel file. The default value is C<Pie_Chart>

$fmt->chart_name("Pie");

=head3 chart_type

C<chart_type> is used as accessor for the type of the charts which are drawn for  
individual test results in the excel file. The default value is C<pie>

List of possible values are: 'bar', 'pie'

pie - Pie Charts are drawn
bar - Bar Charts are drawn

$fmt->chart_type("bar");


=head3 summary_sheet_name

C<summary_sheet_name> is used as accessor for the naming the sheet where the Overall test summaries for  
individual test results are written in the excel file. The default value is C<Summary>

$fmt->summary_sheet_name("Tests_Summary");

=head3 test_todo_ok_format

C<test_todo_ok_format> is used to set the format used for successful to-do.

Default Value is 'bg_color' => 'silver', 'color' => 'pink', 'border' => 1 .

    $fmt->test_todo_ok_format(
        { 'bg_color' => 'magenta', 'color' => 'orange', 'border' => 1 } );

=head3 test_todo_not_ok_format

C<test_todo_not_ok_format> is used to set the format used for unsuccessful to-do.

Default Value is 'bg_color' => 'silver', 'color' => 'orange', 'border' => 1 

    $fmt->test_todo_not_ok_format(
        { 'bg_color' => 'magenta', 'color' => 'pink', 'border' => 1 } );

=head3 test_skip_ok_format

C<test_skip_ok_format> is used to set the format used for successful skip values.

Default Value is 'bg_color' => 'silver', 'color' => 'yellow', 'border' => 1

    $fmt->test_skip_ok_format(
        { 'bg_color' => 'yellow', 'color' => 'pink', 'border' => 1 } );

=head3 test_ok_format

C<test_ok_format> is used to set the format used for successful test 'ok' values.

Default Value is 'bg_color' => 'silver', 'color' => 'green', 'border' => 1

	$fmt->test_ok_format(
        { 'bg_color' => 'pink', 'color' => 'yellow', 'border' => 1 } );

=head3 test_not_ok_format

C<test_not_ok_format> is used to set the format used for unsuccessful test 'not ok' values.

Default Value is 'bg_color' => 'silver', 'color' => 'red', 'border' => 1 }

    $fmt->test_not_ok_format(
        { 'bg_color' => 'yellow', 'color' => 'red', 'border' => 1 } );

=head3 test_plan_format

C<test_plan_format> is used to set the format used for plan values.

Default Value is 'bg_color' => 'silver', 'color' => 'magenta', 'border' => 1 

	$fmt->test_plan_format(
        { 'bg_color' => 'green', 'color' => 'pink', 'border' => 1 } );

=head3 test_comment_format

C<test_comment_format> is used to set the format used for comment values.

Default Value is 'bg_color' => 'silver', 'color' => 'purple', 'border' => 1

	$fmt->test_comment_format(
        { 'bg_color' => 'pink', 'color' => 'blue', 'border' => 1 } );

=head3 summary_header_format

C<summary_header_format> is used to set the format used for summary page header values.

Default Value is 'bg_color' => 'green', 'color'    => 'black', 'border'   => 2, 'bold'     => 1, 'align'    => 'center'


	$fmt->summary_header_format(
        {
            'bg_color' => 'pink',
            'color'    => 'navy',
            'border'   => 1,
            'bold'     => 0,
            'align'    => 'left'
        }
    );

=head3 summary_format

C<summary_format> is used to set the format used for summary text/stat values.

Default Value is 'bg_color' => 'silver', 'color' => 'purple', 'border' => 1 

    $self->summary_format(
        { 'bg_color' => 'green', 'color' => 'red', 'border' => 1 } );

=head2 API METHODS

=head3 summary

  $html = $fmt->summary( $aggregator )

C<summary> produces a summary report after all tests are run.  C<$aggregator>
should be a L<TAP::Parser::Aggregator>.


=head1 SEE ALSO

L<TAP::Formatter::Console> - the default TAP formatter used by L<TAP::Harness>

L<Spreadsheet::WriteExcel> - Module used to write spreadsheet excel 

=head1 AUTHOR

Murugesan Kandasamy E<lt>Murugesan.Kandasamy@gmail.comE<gt>

=head1 ACKNOWLDEGEMENT

John Mcnamara - For his contributions towards Spreadsheet::* Module

Steve Purkis  - For his contributions towards TAP::Formatter::* Module

Prasad JP, Sivaraman M, Rajesh S and Murugaperumal R for their contribution towards testing.

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2011 by Murugesan Kandasamy

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself, either Perl version 5.6.0 or,
at your option, any later version of Perl 5 you may have available.
=cut


