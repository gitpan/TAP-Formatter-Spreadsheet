package TAP::Formatter::Spreadsheet::Session;

use base qw( TAP::Base );

our $VERSION = '0.01';

BEGIN {

    @ACCESSOR = qw( formatter parser show_count results test verbose);

    for my $method (@ACCESSOR) {
        no strict 'refs';
        *$method = sub {
            $class = shift;
            if (@_) {
                $class->{$method} = shift;
            }
            return $class->{$method};
        };
    }
}

sub _initialize {
    my ( $self, $args ) = @_;
    $args ||= {};
    $self->SUPER::_initialize($args);
    foreach my $arg (qw( test parser formatter results show_count verbose)) {
        $self->$arg( $args->{$arg} ) if defined $args->{$arg};
    }
    return $self;
}

#For each of the session TAP::Parser invokes this subroutine. We use this to form a nice DS to use it for later formatting etc.

sub result {
    my ( $self, $result ) = @_;

    # if test, segregate the details for todo, todonot, skip etc
    if ( $result->is_test ) {
        $self->log( $result->as_string );
        if ( $result->has_todo ) {
            if ( $result->is_actual_ok ) {
                $result->{todo_passed} = 1;
                $short = 'todo_ok';
            }
            else {
                $short = 'todo_not_ok';
            }
        }
        elsif ( $result->has_skip ) {
            $short = 'skip_ok';
        }
        elsif ( $result->is_actual_ok ) {
            $short = 'ok';
        }
        else {
            $short = 'not_ok';
        }

        # If passed
        if ( $result->is_ok ) {
            $self->{passed}++;
        }

        # If passed but unplanned
        if ( $result->is_ok || $result->is_unplanned && $result->is_actual_ok )
        {
            $self->{passed_including_unplanned}++;
        }
    }
    else {
        $self->log( $result->as_string );
    }
    if ( $result->is_plan ) {
        $short = "plan";
    }
    if ( $result->is_comment ) {
        $short = "comment";
    }
    $result->{font_help} = $short;
    push @{ $self->{results} }, $result;
    return;
}

### For each test, get each and every details from the parser and feed them to $r DS.
sub process_session_report {
    my ($self) = @_;
    my $parser = $self->parser;
    my $r      = {
        test    => $self->test,
        results => $self->{results},
    };

    # copy contents of Parser to result DS.
    for my $key (
        qw(tests_planned tests_run start_time end_time skip_all has_problems passed failed todo_passed actual_passed actual_failed wait exit)
      )
    {
        $r->{$key} = $parser->$key;
    }

    # helper details for each of pass/fail/error etc for each of the test...
    $r->{num_parse_errors} = scalar $parser->parse_errors;
    $r->{parse_errors}     = [ $parser->parse_errors ];
    $r->{passed_tests}     = [ $parser->passed ];
    $r->{failed_tests}     = [ $parser->failed ];

    $r->{test_status} = $r->{has_problems} ? 'failed' : 'passed';
    $r->{elapsed_time} = $r->{end_time} - $r->{start_time};

    if ( $r->{tests_planned} ) {
        my $num_passed        = $self->{passed}                     || 0;
        my $num_actual_passed = $self->{passed_including_unplanned} || 0;
        my $p                 = $r->{percent_passed} =
          sprintf( '%.1f', $num_passed / $r->{tests_planned} * 100 );
        $r->{percent_actual_passed} =
          sprintf( '%.1f', $num_actual_passed / $r->{tests_planned} * 100 );
        unless ( $p != 100 && $r->{skip_all} ) {
            $r->{percent_passed} = 0;
        }
    }
    return $r;
}

#### This will be called during every test getting closed but need to be present ####
sub close_test {
    my ( $self, @args ) = @_;
    return;
}

sub log {
    my $self   = shift;
    my $string = shift;
    print $string, "\n" if ( $self->verbose );
}
