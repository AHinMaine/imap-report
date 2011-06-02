#!/usr/bin/env perl
##!/bin/sh #eval 'exec `which perl` -x -S $0 ${1+"$@"}'
#    if $runningUnderSomeShell;
#
#!perl
#----------------------------------------------------------------------
#
#         File:  imap-report.pl
#
#        Usage:  imap-report.pl
#
#  Description: Specifically designed to work with gmail, allows
#               you to run reports against an imap mailbox.
#
# Requirements:  Mail::IMAPClient
#                Term::Menus
#                Date::Manip
#
#       Author:  Andy Harrison
#
#                (tld:     com       )
#                (user:    aharrison )
#                (domain:  gmail     )
#
#      Version:  $Id$
#===============================================================================

use strict;
use warnings;

# {{{ progress bar
#
# This is just the String::ProgressBar module with one
# single extra line of code.
#

package IMAP::Progress;

our $VERSION='0.03';

use strict;
use Carp;
use vars qw($VERSION);

#
# LICENSE
# =======
# You can redistribute it and/or modify it under the conditions of LGPL.
#
# AUTHOR
# ======
# Andreas Hernitscheck  ahernit(AT)cpan.org



sub new { # $object ( max => $int )
    my $pkg = shift;
    my $self = bless {}, $pkg;
    my $v={@_};

    # default values
    my $def = {
                value           =>  0,
                length          =>  20,
                border_left     =>  '[',
                border_right    =>  ']',
                bar             =>  '=',
                show_rotation   =>  0,
                show_percent    =>  1,
                show_amount     =>  1,
                text            =>  '',
                info            =>  '',
                print_return    =>  0,
                text_length     =>  14,
               };

    # assign default values
    foreach my $k (keys %$def){
        $self->{ $k } = $def->{ $k };
    }

    if ( not $self->{"text"} ){
        $self->{"text"} = 0;
    }


    foreach my $k (keys %$v){
        $self->{ $k } = $v->{ $k };
    }



    my @req = qw( max );

    foreach my $r (@req){
        if ( ! $self->{$r} ){
            croak "\'$r\' required in constructor";
        }
    }

    return $self;
}


# updates the bar with a new value
# and returns the object itself.
sub update {
    my $self = shift;
    my $value = shift;

    if ( $value > $self->{'max'}  ){
        $value = $self->{'max'};
    }

    $self->{'value'} = $value;

    return $self;
}


# updates text (before bar) with a new value
# and returns the object itself.
sub text { # $object ($string)
    my $self = shift;
    my $value = shift;

    $self->{'text'} = $value;

    return $self;
}


# updates info (after bar) with a new value
# and returns the object itself.
sub info { # $object ($string)
    my $self = shift;
    my $value = shift;

    $self->{'info_last'} = $self->{'info'};
    $self->{'info'} = $value;


    return $self;
}



# Writes the bar to STDOUT.
sub write { # void ()
    my $self = shift;
    my $bar = $self->string();

    $|=1;
    print "$bar\r";

    if ( $self->{'print_return'} && ($self->{'value'} == $self->{'max'}) ){
        print "\n";
    }

}


# returns the bar as simple string, so you may write it by
# yourself.
sub string { # $string
    my $self = shift;
    my $str;

    my $ratio = $self->{'value'} / $self->{'max'};
    my $percent = int( $ratio * 100 );

    my $bar = $self->{'bar'} x ( $ratio *  $self->{'length'} );
    $bar .= " " x ($self->{'length'} - length($bar) );

    $bar = $self->{'border_left'} . $bar . $self->{'border_right'};

    $str = "$bar";

    if ( $self->{'show_percent'} ){
       $str.=" ".sprintf("%3s",$percent)."%";
    }

    if ( $self->{'show_amount'} ){
       $str.=" [".sprintf("%".length($self->{'max'})."s",$self->{'value'})."/".$self->{'max'}."]";
    }

    if ( $self->{'show_rotation'} ){
       my $char = $self->_getRotationChar();
       $str.=" [$char]";
    }

    if ( $self->{'info'} || $self->{'info_used'} ){
       $str.=" ".sprintf("%-".length($self->{'info_last'})."s", $self->{'info'});
       $self->{'info_used'} = 1;
    }



    if ( $self->{'text'} ){
       $str=sprintf("%-".$self->{'text_length'}."s", $self->{'text'})." $str";
    }

    return $str;
}

# Returns a rotating slash.
# With every call one step further
sub _getRotationChar {
    my $self  = shift;

    my @matrix = qw( / - \ | );

    $self->{rotation_counter} = ($self->{rotation_counter}+1) % (scalar(@matrix)-1);

    return $matrix[ $self->{rotation_counter} ];
}


1;

# }}}

package main;

use Data::Dumper;
use Getopt::Long qw/:config auto_help auto_version/;
use Storable qw/store retrieve/;

use Mail::IMAPClient;
use Term::ReadKey qw/GetTerminalSize/;
use Term::Menus;
use Date::Manip::Date;

our $VERSION = sprintf "%d.%d", q$Revision: 1.1 $ =~ /(\d+)/g;

# Handle signals gracefully
#
$SIG{'INT'}  = 'die_signal';
$SIG{'QUIT'} = 'die_signal';
$SIG{'USR1'} = 'die_signal';
#$SIG{'CHLD'} = 'IGNORE';
#$SIG{'CHLD'} = sub { print "\n\n!!CHLD SIG!!\n\n"; };
my $break = 0;

# {{{ Handle commandline options
#

my $opts = {};

$opts->{server}            = 'imap.gmail.com';
$opts->{port}              = 993;
$opts->{debug}             = 0;
$opts->{verbose}           = 0;
$opts->{pager}             = '/usr/bin/less';
$opts->{log}               = '/tmp/imap-report.debuglog';
$opts->{top}               = 10;      # default for top 10 style reports
$opts->{min}               = 100_000; # default minimum size in bytes
$opts->{Maxcommandlength}  = 1_000;   # Number of messages in a single fetch operation
$opts->{Keepalive}         = 1;
$opts->{Fast_io}           = 1;
$opts->{Reconnectretry}    = 3;
$opts->{Ssl}               = 1;
$opts->{cache}             = 1;
$opts->{cache_file}        = "$ENV{HOME}/.imap-report.cache";
$opts->{cache_age}         = 24 * 60 * 60;
$opts->{conf}              = "$ENV{HOME}/.imapreportrc";
$opts->{force}             = 0;
$opts->{list}              = 0;
$opts->{types}             = 0;
$opts->{threads}           = 0;
$opts->{use_threaded_mode} = 0;

GetOptions(

    $opts,

        'server=s',
        'port=i',
        'user=s',
        'password=s',
        'top=i',
        'min=i',
        'filters|folder|search=s@{,}',
        'exclude=s@{,}',
        'report=s',
        'types!',
        'list!',
        'log=s',
        'conf=s',
        'Keepalive!',
        'Fast_io!',
        'Maxcommandlength=i',
        'Reconnectretry=i',
        'Ssl!',
        'cache!',
        'cache_file=s',
        'cache_age=i',
        'threads=i',
        'use_threaded_mode!',
        'pager=s',
        'force!',
        'debug!',
        'verbose!',

);

if ( $opts->{list} ) {
    $opts->{report} = 'list';
}

if ( $opts->{types} ) {
    show_report_types();
}

if ( $opts->{Ssl} ) {
    eval {
        require IO::Socket::SSL;
        import IO::Socket::SSL;
        1;
    } || die_clean( 1, "IO::Socket::SSL module not found\n" );
}

read_config_file();

die "--server option required.\n" unless $opts->{server};
die "--port option required.\n"   unless $opts->{port};

# Default to unthreaded. Only turn on threading if our
# conditions are met...
#
my $use_threaded_mode = 0;

my @valid_threads = ( 2, 3, 5, 7, 11, 13, 17 );

if ( $opts->{threads} >= 2 && grep $opts->{threads} == $_, @valid_threads ) {
    if ( eval 'require threads;' ) {
        import threads ( 'exit' => 'threads_only' );
        require threads::shared;
        import threads::shared qw(shared);
        require Thread::Queue;
        import Thread::Queue ();
        $use_threaded_mode = 1;
    }
}


if ( ! defined $opts->{user} && ! $opts->{user} ) {
    $opts->{user} = user_prompt();
}

if ( ! defined $opts->{password} && ! $opts->{password} ) {
    $opts->{password} = password_prompt();
}

# Just shovel any remaining args onto the list of filters.
#
if ( defined @ARGV && scalar @ARGV ) {
    push @{$opts->{filters}}, $_ for @ARGV;
}

verbose( qq[

           Server: $opts->{server}
             Port: $opts->{port}
         Username: $opts->{user}

              min: $opts->{min}
              top: $opts->{top}
 Maxcommandlength: $opts->{Maxcommandlength}
          threads: $opts->{threads}
use_threaded_mode: $opts->{use_threaded_mode}

            cache: $opts->{cache}
       cache_file: $opts->{cache_file}
        cache_age: $opts->{cache_age}

          verbose: $opts->{verbose}


]);


# }}}

if ( $opts->{debug} ) {
    open( DBG, '>>' . $opts->{log} )
        or die_clean( 1, "Unable to open debuglog: $!\n" );
}

# Lazy imap to english translation table
#
my %header_table = (

                'Date'                          => 'INTERNALDATE',
                'INTERNALDATE'                  => 'Date',

                'Subject'                       => 'BODY[HEADER.FIELDS (SUBJECT)]',
                'BODY[HEADER.FIELDS (SUBJECT)]' => 'Subject',

                'Size'                          => 'RFC822.SIZE',
                'RFC822.SIZE'                   => 'Size',

                'To'                            => 'BODY[HEADER.FIELDS (TO)]',
                'BODY[HEADER.FIELDS (TO)]'      => 'To',

                'From'                          => 'BODY[HEADER.FIELDS (FROM)]',
                'BODY[HEADER.FIELDS (FROM)]'    => 'From',

);


print "Using IMAP Server: "
    . $opts->{server}
    . ':'
    . $opts->{port}
    . "\n"
    . "Connecting...\n"
    ;

my $imap = Mail::IMAPClient->new(

                        Server           => $opts->{server},
                        Port             => $opts->{port},
                        User             => $opts->{user},
                        Password         => $opts->{password},
                        Keepalive        => $opts->{Keepalive},
                        Fast_io          => $opts->{Fast_io},
                        Ssl              => $opts->{Ssl},
                        Reconnectretry   => $opts->{Reconnectretry},
                        Maxcommandlength => $opts->{Maxcommandlength},
                        Uid              => 0,
                        Clear            => 100,
                        Debug            => $opts->{debug},
                        Debug_fh         => *DBG,

) or die "Cannot connect to host : $@";

if ( $imap->IsAuthenticated ) {
    print "Login successful.\n";
}

our ( $progress_queue, $fetched_queue, $fetcher_status, $sequence_queue );

if ( $use_threaded_mode ) {
    $fetched_queue  = Thread::Queue->new();
    $fetcher_status = Thread::Queue->new();
    $progress_queue = Thread::Queue->new();
    $sequence_queue = Thread::Queue->new();
}

my $reports = report_types();

my $cache = init_cache( $opts->{cache_file} );

# We're connected to the imap server.  Fetch the list of
# folders now, rather than repeatedly running imap commands
# to gather the list of folders each time the menu is
# displayed.
#
my $imap_folders = fetch_folders({ filters => $opts->{filters}, excludes => $opts->{exclude} });

die_clean( 1, "No folders in fetched lists!" ) unless scalar(@$imap_folders);

# Keep looping until 'quit'
#
while (1) {

    my $banner =
        "\n(Current minimum size of message to report: "
        . $opts->{min}
        . ' bytes, '
        . convert_bytes( $opts->{min} ) . ")\n\n"
        .  '(Number of folders in list: ' . scalar(@$imap_folders) . ")\n\n\n"
        . "Choose your report type: \n\n\n";

    # Choose what type of report we want to run.
    #
    my $action = $opts->{report}
        ? $reports->{$opts->{report}}
        : choose_action({ banner  => $banner,
                          reports => $reports });

    die_clean( 0, 'Quitting...' ) if $action eq ']quit[';

    print "\n\n Action selected: $action\n";

    my @report;

    # Take action based on our chosen report type.  The
    # reports themselves don't print any output, but just
    # populate an array for display at the end of the job.
    #
    # TODO
    #
    # (This could be handled a LOT more elegantly...)
    #
    if ( $action eq $reports->{biggest_messages_report} ) {

        @report = biggest_messages_report();

    } elsif ( $action eq $reports->{size_report} ) {

        @report = size_report();

    } elsif ( $action eq $reports->{folder_message_count_report} ) {

        @report = folder_message_count_report();

    } elsif ( $action eq $reports->{folder_message_sizes_report} ) {

        @report = folder_message_sizes_report();

    } elsif ( $action eq $reports->{messages_by_subject_report} ) {

        #@report = messages_by_subject_report();
        @report = messages_by_header_report({ header => 'Subject' });

    } elsif ( $action eq $reports->{messages_by_from_address_report} ) {

        #@report = messages_by_from_address_report();

        @report = messages_by_header_report({ header => 'From' });

    } elsif ( $action eq $reports->{messages_by_to_address_report} ) {

        #@report = messages_by_to_address_report();

        @report = messages_by_header_report({ header => 'To' });

    } elsif ( $action eq $reports->{list} ) {

        @report = list_folders();

    }

    next unless scalar(@report);

    print_report(\@report);

}

1;

# {{{ subs
#

# {{{ report_types
#
sub report_types {

    # These are the types of reports that can be run.
    #
    return {
        folder_message_count_report     => 'Total count of messages in ALL folders',
        folder_message_sizes_report     => 'Total size of messages in ALL folders',
        messages_by_subject_report      => 'Folder report: All messages sorted by Subject',
        messages_by_from_address_report => 'Folder report: All messages sorted by From address',
        messages_by_to_address_report   => 'Folder report: All messages sorted by To address',
        biggest_messages_report         => 'Folder report: Largest messages',
        size_report                     => 'Folder report: Total size of messages',
        list                            => 'Display the current list of folders',
    };

} # }}}

# {{{ list_folders
#
sub list_folders {

    return join( "\n", @$imap_folders );

}

# }}}

# {{{ biggest_messages_report
#
# This report is to give us a top-ten style report to see
# the largest messages in a folder.  It disregards any
# messages below our 'min' size.
#
sub biggest_messages_report {

    my $report_type = $reports->{biggest_messages_report};

    my $folder = folder_choice({ folders => $imap_folders, report_type => $report_type });

    return unless $folder;

    my $num = $imap->message_count;

    print "Selected folder '$folder' contains $num messages\n\n";

    print "\n\nFetching message details...\n";

    my $stime = time;

    my $fetched_messages =
        $use_threaded_mode
        ? threaded_fetch_msgs( $folder )
        : fetch_msgs( $folder )
        ;

    #ddump( 'fetched_messages', $fetched_messages ) if $opts->{debug};

    my $ftime = time;

    my $elapsed = $ftime - $stime;

    my @breport;

    push @breport, "\nTotal time to fetch all messages: $elapsed seconds\n";
    push @breport, 'Iterated ' . scalar( keys %$fetched_messages ) . " messages.\n";
    push @breport, '(Ignored messages smaller than ' . $opts->{min} . " bytes.)\n";
    push @breport, "\nReporting on the top " . $opts->{top} . " messages.\n";

    my $totalsize;

    # quick calculation on the total size of the messages so
    # it can appear at the top of the report.
    #
    for ( keys %$fetched_messages ) {
        $totalsize += $fetched_messages->{$_}->{$header_table{'Size'}};
    }

    push @breport, 'Total folder size: ' . convert_bytes($totalsize) . "\n\n\n";

    my @msglist;

    my $counter = 0;
    my $reportsize;

    # Now iterate the messages, sorted by size, greatest to
    # smallest and just pull out the values we care about
    # from each message.  Increment the counter only for
    # messages that pass the size test.
    #
    for my $curmsg (
        reverse sort {
                $fetched_messages->{$a}->{$header_table{'Size'}}
            <=> $fetched_messages->{$b}->{$header_table{'Size'}} } keys %$fetched_messages
        ) {

        my $Size = $fetched_messages->{$curmsg}->{ $header_table{Size} };

        next if $Size <= $opts->{min};

        $reportsize += $Size;

        my $Subject = $fetched_messages->{$curmsg}->{ $header_table{Subject} };

        my $Date = $fetched_messages->{$curmsg}->{ $header_table{Date} };

        push @msglist, [ convert_bytes($Size), $Date, $Subject ];

        last if $counter++ >= ( $opts->{top} - 1 );

    }

    # If the counter never incremented, obviously no
    # messages were larger than our --min threshold.
    #
    if ( ! $counter ) {
        show_error(   'Error: No messages greater than minimum size ('
                    . $opts->{min}
                    . ' bytes) to report...' );
        next;
    } else {

        push @breport, "Size\t\tDate\t\t\t\tSubject\n";
        push @breport, '-' x 60 . "\n";

        for ( @msglist ) {
            push @breport,
                defined $_->[0] && $_->[0]
                ? $_->[0]
                : '',
                "\t\t",
                defined $_->[1] && $_->[1]
                ? $_->[1]
                : '',
                "\t",
                defined $_->[2] && $_->[2]
                ? $_->[2]
                : '',
                "\n"
                ;

            # Display the necessary search needed to cut and
            # paste into gmail directly to quickly find the
            # current message.
            #
            push @breport, generate_subject_search_string( $folder, $_->[1], $_->[2] ) . "\n\n"
                if $opts->{server} eq 'imap.gmail.com';

        }

        push @breport, '-' x 60 . "\n";
        push @breport, "Total size of reported messages in folder '$folder': " . convert_bytes($reportsize) . "\n\n";

    }

    return @breport;

} # }}}

# {{{ size_report
#
sub size_report {

    my $report_type = $reports->{size_report};

    my $folder = folder_choice({ folders => $imap_folders, report_type => $report_type });

    return unless $folder;

    print "\n\nFetching message details for folder '$folder'...\n\n";

    my $stime = time;

    my $msgs =
        $use_threaded_mode
        ? threaded_fetch_msgs( $folder )
        : fetch_msgs( $folder )
        ;

    my $ftime = time;

    my $elapsed = $ftime - $stime;

    my @sreport;

    push @sreport, "\nTime to fetch: $elapsed seconds\n";
    push @sreport, "Iterated " . scalar( keys %$msgs ) . " messages.\n\n";

    my $totalsize;

    # TODO
    #
    # need to clean up this counter garbage...
    #
    my $counter = 0;

    for ( keys %$msgs ) {
        $totalsize += $msgs->{$_}->{$header_table{'Size'}};
        $counter++;
    }

    #ddump( 'msgs', $msgs ) if $opts->{debug};


    if ( ! $counter ) {
        show_error(   'Error: No messages to report...' );
        next;
    }

    #ddump( 'reportsize', convert_bytes($totalsize) ) if $opts->{debug};

    push @sreport, '-' x 60 . "\n";
    push @sreport, "\n\nTotal size of all messages in '$folder' = " . convert_bytes($totalsize) . "\n\n";
    push @sreport, '-' x 60 . "\n";

    return @sreport;

} # }}}

# {{{ messages_by_subject_report
#
sub messages_by_subject_report {

    my $report_type = $reports->{messages_by_subject_report};

    my $folder = folder_choice({ folders => $imap_folders, report_type => $report_type });

    return unless $folder;

    my $stime = time;

    #my $msgs = fetch_msgs( $folder );
    #my $msgs = threaded_fetch_msgs( $folder );

    my $msgs =
        $use_threaded_mode
        ? threaded_fetch_msgs( $folder )
        : fetch_msgs( $folder )
        ;

    my $ftime = time;

    my $elapsed = $ftime - $stime;

    my @cur_report;

    my $subject_stats     = {};
    my $massaged_subjects = {};


    # This produces two separate reports.  One with raw
    # untouched subjects.  Another that strips off all the
    # Re: (and Re: Re: RE: Re: junk).
    #
    for ( keys %$msgs ) {

        my $subj = $msgs->{$_}->{$header_table{'Subject'}};

        if ( defined $subject_stats->{$subj} ) {
            $subject_stats->{$subj}++;
        } else {
            $subject_stats->{$subj} = 1;
        }

        $subj =~ s/^Re:\s+//gi;
        $subj =~ s/^Fwd:\s+//gi;
        $subj =~ s/\s+Re:\s+//gi;
        $subj =~ s/\s+Fwd:\s+//gi;

        if ( defined $massaged_subjects->{$subj} ) {
            $massaged_subjects->{$subj}++;
        } else {
            $massaged_subjects->{$subj} = 1;
        }


    }

    push @cur_report, "\n\nTotal messages processed: " . scalar(keys %$msgs) . "\n\n";
    push @cur_report, 'Top ' . $opts->{top} . " subjects: \n\n\n";
    push @cur_report, "Count\t\t\tSubject\n" . '-' x 60 . "\n";

    my $counter = 1;

    # Sort the first report numerically by the number of
    # occurrences of each subject to produce our top-ten
    # style report.
    #
    for ( reverse sort { $subject_stats->{$a} <=> $subject_stats->{$b} } keys %$subject_stats ) {
        push @cur_report, $subject_stats->{$_}
        . "\t\t\t"
        . $_
        . "\n"
        ;

        $counter++;

        last if $counter >= $opts->{top};

    }

    push @cur_report, "\n\n" . '-' x 60 . "\n\n";

    push @cur_report,
          "\n\n" . 'Top '
        . $opts->{top}
        . " report after a bit of subject massage...\n\n";

    push @cur_report, "Count\t\t\tSubject\n" . '-' x 60 . "\n";

    my $mcounter = 1;

    # Sort the second report numerically by the number of
    # occurrences of each massaged subject to produce our
    # top-ten style report.
    #
    for ( reverse sort { $massaged_subjects->{$a} <=> $massaged_subjects->{$b} } keys %$massaged_subjects ) {

        push @cur_report, $massaged_subjects->{$_}
        . "\t\t\t"
        . $_
        . "\n"
        ;

        $mcounter++;

        last if $mcounter >= $opts->{top};

    }


    # Finally, just shovel onto the report ALL message
    # subjects, sorted alphabetically, to page through it
    # manually.
    #
    push @cur_report,
          "\n\n\n"
        . '-' x 60 . "\n\n"
        . "All messages, sorted by subject:\n\n\n"
        . "Date\t\t\t\t\tSize\t\tSubject\n"
        . '-' x 60 . "\n\n";

    for ( sort { $msgs->{$a}->{$header_table{'Subject'}} cmp $msgs->{$b}->{$header_table{'Subject'}} } keys %$msgs ) {
        push @cur_report,
            $msgs->{$_}->{$header_table{'Date'}}
            . "\t\t"
            . convert_bytes( $msgs->{$_}->{$header_table{'Size'}} )
            . "\t\t"
            . $msgs->{$_}->{$header_table{'Subject'}}
            . "\n"
            ;

    }


    push @cur_report, "\n\n\nTime to fetch: $elapsed seconds\n";
    push @cur_report, "Iterated " . scalar( keys %$msgs ) . " messages.\n\n";

    return @cur_report;

} # }}}

# {{{ messages_by_header_report
#
# Report sorted by whatever header...
#
sub messages_by_header_report {

    my $args = shift;

    my $header = $args->{header};

    return unless $header;

    return unless grep $header eq $_, keys %header_table;

    my @cur_report;

    my $report_type = $reports->{messages_by_from_address_report}
        if $header eq 'From';

    my $folder = folder_choice({ folders => $imap_folders, report_type => $report_type });

    return unless $folder;

    my $stime = time;

    my $num = $imap->message_count;

    my $msgs =
        $use_threaded_mode
        ? threaded_fetch_msgs( $folder )
        : fetch_msgs( $folder )
        ;

    my $ftime = time;
    my $elapsed = $ftime - $stime;

    my $header_stats = {};

    for ( keys %$msgs ) {

        my $imap_header = $msgs->{$_}->{ $header_table{$header} };

        if ( defined $header_stats->{$imap_header} ) {
            $header_stats->{$imap_header}++;
        } else {
            $header_stats->{$imap_header} = 1;
        }

    }

    push @cur_report, 'Total messages processed: ' . scalar( keys %$msgs ) . "\n\n";
    push @cur_report, 'Top ' . $opts->{top} . " $header addresses: \n\n\n";
    push @cur_report, "Count\t\t\t$header\n" . '-' x 60 . "\n";

    my $counter = 1;

    # sort the header field by the count of their
    # occurrences.
    #
    for ( reverse sort { $header_stats->{$a} <=> $header_stats->{$b} } keys %$header_stats ) {

        push @cur_report, $header_stats->{$_}
        . "\t\t\t"
        . $_
        . "\n"
        ;

        $counter++;

        last if $counter >= $opts->{top};

    }

    push @cur_report, "\n\n" . '-' x 60 . "\n\n";

    push @cur_report,
          "\n\n\n"
        . '-' x 60 . "\n\n"
        . "All messages, sorted by $header\n\n\n"
        . "Date\t\t\t\t\tSize\t\t$header\n"
        . '-' x 60 . "\n\n";

    # The entire list of messages, sorted by the chosen
    # header field.
    #
    for ( sort { $msgs->{$a}->{$header_table{$header}} cmp $msgs->{$b}->{$header_table{$header}} } keys %$msgs ) {

        push @cur_report,
            $msgs->{$_}->{$header_table{'Date'}}
            . "\t\t"
            . convert_bytes( $msgs->{$_}->{$header_table{'Size'}} )
            . "\t\t"
            . $msgs->{$_}->{$header_table{$header}}
            . "\n"
            ;

    }

    push @cur_report, "\n\n\nTime to fetch: $elapsed seconds\n";
    push @cur_report, "Iterated " . scalar( keys %$msgs ) . " messages.\n\n";

    return @cur_report;

} # }}}

# {{{ messages_by_from_address_report
#
# Same to the subjects report, but changed subjects to from
# addresses and no massaging business.
#
sub messages_by_from_address_report {

    my $report_type = $reports->{messages_by_from_address_report};

    my $folder = folder_choice({ folders => $imap_folders, report_type => $report_type });

    return unless $folder;

    my $stime = time;

    #my $msgs = fetch_msgs( $folder );
    #my $msgs = threaded_fetch_msgs( $folder );
    my $msgs =
        $use_threaded_mode
        ? threaded_fetch_msgs( $folder )
        : fetch_msgs( $folder )
        ;

    my $ftime = time;

    my $elapsed = $ftime - $stime;

    my @cur_report;

    my $from_stats = {};

    for ( keys %$msgs ) {

        my $from = $msgs->{$_}->{$header_table{'From'}};

        if ( defined $from_stats->{$from} ) {
            $from_stats->{$from}++;
        } else {
            $from_stats->{$from} = 1;
        }

    }

    push @cur_report, 'Total messages processed: ' . scalar(keys %$msgs) . "\n\n";
    push @cur_report, 'Top ' . $opts->{top} . " from addresses: \n\n\n";
    push @cur_report, "Count\t\t\tFrom\n" . '-' x 60 . "\n";

    my $counter = 1;

    # sort the from address by the count of their
    # occurrences.
    #
    for ( reverse sort { $from_stats->{$a} <=> $from_stats->{$b} } keys %$from_stats ) {
        push @cur_report, $from_stats->{$_}
        . "\t\t\t"
        . $_
        . "\n"
        ;

        $counter++;

        last if $counter >= $opts->{top};

    }

    push @cur_report, "\n\n" . '-' x 60 . "\n\n";


    push @cur_report,
          "\n\n\n"
        . '-' x 60 . "\n\n"
        . "All messages, sorted by from:\n\n\n"
        . "Date\t\t\t\t\tSize\t\tFrom\n"
        . '-' x 60 . "\n\n";

    # The entire list of from addresses, sorted by from
    # address alphabetically.
    #
    for ( sort { $msgs->{$a}->{$header_table{'From'}} cmp $msgs->{$b}->{$header_table{'From'}} } keys %$msgs ) {
        push @cur_report,
            $msgs->{$_}->{$header_table{'Date'}}
            . "\t\t"
            . convert_bytes( $msgs->{$_}->{$header_table{'Size'}} )
            . "\t\t"
            . $msgs->{$_}->{$header_table{'From'}}
            . "\n"
            ;

    }


    push @cur_report, "\n\n\nTime to fetch: $elapsed seconds\n";
    push @cur_report, "Iterated " . scalar( keys %$msgs ) . " messages.\n\n";

    return @cur_report;

} # }}}

# {{{ messages_by_to_address_report
#
# Cut and paste of the above report.
#
# TODO
#
# Abstract this function so that it can be used for any
# header field.
#
sub messages_by_to_address_report {

    my $report_type = $reports->{messages_by_to_address_report};

    my $folder = folder_choice({ folders => $imap_folders, report_type => $report_type });

    return unless $folder;

    my $stime = time;

    #my $msgs = fetch_msgs( $folder );
    #my $msgs = threaded_fetch_msgs( $folder );

    my $msgs =
        $use_threaded_mode
        ? threaded_fetch_msgs( $folder )
        : fetch_msgs( $folder )
        ;

    my $ftime = time;

    my $elapsed = $ftime - $stime;

    my @cur_report;

    my $to_stats = {};

    for ( keys %$msgs ) {

        my $to = $msgs->{$_}->{$header_table{'To'}};

        if ( defined $to_stats->{$to} ) {
            $to_stats->{$to}++;
        } else {
            $to_stats->{$to} = 1;
        }

    }

    push @cur_report, 'Total messages processed: ' . scalar(keys %$msgs) . "\n\n";
    push @cur_report, 'Top ' . $opts->{top} . " to addresses: \n\n\n";
    push @cur_report, "Count\t\t\tTo\n" . '-' x 60 . "\n";

    my $counter = 1;

    for ( reverse sort { $to_stats->{$a} <=> $to_stats->{$b} } keys %$to_stats ) {
        push @cur_report, $to_stats->{$_}
        . "\t\t\t"
        . $_
        . "\n"
        ;

        $counter++;

        last if $counter >= $opts->{top};

    }

    push @cur_report, "\n\n" . '-' x 60 . "\n\n";


    push @cur_report,
          "\n\n\n"
        . '-' x 60 . "\n\n"
        . "All messages, sorted by to:\n\n\n"
        . "Date\t\t\t\t\tSize\t\tTo\n"
        . '-' x 60 . "\n\n";

    for ( sort { $msgs->{$a}->{$header_table{'To'}} cmp $msgs->{$b}->{$header_table{'To'}} } keys %$msgs ) {
        push @cur_report,
            $msgs->{$_}->{$header_table{'Date'}}
            . "\t\t"
            . convert_bytes( $msgs->{$_}->{$header_table{'Size'}} )
            . "\t\t"
            . $msgs->{$_}->{$header_table{'To'}}
            . "\n"
            ;

    }


    push @cur_report, "\n\n\nTime to fetch: $elapsed seconds\n";
    push @cur_report, "Iterated " . scalar( keys %$msgs ) . " messages.\n\n";

    return @cur_report;

} # }}}

# {{{ folder_message_count_report
#
# Displays a count of the number of messages in each folder.
# While it can use cached values, the message counts
# themselves are not a cached value.
#
sub folder_message_count_report {

    my $report_type = $reports->{folder_message_count_report};

    my $total_message_count = 0;

    my @fsize_report = ();

    push @fsize_report, "\n\n$report_type\n\n";
    push @fsize_report, "Count\t\t\tFolder\n";
    push @fsize_report, '-' x 60 . "\n";

    my $raw_report = {};
    my $total_num_messages = 0;
    my $folders_counted = 0;

    my $num_folders = scalar(@$imap_folders);

    my $bar = IMAP::Progress->new( max => $num_folders, length => 10 );

    for ( @$imap_folders ) {

        # Quick sanity check to see if the chosen folder really exists, not all
        # that necessary because the folder list is well validated at the
        # beginning of the script, but, just to be safe in case anything was
        # modified during the runtime of this script...
        #
        if ( ! check_cache( 'validated_folder_list', $_ ) ) {

            if ( ! $imap->exists($_) ) {
                show_error( "Error: $_ not a valid folder: $@\nLastError: " . $imap->LastError );
                next;
            }

        }

        # Try to get the count of messages from cache,
        # otherwise get the count manually.
        #
        my $count = check_cache( 'message_count', $_ );

        if ( ! $count ) {

            #show_error( "COUNT is now\n" . Dumper( $count ) );

            $imap->examine($_)
                or show_error( "Error selecting $_: $@\n" );

            print "Checking folder: $_\n" if $opts->{verbose};

            $count = $imap->message_count();

            #show_error( "COUNT is now\n" . Dumper( $count ) );

        }

        $raw_report->{$_} = $count;

        # We don't want the 'All Mail' gmail label to skew
        # our total since it represents ALL messages in a
        # gmailbox.
        #
        $total_num_messages += $count unless $_ eq '[Gmail]/All Mail';

        $bar->info( "Sent: " . $imap->Transaction . " STATUS \"$_\"" );
        $bar->update( ++$folders_counted );

        $bar->write;

    }

    # Iterate the hashref of raw data and make it prettier.
    #
    for my $fc ( reverse sort { $raw_report->{$a} <=> $raw_report->{$b} } keys %$raw_report ) {
        push @fsize_report, $raw_report->{$fc} . "\t\t\t$fc\n";
    }

    push @fsize_report, '-' x 60 . "\n\n";
    push @fsize_report, "Total messages found: $total_num_messages\n\n";

    return @fsize_report;

} # }}}

# {{{ folder_message_sizes_report
#
sub folder_message_sizes_report {

    my $report_type = $reports->{folder_message_sizes_report};

    # Extra press-enter-to-continue prompts to make damn
    # sure that we understand that this is long, heavy
    # operation on a mailbox with a ton of messages.
    #
    show_error(
          "Caution: This operation will iterate EVERY SINGLE message\n"
        . "in your IMAP mailbox.  While only the message size\n"
        . "attribute is collected (full email messages are not\n"
        . "downloaded), this will still take a VERY long time!\n"
        . "Cancel out of this script if you do not wish to procede.\n"
        . "Hint, try using the --filter option to pare down the list\n"
        . "of folders before running this report.\n"
    );

    show_error(
        "\n\n"
        . join( "\n", @$imap_folders )
        . "\n\nAbove is the list of folders that will be iterated for the report.\n\n"
        . "Last chance to cancel!\n\n"
    );

    my $total_message_count = 0;

    my @msize_report = ();

    push @msize_report, "\n\n$report_type\n\n";
    push @msize_report, "Size\t\t\tCount\t\t\tFolder\n";
    push @msize_report, '-' x 60 . "\n";

    my $raw_report         = {};
    my $total_num_messages = 0;
    my $current_iter       = 0;

    for ( @$imap_folders ) {

        # Always skip the 'All Mail' gmail label.  This
        # represents every message, so it will just skew the
        # results.
        #
        if ( $_ eq '[Gmail]/All Mail' ) {
            print "...Skipping the $_ Folder...\n";
            next;
        }

        my ( $cur_size, $message_count ) = get_folder_size($_);

        if ($cur_size) {
            $raw_report->{$_}->{size}  = $cur_size;
            $raw_report->{$_}->{count} = $message_count;
            $total_num_messages += $message_count;
        }

    }

    for my $cur_folder ( reverse sort { $raw_report->{$a}->{size} <=> $raw_report->{$b}->{size} } keys %$raw_report ) {

        push @msize_report,
              convert_bytes( $raw_report->{$cur_folder}->{size} )
            . "\t" x 3
            . $raw_report->{$cur_folder}->{count}
            . "\t" x 3
            . $cur_folder
            . "\n";

    }

    push @msize_report, '-' x 60 . "\n\n";
    push @msize_report, "Total messages found: $total_num_messages\n\n";

    return @msize_report;

} # }}}

# {{{ choose_action
#
# Expects to receive a string to use as the banner to
# display at the top of the menu screen.
#
# Returns the choice corresponding with the menu item
# selected from Term::Menus, which is the verbatim
# description of that particular menu option.
#
sub choose_action {

    my $args = shift;

    my $banner       = $args->{banner};
    my $report_types = $args->{reports};

    my @report_menu_items = sort map { $report_types->{$_} } keys %$report_types;

    my ( undef, $term_height, undef, undef ) = GetTerminalSize();
    my $display_this_many_items = $term_height - 16;

    my @menu_options;

    $menu_options[0]  = \@report_menu_items;
    $menu_options[1]  = $banner;
    $menu_options[2]  = $display_this_many_items;
    $menu_options[10] = 'One';

    my $choice = &pick(@menu_options);

    #ddump( 'choice', $choice ) if $opts->{debug};

    return $choice;


} # }}}

# {{{ folder_choice
#
# Shunted in this function to allow for report types that
# don't include selecting a single folder.
#
# Expects to receive a hashref to a list of folders and the
# name of the current action for display at the top of the
# menu.
#
# Returns the name of the folder after validating it and
# setting it to the selected state.
#
sub folder_choice {

    my $args = shift;

    my $folders = $args->{folders};
    my $type    = $args->{report_type};

    # Here's where we call to our picker function which uses
    # Term::Menus to allow us to pick from the list of
    # folders.
    #
    my $choice = show_folder_picker({
                    imap_folders => $folders,
                    banner => "Choose folder for report type: $type" });

    next unless $choice;

    next if $choice eq ']quit[';

    # Quick sanity check to see if the chosen folder really exists, not all that
    # necessary because the use of Term::Menus pretty much guarantees that we'll
    # get a valid folder name, but this should at least catch it if there is a
    # foldername with odd characters.
    #
    if ( ! $imap->exists($choice) ) {
        show_error( "Error: $choice not a valid folder.\n\nRaw IMAP error output: \n" . $imap->LastError );
        return;
    }


    # The examine method is a read-only version of select.
    # The act of iterating a folder on an imap server can
    # affect the status of messages and we want this to be a
    # completely transparent function.
    #
    $imap->examine($choice)
        or show_error( "Error: Error selecting folder: ${choice}...\n" . $imap->LastError );

    print "\n\n Folder selected: $choice\n";

    return $choice;

} # }}}

# {{{ show_folder_picker
#
# Expects to receive two params, the banner to display at
# the top of the menu and a arrayref of the list of imap
# folders from which to choose.
#
# Returns the full foldername.
#
sub show_folder_picker {

    my $args = shift;

    my $banner      = $args->{banner};
    my $folder_list = $args->{imap_folders};

    # Don't bother asking if we only have one folder from
    # which to choose.
    #
    if ( scalar(@$folder_list) == 1 ) {
        return $folder_list->[0];
    }

    #ddump( 'folders', $folder_list ) if $opts->{debug};

    # Dynamically size our picker menu to the size of the
    # terminal window.
    #
    my ( undef, $term_height, undef, undef ) = GetTerminalSize();
    my $display_this_many_items = $term_height - 14;
    my @menu_options;

    $menu_options[0]  = $folder_list;
    $menu_options[1]  = $banner;
    $menu_options[2]  = $display_this_many_items;
    $menu_options[10] = 'One';

    # Make the call to Term::Menus
    #
    my $choice = &pick(@menu_options);

    return $choice if $choice eq ']quit[';

    #ddump( 'choice', $choice ) if $opts->{debug};

    return $choice;


} # }}}

# {{{ fetch_folders
#
# Expects to receive an anon hashref of arguments containing
# lists for filters and excludes.
#
# Return an arrayref of folders after filtering and
# validating.
#
sub fetch_folders {

    my $args = shift;

    my $filter_list   =
        $args->{filters}
        ? $args->{filters}
        : []
        ;

    my $excludes_list =
        $args->{excludes}
        ? $args->{excludes}
        : []
        ;

    my $list         = [];
    my $menu_folders = [];

    $list = check_cache('folder_list');

    if ( ! defined $list ) {

        $list = $imap->folders
            or die_clean( 1, "Error fetching folders: $!\n" . $imap->LastError );

        put_cache({ content_type => 'folder_list', values => $list });

    }

    print "Processing list of IMAP folders...\n";

    # Here's where we filter out what we want from the list
    # of folders.
    #
    my @filtered_list =
                           grep { my $item = $_; not grep { $item =~ m/$_/i } @$excludes_list }
           @$filter_list ? grep { my $item = $_;     grep { $item =~ m/$_/i } @$filter_list   } @$list : @$list;

    # Run the exists method on the folder to compensate for
    # the odd behavior that can come from using nested gmail
    # labels.
    #
    my $max = scalar(@filtered_list);

    die_clean( 1, "No folders..." ) unless $max;

    my $bar = IMAP::Progress->new( max => $max, length => 10 );
    my $counter= 1;

    for my $ff ( @filtered_list ) {
        if ( $opts->{force} || check_cache( 'validated_folder_list', $ff ) ) {

            # The force option is allowed to short circuit
            # the imap->exists method...
            #
            push @$menu_folders, $ff;

        } else {

            if ( $imap->exists($ff) ) {
                push @$menu_folders, $ff;
                put_cache({ content_type =>  'validated_folder_list', folder => $ff });
            }

        }

        $bar->info( "Sent: " . $imap->Transaction . " STATUS \"$ff\"" );
        $bar->update( $counter++ );
        $bar->write;

    }

    return $menu_folders;

} # }}}

# {{{ fetch_msgs
#
# Expects to receive a list of items representing the
# message attributes we want to fetch.
#
# Returns a hashref of message id's as the keys, and the
# values for each key are hashrefs of the message attributes
# on which we want to report.
#
sub fetch_msgs {

    my $folder = shift;

    my @headers;

    push @headers, $header_table{$_} for qw/Date Subject Size To From/;

    #ddump( 'headers',        \@headers )        if $opts->{debug};

    # This will hold the entire result of fetching
    # operations.
    #
    my $fetcher = {};

    $fetcher = check_cache( 'fetched_messages', $folder );

    return $fetcher if $fetcher;

    # Quick sanity check to see if the chosen folder really exists, not all that
    # necessary because the folder list is well validated at the beginning of
    # the script, but, just to be safe in case anything was modified during the
    # runtime of this script...
    #
    if ( ! $imap->exists($folder) ) {
        show_error( "Error: $folder not a valid folder: $@\nLastError: " . $imap->LastError );
        return ( 0, 0 );
    }

    $imap->examine($folder)
        or show_error( "Error selecting $folder: $@\n" );


    # One more quick sanity check to make sure we really are
    # in a folder and that folder is in a select (examine)
    # state.
    #
    if ( $imap->Folder() ) {

        my $msg_count = $imap->message_count;

        # Return empty handed if there are no messages in
        # the folder.
        #
        return unless $msg_count;

        # Take the list of message id's and break them up
        # into smaller chunks in the form of an array of
        # MessageSet objects.
        #
        my @sequences = $imap->messages;

        # Trying to come up with a way of trapping a Ctrl-C
        # to gracefully finish the current iteration and
        # finish producing the report.  It doesn't work very
        # well.
        #
        $SIG{'INT'} = 'break_fetch';

        print "Fetching $msg_count messages...\n";

        $fetcher = $imap->fetch_hash( \@sequences, @headers );


        # Set the break function back to what it was.
        #
        $SIG{'INT'} = 'die_signal';

    } else {
        die_clean( 1, "Error checking current folder selection: $! " . $imap->LastError );
    }

    #ddump( 'fetcher', $fetcher ) if $opts->{debug};

    my $max = ( scalar(keys %$fetcher) );
    my $sbar = IMAP::Progress->new( max => $max, length => 10 );

    print "\n\n";

    $sbar->text('Processing headers:');

    my $scounter = 0;

    # Ugly.
    #
    # Iterate the whole list of fetched messages and fix
    # each value returned.
    #
    for my $cur_id ( keys %$fetcher ) {

        for my $cur_header (@headers) {

            $fetcher->{$cur_id}->{$cur_header} =
                stripper( $header_table{$cur_header},
                          $fetcher->{$cur_id}->{$cur_header} );

            $fetcher->{$cur_id}->{$cur_header} = '[EmptyValue]' unless $fetcher->{$cur_id}->{$cur_header};

        }

        $sbar->update( ++$scounter );
        $sbar->write;

    }

    put_cache({ content_type => 'fetched_messages', folder => $folder, values => $fetcher });

    return $fetcher;

} # }}}

# {{{ threaded_fetch_msgs
#
# Expects to receive a list of items representing the
# message attributes we want to fetch.
#
# Returns a hashref of message id's as the keys, and the
# values for each key are hashrefs of the message attributes
# on which we want to report.
#
sub threaded_fetch_msgs {

    my $folder = shift;

    verbose( "THREADED FOLDER: " . Dumper( $folder ) );

    my @headers;

    # TODO
    #
    # Fix this header handling...
    #
    push @headers, $header_table{$_} for qw/Date Subject Size To From/;

    #ddump( 'headers', \@headers ) if $opts->{debug};

    # This will hold the entire result of fetching
    # operations.
    #
    my $fetcher = {};

    $fetcher = check_cache( 'fetched_messages', $folder );

    verbose( "FETCHER NOW: " . Dumper( $fetcher ) );

    return $fetcher if $fetcher;

    # Quick sanity check to see if the chosen folder really exists, not all that
    # necessary because the folder list is well validated at the beginning of
    # the script, but, just to be safe in case anything was modified during the
    # runtime of this script...
    #
    if ( ! $imap->exists($folder) ) {
        show_error( "Error: $folder not a valid folder: $@\nLastError: " . $imap->LastError );
        return ( 0, 0 );
    }

    $imap->examine($folder)
        or show_error( "Error selecting $folder: $@\n" );


    # One more quick sanity check to make sure we really are
    # in a folder and that folder is in a select (examine)
    # state.
    #
    if ( $imap->Folder() ) {

        # This object will hold our message ids.
        #
        my $msgset = Mail::IMAPClient::MessageSet->new( $imap->messages );

        my $msg_ids = [];

        if ( $msgset ) {
            $msg_ids = $msgset->unfold;
        }

        my $msg_count = scalar(@$msg_ids);

        # Return empty handed if there are no messages in
        # the folder.
        #
        return unless $msg_count;

        # Take the list of message id's and break them up
        # into smaller chunks in the form of an array of
        # MessageSet objects.
        #
        my $threaded_sequences = threaded_sequence_chunker( $msg_ids );

        #ddump( 'sequences', \@sequences ) if $opts->{debug};

        #my $num_blocks;

       #for my $cur_bucket ( keys %$threaded_sequences ) {
       #    for my $cur_msg_set_list ( @{$threaded_sequences->{$cur_bucket}} ) {
       #        $num_blocks += scalar @$cur_msg_set_list;
       #    }
       #}

        #my $prog_bar = IMAP::Progress->new( max => $num_blocks, length => 10 );

        #$prog_bar->text( "$msg_count remaining" );
        #$prog_bar->update( 0 );
        #$prog_bar->write;

        #my $iter = 1;

        #my $eta = [];

        # Trying to come up with a way of trapping a Ctrl-C
        # to gracefully finish the current iteration and
        # finish producing the report.  It doesn't work very
        # well.
        #
        #$SIG{'INT'} = 'break_fetch';

        my $threads = [];

        # Each thread will stuff stats in the global progress_queue object.
        # Spawn this thread which will keep watching the queue and computing
        # stats.
        #
        {

      #_#   my $progress_bar = IMAP::Progress->new( max => $msg_count, length => 10 );

      #_#   my $progress_thread = threads->create( \&threaded_progress_bar, $progress_bar )->detach();

            print "\nInitiating fetcher threads: ";

            for my $cur_bucket ( sort keys %$threaded_sequences ) {

                print "$cur_bucket ";

                #$threads->[$cur_bucket] = threads->create(

                {

                    # Detaching thread because we will read our results from the
                    # fetched messages queue so we don't need to have the thread
                    # return anything.
                    #
                    #)->detach();

                    $threads->[$cur_bucket] = threads->create(

                        { 'context' => 'void',
                          'exit'    => 'thread_only' },

                        \&imap_thread,

                            $folder,
                            $cur_bucket,
                            $threaded_sequences->{$cur_bucket}

                    );

                }

            }

            print "\n";


            # Before proceding, wait for each thread to finish
            # collecting messages by checking the fetcher status
            # queue.
            #
            while ( 1 ) {

                print "\nWaiting for threads to complete...\n" if $opts->{verbose};

                sleep 1;

                my $statuses;

                {
                    lock($fetcher_status);
                    $statuses = $fetcher_status->pending();
                }

                #$Data::Dumper::Varname = 'statuses';

                #print "\n\n" . Dumper( $statuses ) . "\n\n";

                last if $statuses == $opts->{threads};

            }

            # Rejoin the threads for a clean exit.
            #
            for my $cur_bucket ( sort keys %$threaded_sequences ) {
                $threads->[$cur_bucket]->join();
            }

            print "\nThreads complete...\n" if $opts->{verbose};

        }

#_#     print "Progress bar thread ended...\n" if $opts->{verbose};


        # Now we that the threads have completed, iterate
        # the values in our fetched messages queue and merge
        # them into the big fetcher hashref.
        #
        {

            print "Parsing fetched messages...\n";

            lock($fetched_queue);

            my $pending = $fetched_queue->pending();

            print "Total messages pending for processing: $pending\n";

           #my $bar = IMAP::Progress->new( max => $pending, length => 10 );

           #$bar->text('Parsing fetched messages:');

           #my $parse_counter = 0;

            my %from_fetched_queue = $fetched_queue->extract( 0, $pending );

            #verbose( 'from_fetched_queue_hash: ' . Dumper( \%from_fetched_queue ) );

            @{$fetcher}{ keys %from_fetched_queue } = values %from_fetched_queue;

            print "\nParse complete...\n";


            #$Data::Dumper::Varname = 'fetcher_';
            #print "\n" . Dumper( $fetcher ) . "\n\n";

        }

        # Set the break function back to what it was.
        #
        #$SIG{'INT'} = 'die_signal';

    } else {
        die_clean( 1, "Error checking current folder selection: $! " . $imap->LastError );
    }

    #ddump( 'fetcher', $fetcher ) if $opts->{debug};

    #show_error( "FETCHER_FINAL: " . Dumper( $fetcher ) );

    my $max = ( scalar(keys %$fetcher) );

    my $sbar = IMAP::Progress->new( max => $max, length => 10 );

    print "\n\n";

    $sbar->text('Processing headers:');

    my $scounter = 0;

    # Ugly.
    #
    # Iterate the whole list of fetched messages and fix
    # each value returned.  The header information has some
    # issues like CRLF and such.
    #
    print "Fixing message headers...\n";
    for my $cur_id ( keys %$fetcher ) {

        for my $cur_header (@headers) {

            $fetcher->{$cur_id}->{$cur_header} =
                stripper( $header_table{$cur_header},
                          $fetcher->{$cur_id}->{$cur_header} );


            $fetcher->{$cur_id}->{$cur_header} = '[EmptyValue]'
                unless $fetcher->{$cur_id}->{$cur_header};

        }

        # Keep the counter from updating too frequently...
        #
        if ( ( ( $scounter % 20 ) + 1 )  == 20 ) {
            $sbar->update( $scounter++ );
            $sbar->write;
        }

    }

    #show_error( "FETCHER: " . Dumper( $fetcher ) );
    #show_error( "FETCHER REF: " . Dumper( ref $fetcher ) );

    # Store our results in the cache then return the
    # results.
    #
    print "Storing messages in cache...\n";
    put_cache({ content_type => 'fetched_messages', folder => $folder, values => $fetcher });

    return $fetcher;

} # }}}

# {{{ imap_thread
#
sub imap_thread {

    # Each thread gets an appropriate list of M::I::MessageSet objects
    #
    my $folder       = shift;
    my $cur_bucket   = shift;
    my $msg_set_list = shift;

    my @headers;

    # TODO
    #
    # Fix this header handling...
    #
    push @headers, $header_table{$_} for qw/Date Subject Size To From/;




    # Each thread gets its own imap object...
    #
    my $imap_thread  = Mail::IMAPClient->new(
                Server           => $opts->{server},
                Port             => $opts->{port},
                User             => $opts->{user},
                Password         => $opts->{password},
                Keepalive        => $opts->{Keepalive},
                Fast_io          => $opts->{Fast_io},
                Ssl              => $opts->{Ssl},
                Maxcommandlength => $opts->{Maxcommandlength},
                Reconnectretry   => $opts->{Reconnectretry},
                Uid              => 0,
                Clear            => 100,
                #Debug            => $opts->{debug} . '.' . $cur_bucket,
                #Debug_fh         => *DBG,

    ) or die "Cannot connect to host : $@";

    #show_error( "CUR BUCKET: " . Dumper( $cur_bucket ) );

    # Reselect the folder so this thread is in the right place.  No validation
    # or exists check necessary at this stage.
    #
    if ( ! $imap_thread->examine($folder) ) {
        warn "Error selecting folder $folder in thread $cur_bucket: $@\n";
        last;
    }

    my $msg_set_object = shift( @$msg_set_list );

    my @cur_msg_id_list = $msg_set_object->unfold;

    my $cur_fetcher = $imap_thread->fetch_hash( \@cur_msg_id_list, @headers);

    # This thread is done, tear down the imap
    # connection.
    #
    $imap_thread->disconnect;

    {
        lock( $fetched_queue );

        for my $m_id ( keys %$cur_fetcher ) {
            next unless $m_id;
            $fetched_queue->enqueue( $m_id => $cur_fetcher->{$m_id} );
        }

    }


    # Stick a message on the status queue that we're done.
    #
    {
        lock($fetcher_status);
        $fetcher_status->enqueue( 'complete' );
    }

    return;

} # }}}

# {{{ get_folder_size
#
# TODO
#
# Simplify this mess.
#
sub get_folder_size {

    my $folder = shift;

    print "\n\nFetching message details for folder '$folder'...\n\n";

    my $msgs =
        $use_threaded_mode
        ? threaded_fetch_msgs( $folder )
        : fetch_msgs( $folder )
        ;

    return ( 0, 0 ) unless scalar( keys %$msgs );

    my $totalsize;

    my $counter = 0;

    for ( keys %$msgs ) {
        $totalsize += $msgs->{$_}->{$header_table{'Size'}};
        $counter++;
    }

    #ddump( 'msgs', $msgs ) if $opts->{debug};

    if ( ! $counter ) {
        show_error( "Error: No messages to report for '$folder'..." );
        return ( 0, 0 );
    }

    #ddump( 'totalsize', $totalsize ) if $opts->{debug};

    return ( $totalsize, $counter );

} # }}}

# {{{  init_cache
#
# Crude implementation of a cache using Storable.
#
sub init_cache {

    my $cfile = shift;

    my $c;

    # If 'cache' file exists, load it, otherwise,
    # instantiate a cache hashref.
    #
    if ( -r $cfile ) {
        print "Loading cache...\n";
        $c = retrieve($cfile);
    } else {

        # Not particularly necessary, mainly for
        # documentation purposes.
        #
        # All cache elements are stored as hashrefs because
        # it's a lazy way of ensuring that I'm not stuffing
        # duplicates into the cache.
        #
        $c->{ $opts->{server} }->{imap_folders}->{fetched_messages} = {};
        $c->{ $opts->{server} }->{imap_folders}->{folders}          = {};
        $c->{ $opts->{server} }->{imap_folders}                     = {};
        $c->{ $opts->{server} }->{imap_folders}->{validated}        = {};

    }

    $c->{updated} = 0;

    return $c;


} # }}}

# {{{ check_cache
#
# My crude method of a caching mechanism.
#
# Feed this function the type of content and value for which
# to look.  Returns cached elements based on the different
# types of information, arrayrefs, hashrefs, bools, etc.
#
# TODO
#
# Implement cache aging
#
sub check_cache {

    my $content_type = shift;
    my $value = shift;

    my $cur_time = time;

    return unless $opts->{cache};

    if ( $content_type eq 'folder_list' ) {

        # Checks the cached list of folders and returns an
        # arrayref list of them.

        if ( defined $cache->{$opts->{server}}->{imap_folders}->{folders} ) {

            # First attempt at cache aging.  Screwed
            # something else up which broke this.  Need to
            # revisit and fix...

           #my $last_update =
           #    defined $cache->{imap_folders}->{age} && $cache->{imap_folders}->{age}
           #    ? $cache->{imap_folders}->{age}
           #    : 0
           #    ;

           #my $age =
           #    defined $last_update && $last_update
           #    ? $cur_time - $last_update
           #    : 0
           #    ;

           ## Return emptyhanded if age has no value or if
           ## it has a value and it is greater than the
           ## maximum cache age.
           ##
           #show_error( "LAST UPDATE: " . Dumper( $last_update ) );
           #show_error( "AGE: " . Dumper( $age ) );

           #if ( defined $age && $age && $age > 0 && $age > $opts->{cache_age} ) {
           #    return;
           #}

            show_error('Fetched folder list from cache.')
                if $opts->{verbose};

            #ddump( 'Folders from cache', $cache->{$opts->{server}}->{imap_folders}->{folders} )
                #if $opts->{debug};

            my @f = sort keys %{ $cache->{$opts->{server}}->{imap_folders}->{folders} };
            return \@f;

        }

    } elsif ( $content_type eq 'validated_folder_list' ) {

        # The list of VALIDATED folders is cached
        # separately.  This just returns a true/false if a
        # folder appears in the list of validated folders.
        # (Validated folders have passed a test using the
        # 'exists' method.)
        #
        return unless defined $value && $value;

        if ( defined $cache->{$opts->{server}}->{imap_folders}->{validated}
             && ref $cache->{$opts->{server}}->{imap_folders}->{validated} eq 'HASH' ) {

            my $result = grep $value eq $_, keys %{ $cache->{$opts->{server}}->{imap_folders}->{validated} };

            return $result;

        }

    } elsif ( $content_type eq 'fetched_messages' ) {

        return unless defined $value && $value;

        if ( defined $cache->{$opts->{server}}->{imap_folders}->{fetched_messages}->{$value}->{messages}
             && ref $cache->{$opts->{server}}->{imap_folders}->{fetched_messages}->{$value}->{messages} eq
             'HASH' ) {

            return $cache->{$opts->{server}}->{imap_folders}->{fetched_messages}->{$value}->{messages};

        }

    } elsif ( $content_type eq 'message_count' ) {

        # This looks into the cache of messages and if
        # messages have been cached, returns the count of
        # the number of messages stored there.  Folder
        # message counts themselves are not actually cached.
        #

        return unless defined $value && $value;

        if ( defined $cache->{$opts->{server}}->{imap_folders}->{fetched_messages}->{$value}->{messages}
             && ref $cache->{$opts->{server}}->{imap_folders}->{fetched_messages}->{$value}->{messages} eq 'HASH' ) {

            my $result = scalar( keys %{ $cache->{$opts->{server}}->{imap_folders}->{fetched_messages}->{$value}->{messages} } );
            #show_error( "VALID FOLDER RESULT: " . Dumper( $result ) );

            return $result;

        }

    }


    return;

} # }}}

# {{{ put_cache
#
# Handle inserting the various types of information we want
# to cache.
#
# Sticks in the current time value so for cache aging
# purposes later.
#
sub put_cache {

    my $args = shift;

    my $content_type = $args->{content_type};
    my $values       = $args->{values};
    my $folder       = $args->{folder};

    return unless $content_type;

    if ( $content_type eq 'folder_list' ) {

        if ( ref $values eq 'ARRAY' ) {
            %{$cache->{$opts->{server}}->{imap_folders}->{folders}} = map { $_ => 1 } @$values;
            $cache->{$opts->{server}}->{imap_folders}->{age} = time;
            $cache->{updated} = 1;
        } else {
            return;
        }

    } elsif ( $content_type eq 'validated_folder_list' ) {

        return unless $folder;
        return if ref $folder;

        # Not sticking the age here yet because I'm not
        # quite yet sure how I'm going to handle the aging
        # of the cache of validated folder list.
        #
        $cache->{$opts->{server}}->{imap_folders}->{validated}->{$folder} = 1;
        $cache->{updated} = 1;

    } elsif ( $content_type eq 'fetched_messages' ) {

        return unless ref $values eq 'HASH';
        return unless $folder;

        $cache->{$opts->{server}}->{imap_folders}->{fetched_messages}->{$folder}->{messages} = $values;
        $cache->{$opts->{server}}->{imap_folders}->{fetched_messages}->{$folder}->{age}      = time;
        $cache->{updated} = 1;

    }


    return;

} # }}}

# {{{ threaded_progress_bar
#
# Display our multithreaded progress bar.
#
sub threaded_progress_bar {

    my $pbar = shift;

    my $stats = [];

    print "\n\nJob progress...\n";
    my $scounter = 0;

    while (1) {

        sleep 1;

        my $item;

        {
            lock($progress_queue);
            $item = $progress_queue->extract();

        }

        next unless $item;

        if ( $item eq 'quit' ) {
            last;
        }

        if ( ref $item eq 'ARRAY' ) {

            # This just continuously pushes stats into
            # an arrayref.

            my $elapsed             = $item->[0];
            my $fetched_count       = $item->[1];
            my $total_message_count = $item->[2];

            next unless ( $elapsed && $fetched_count && $total_message_count );

            push @$stats, [ $elapsed, $fetched_count, $total_message_count ];

        }

        # The statistics are fully recomputed from scratch
        # with every iteration.
        #

        my $total_elapsed             = 0;
        my $total_fetched_count       = 0;
        my $remaining_messages        = 0;
        my $total_message_count       = 0;
        my $elapsed_seconds_completed = 0;

        for ( @$stats ) {
            $elapsed_seconds_completed += $_->[0];
            $total_fetched_count       += $_->[1];
            $total_message_count        = $_->[2]; #yuck
        }

        next unless $elapsed_seconds_completed;
        next unless $total_fetched_count;

        $remaining_messages = $total_message_count - $total_fetched_count;

        last unless $remaining_messages;

        my $rate          = $elapsed_seconds_completed / $total_fetched_count;
        my $total_eta     = $rate * $total_message_count;
        my $remaining_eta = $total_eta - $elapsed_seconds_completed;

        my $text = "[$opts->{threads} threads] "
            . sprintf( "%8d remaining", $remaining_messages );

        my $info = 'messages';
        $info .= $remaining_eta ? ' - ETA: ' . convert_seconds($remaining_eta) : '';

        $pbar->info( $info );
        $pbar->text( $text );

        # Keep the counter from updating too frequently...
       #if ( ( ( $scounter++ % 100 ) + 1 )  == 100 ) {
            $pbar->update( $total_fetched_count );
            $pbar->write;
       #}

    }

    {
        lock($progress_queue);
        $progress_queue->enqueue('ended');
    }


} # }}}

# {{{ show_current_fetch_stats
#
# Pretty print the stats of the current fetch operation.
# Will replace with a progressbar.
#
#sub show_current_fetch_stats {

#my $cur_msgs         = shift;
#my $current_block    = shift;
#my $total_num_blocks = shift;
#my $elapsed          = shift;
#my $remaining        = shift;

#format STDOUT =
#@<<<<<<< @>>>>> @<<<<<<<<< @>>>>>>>>>>> @>>>> @< @<<<<  @<<<<<<<<<<<<<<<<<<<<<<< @*
#'Iterated', $cur_msgs, 'messages, ', 'block number', $current_block, 'of', $total_num_blocks, "(fetch time: $elapsed seconds)", $remaining ? "ETA: $remaining" : ''
#.
#
#write;
#

#} # }}}

# {{{ estimate_completion_time
#
# Compute the ETA.
#
# TODO
#
# fix this mess.
#
sub estimate_completion_time {

    my $stats                  = shift;
    my $total_number_of_blocks = shift;

    my $number_of_blocks_completed = scalar(@$stats);

    my $elapsed_seconds_completed = 0;

    for (@$stats) {
        $elapsed_seconds_completed += $_->[0];
    }

    my $rate          = $elapsed_seconds_completed / $number_of_blocks_completed;
    my $total_eta     = $rate * $total_number_of_blocks;
    my $remaining_eta = $total_eta - $elapsed_seconds_completed;

    return convert_seconds($remaining_eta);

} # }}}

# {{{ sequence_chunker
#
# This is where we take our big list of message id's and
# split it into chunks.  Expects to receive two params.
# First, a Mail::IMAPClient::MessageSet object containing
# our message id's.  Second, the maximum size of each chunk.
#
# Returns an array of M::I::MessageSet objects.
#
# The reason for the back-and-forth switching from lists to
# MessageSet objects is because M::I::MessageSet makes the
# effort to express the list of message ids in optimal
# RFC2060 representation.
#
sub sequence_chunker {

    my $all_msg_ids = shift;
    my $max         = shift;

    show_error( 'Error processing sequences' ) unless $all_msg_ids;

    my @cur_block;
    my @seq_objects;

    while ( my @cur_block = splice @$all_msg_ids, 0, $max ) {

            my $cur_set = Mail::IMAPClient::MessageSet->new( @cur_block );
            push @seq_objects, $cur_set;

    }

    #ddump( 'seq_objects', \@seq_objects ) if $opts->{debug};

    verbose( "Sequences: " . Dumper( \@seq_objects ) );

    return @seq_objects;

} # }}}

# {{{ threaded_sequence_chunker
#
# This is where we take our big list of message id's and
# split it into chunks.  Expects to receive two params.
# First, a Mail::IMAPClient::MessageSet object containing
# our message id's.  Second, the maximum size of each chunk.
#
# Returns an array of well-distributed M::I::MessageSet
# objects, carefully sorted into different buckets so that
# no threads will be attempting to fetch the same messages.
#
# The reason for the back-and-forth switching from lists to
# MessageSet objects is because M::I::MessageSet makes the
# effort to express the list of message ids in optimal
# RFC2060 representation.
#
sub threaded_sequence_chunker {

    my $all_msg_ids = shift;

    show_error( 'Error processing sequences' ) unless $all_msg_ids;

    my $message_count = scalar(@$all_msg_ids);

    my $max = int( $message_count / $opts->{threads} );

    my @cur_block;

    my $msg_id_buckets  = {};
    my $msg_set_buckets = {};

    my $counter = 0;

    while ( my @cur_block = splice @$all_msg_ids, 0, $max ) {

        my $bucket_id = $counter % $opts->{threads};
        my $cur_set = Mail::IMAPClient::MessageSet->new( @cur_block );
        push @{$msg_set_buckets->{$bucket_id}}, $cur_set;
        $counter++;

    }

    #show_error( "MSG SET BUCKETS: " . Dumper( $msg_set_buckets ) );

    #die_clean( 0, "Quitting..." );

    #return $seq_objects;

    verbose( "Sequences: " . Dumper( $msg_set_buckets ) );

    return $msg_set_buckets;

} # }}}

# {{{ stripper
#
# Oddly, the chomp function behaves in an unexpected way on
# subjects and other headers returned from the imap server.
# I know it's an issue of LF vs. CR, but I still couldn't
# get it to behave cleanly, so I did it this way rather than
# any chomp chop chomp monkey business...
#
sub stripper {

    my $name  = shift;
    my $field = shift;

    return unless $name;
    return unless $field;

    $field =~ s/\n+//;
    $field =~ s/\r+//;
    $field =~ s/\R+//;
    $field =~ s/^\s+//;
    $field =~ s/\s+$//;

    # Strip off the name of the envelope attribute
    #
    if ( $field =~ m/^$name:\s+(.*)$/ ) {
        $field = $1;
    }

    # For from addresses, just grab the address.
    #
    if ( $name eq 'From' ) {
        $field =~ m/[<](.*)[>]/;
        $field = $1;
    }

    # Strip off the subject cruft...
    #
    if ( $name eq 'Subject' ) {
        $field =~ s/^Re:\s+//gi;
        $field =~ s/^Fwd:\s+//gi;
        $field =~ s/\s+Re:\s+//gi;
        $field =~ s/\s+Fwd:\s+//gi;
    }


    if ( ! $field ) {
        $field = ']]EMPTY[[';
    }

    return $field;

} # }}}

# {{{ show_report_types
#
sub show_report_types {

my $types = report_types();

print "\n\nType                                 Description\n"
    . '-' x 75
    . "\n";

for ( sort { $types->{$a} cmp $types->{$b} } keys %$types ) {

# pretty print the types of reports...
#
format STDOUT = 
@<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  @*
$_, $types->{$_}
.
write;
}

die_clean( 0, '' );

} # }}}

# {{{ show_error
#
sub show_error {

    my $error = shift;

    print "\n\n\n$error\n";

    enter();

} # }}}

# {{{ break_fetch
#
# Yeah.  This doesn't work right.
#
sub break_fetch {

    print "Will break after the current fetch...\n";
    $break = 1;

    return;

} # }}}

# {{{ generate_subject_search_string
#
# To display a helpful string you can cut and paste directly
# into the gmail search pane to search for the reported
# message.
#
sub generate_subject_search_string {

    my ( $folder, $date, $subject ) = @_;


    my %folders = (
                    'INBOX'             => 'in:inbox',
                    '[Gmail]/Sent Mail' => 'is:sent',
                    'Sent'              => 'is:sent',
                    'Sent Items'        => 'is:sent',
                    '[Gmail]/Spam'      => 'is:spam',
                    '[Gmail]/Starred'   => 'is:starred',
                    '[Gmail]/All Mail'  => 'in:anywhere',
    );

    my $label;

    if ( defined $folders{$folder} ) {
        $label = $folders{$folder};
    } else {
        $label = 'label:' . '"' . $folder . '"';
    }

    my $dm = Date::Manip::Date->new();

    my ( $bdate, $adate );

    my $aresult = $dm->parse($date);

    if ( $aresult ) {
        $bdate = '';
        $adate = '';
    } else {

        my $parsed_adate = $dm->printf('%Y/%m/%d');
        my $aepoch       = $dm->printf('%s');
        my $bepoch       = $aepoch + ( 24 * 60 * 60 );
        my $bresult      = $dm->parse("epoch $bepoch");
        my $parsed_bdate = $dm->printf('%Y/%m/%d');

        $adate = "after:$parsed_adate";
        $bdate = "before:$parsed_bdate";

    }

    return
        join( ' ',
              "subject:\"$subject\"",
              $adate,
              $bdate,
              $label );

} # }}}

# {{{ generate_search_string
#
# To display a helpful string you can cut and paste directly
# into the gmail search pane to search for the reported
# message.
#
sub generate_search_string {

    my $args = shift;


    my $folder = $args->{folder};
    my $date   = $args->{date};
    my $header = $args->{header};
    my $value  = $args->{value};

    my %folders = (
                    'INBOX'             => 'in:inbox',
                    '[Gmail]/Sent Mail' => 'is:sent',
                    'Sent'              => 'is:sent',
                    'Sent Items'        => 'is:sent',
                    '[Gmail]/Spam'      => 'is:spam',
                    '[Gmail]/Starred'   => 'is:starred',
                    '[Gmail]/All Mail'  => 'in:anywhere',
    );

    my $label;

    if ( defined $folders{$folder} ) {
        $label = $folders{$folder};
    } else {
        $label = 'label:' . '"' . $folder . '"';
    }

    my $dm = Date::Manip::Date->new();

    my ( $bdate, $adate );

    my $aresult = $dm->parse($date);

    if ( $aresult ) {
        $bdate = '';
        $adate = '';
    } else {

        my $parsed_adate = $dm->printf('%Y/%m/%d');
        my $aepoch       = $dm->printf('%s');
        my $bepoch       = $aepoch + ( 24 * 60 * 60 );
        my $bresult      = $dm->parse("epoch $bepoch");
        my $parsed_bdate = $dm->printf('%Y/%m/%d');

        $adate = "after:$parsed_adate";
        $bdate = "before:$parsed_bdate";

    }

    return
        join( ' ',
              "$header:\"$value\"",
              $adate,
              $bdate,
              $label );

} # }}}

# {{{ read_config_file
#
# Allow a home directory config file containing your
# password and other options to be set.
#
# Simple foo=bar syntax.  Ignore comments, strip leading and
# trailing spaces.
#
sub read_config_file {

    my $cf = $opts->{conf};

    if ( $cf && ! -f $cf ) {
        return;
    }

    my $mode = ( stat($cf) )[2];

    # Abort unless the file is ONLY readable by the user.
    #
    if ( ( $mode & 00004 ) == 00004 || ( $mode & 00040 ) == 00040 ) {
        show_error( 'Configuration is readable by group or other. Restrict permissions to 600.  Skipping config file.' );
        return;
    }

    open my $conf_fh, '<', $cf
        or die "Error reading configuration file ($cf): $!\n\n";

    for (<$conf_fh>) {

        chomp;

        if ( $_ =~ m/[=]/ ) {

            s/#.*//;
            s/^\s+//;
            s/\s+$//;
            next unless length $_;

            my ( $key, $value ) = split( /\s*=\s*/, $_, 2 );

            print "Reading value from conf file: $key\n"
                if $opts->{verbose};

            $opts->{$key} = $value;

        }
    }

    close $conf_fh;

    return;

} # }}}

# {{{ sub convert_bytes
#
# For pretty printing byte numbers.
#
sub convert_bytes {

    my $bytes = shift;

    my $KB = $bytes / 1024;

    return sprintf( '%.1fKB', $KB )               if $KB < 1000;
    return sprintf( '%.1fMB', $KB / 1024 )        if $KB < 1000000;
    return sprintf( '%.1fGB', $KB / 1024 / 1024 ) if $KB < 100000000;

} # }}}

# {{{ sub convert_seconds
#
# For pretty printing the number of seconds.
#
sub convert_seconds {

    my $seconds = shift;

    my $days  = int( $seconds / ( 24 * 60 * 60 ) );
    my $hours = ( $seconds / ( 60 * 60 ) ) % 24;
    my $mins  = ( $seconds / 60 ) % 60;
    my $secs  = $seconds % 60;

    my $days_string  = $days  ? "$days days " : '';
    my $hours_string = $hours ? sprintf( '%02d', $hours ) . ' hours '   : '';
    my $mins_string  = $mins  ? sprintf( '%02d', $mins  ) . ' minutes ' : '';
    my $secs_string  = $secs  ? sprintf( '%02d', $secs  ) . ' seconds ' : '';

    return join( '', $days_string, $hours_string, $mins_string, $secs_string );

} # }}}

# {{{ sub die_signal
#
# Die cleanly on a signal.
#
sub die_signal {

    my @args = @_;

    if ( scalar @args > 0 ) {

        my $signal = shift(@args);
        die_clean( 1, "Died on signal: $signal" );

    }

} # }}}

# {{{ sub die_clean
#
# Die cleanly and display a message as well as write the
# cache.
#
sub die_clean {

    my $err = shift;
    my $msg = shift;

    #if ( defined $imap && ref $imap && $imap->IsConnected ) {
        #$imap->disconnect;
    #}

    close DBG;

    print "\n$msg\n";

   if ( $cache->{updated} && $err == 0 ) {
       print "\nWriting cache...\n";
       store( $cache, $opts->{cache_file} );
       my $cache_size = (stat( $opts->{cache_file} ))[7];
       print 'Cache size: ' . convert_bytes($cache_size) . "\n";
   }

    exit $err;

} # }}}

# {{{ debugging output
#
sub ddump {

    $Data::Dumper::Varname = shift;

    open( DDBG, '>>' . $opts->{log} . 'dumperlog' )
        or die_clean( 1, "Error opening dumperlog: $!\n" );

    print DDBG Dumper( @_ );

    close DDBG;

}

sub verbose {

    return unless $opts->{verbose};

    my $v = shift;

    print "\n$v\n";

    enter();

}

sub enter {

    print "\nPress [Enter] to continue: ";
    <>;

}

# }}}

# {{{ login/password reading...
#
sub disable_echo {
    print `/bin/stty -echo`;
}

sub enable_echo {
    print `/bin/stty echo`;
}

sub password_prompt {

    my $password = '';

    while ( ! $password ) {

        print 'Password : ';
        disable_echo();
        chomp( $password = <> );
        enable_echo();
        print "\n";

    }

    return $password;

}

sub user_prompt {

    my $user = '';

    while ( ! $user ) {
        print 'Email : ';
        chomp( my $user = <> );
        print "\n";
    }
    return $user;

}

# }}}

# {{{ print_report
#
# Display our report.  Put in its own function because I
# intend to complicate this later.
#
# TODO
#
# Better pager handling.
#
sub print_report {

    my $report = shift;

    my $file = "$ENV{HOME}/imap-report.txt";
    open ( RPT, ">" . $file )
        or die_clean( 1, "Unable to write report.\n" );

    print RPT $_ for @$report;
    close RPT;

    system( "less -niSRX $file" );

    die_clean( 0, "Quitting" )
        if $opts->{list};

} # }}}

# }}}

__END__

# {{{ POD

=pod

=head1 NAME

imap-report.pl - Generate reports on an imap account.

=head1 SCRIPT CATEGORIES

Mail

=head1 README

Primarily intended for use with a gmail account, this script can generate various reports on an imap mailbox.  I wrote this mostly out of frustration from google's lack of features to allow you to prune your mailbox.  Then it started to amuse me and I decided to try my first attempt at writing threaded perl.  It's not going well...

There is a crude caching mechanism present to speed things up after the message envelope information is loaded.  Even though only header information is fetched and cached, this is still a very heavy, time consuming, memory hungry operation on a huge mailbox.  The only operation that doesn't populate the cache automatically is the counting of all folders.  All other report types end up needing to actually iterate messages and therefor populates the cache.  Otherwise the count operation just uses the simple messages_count method of Mail::IMAPClient which uses the STATUS function of IMAP on an individual folder.

All message fetch operations are broken up into small (--maxfetch) chunks so that if there is a problem during the fetch, such as getting disconnected or other imap error, it won't abort the whole operation, just the current chunk of messages being fetched.

=head1 OSNAMES

any

=head1 PREREQUISITES

 Mail::IMAPClient >= 3.24
 Term::ReadKey
 Term::Menus
 Date::Manip

=head1 COREQUISITES

 IO::Socket::SSL - Needed by Mail::IMAPClient

=head1 SYNOPSIS

=head2 OPTIONS AND ARGUMENTS

=over 15

=item B<--user> I<username>

Optional username for IMAP authentication.  If omitted, you will be prompted after running the script.

=item B<--password> I<password>

Optional password for IMAP authentication.  If omitted, you will be prompted after running the script.

=item B<--server> I<server hostname or ip>

The identity of the IMAP server.

(default: imap.gmail.com)

=item B<--port> I<IMAP Port>

The port used to connect to the IMAP server.

(default: 993)

=item B<--top> I<integer number>

The number of messages in top ten style reports.

(default: 10)

=item B<--min> I<number of bytes>

Messages smaller than this size will be ignored when gathering statistics on individual messages.

(default: 100000)

=item B<--maxfetch> I<integer number>

The size of an individual fetch operation.  Breaks a fetch operation up into smaller chunks so that an individual folder with a massive number of messages will not kill the operation of there is a timeout or other communication problem.

=item B<--filters> I<string>

Folder filters.  Restrict all operations to folders matching the specified string.  This option can be specified multiple times.

=item B<--exclude> I<string>

Folder exclusions.  The list of folders will be pruned of the ones matching the specified string.  This option can be specified multiple times.  Perl compatible regex should work as long as you take care not to allow your shell to swallow up the expression.

=item B<--cache> I<cache_filename>

Name of the file used to store cached information.

(default: $HOME/.imap-report.cache)

=item B<--cache_age> I<integer>

Maximum age of cached information.

(currently non-functional)

(default: 1 day)

=item B<--conf> I<config_filename>

Name of the file in which to read configuration options.

All of these configuration options can be stored in this file using the same names listed here.  Must only be readable by the user.

(default: $HOME/.imapreportrc)

=item B<--list>

Just show the list of folders.

=item B<--Keepalive>

Corresponds to the Mail::IMAPClient Keepalive option.

(default: true)

=item B<--Fast_io>

Corresponds to the Mail::IMAPClient Fast_io option to allow buffered I/O.

(default: true)

=item B<--Reconnectretry>

Corresponds to the Mail::IMAPClient Reconnectretry option to and re-establish lost connections.

(default: 3)

=item B<--Maxcommandlength>

Corresponds to the Mail::IMAPClient Maxcommandlength option to limit the size of individual fetches.

(default: 1000)

=item B<--Ssl>

Corresponds to the Mail::IMAPClient Ssl option.

(default: true)

=item B<--debug>

Lots of ugly debugging output to a logfile...

=item B<--verbose>

A bit more output than usual

=back

=head2 EXAMPLE

C<./imap-report.pl>

(No options are necessary to run this script.  See the description of options below for how to override the default settings.  Run perldoc imap-report.pl to for further instruction.)

=head1 ACKNOWLEDGEMENTS

Built largely using Mail::IMAPClient currently maintained by E<lt>L<PLOBBES|http://search.cpan.org/~plobbes/>E<gt>
and the Term::Menus module by E<lt>L<REEDFISH|http://search.cpan.org/~reedfish/>E<gt>, along with
String::ProgressBar from E<lt>L<AHERNIT|http://search.cpan.org/~ahernit/>E<gt>.

=head1 TODO

=over 15

=item Better caching method.

=item Implement a cache aging mechanism.

=item Better report action handling.

=item Function to produce a report on any header field.

=item Better pager handling.

=item Clean up all that recon() rubbish.

=item Code refactoring...

=item

=back

=head1 LICENSE AND COPYRIGHT

Copyright (c) 2011 Andy Harrison

You can redistribute and modify this work under the conditions of the GPL.

=cut

# }}}

#  vim: set et ts=4 sts=4 sw=4 tw=80 nowrap ff=unix ft=perl fdm=marker :

