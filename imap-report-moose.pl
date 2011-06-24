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

# {{{ package IMAP::Report::Progress
#
# This is just the String::ProgressBar module with one
# single extra line of code.
#

package IMAP::Report::Progress;

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

    if ( ! defined $self->{rotation_counter} ) {
        $self->{rotation_counter} = 0;
    }

    $self->{rotation_counter} = ($self->{rotation_counter}+1) % (scalar(@matrix)-1);

    return $matrix[ $self->{rotation_counter} ];
}


1;

# }}}

# {{{ package IMAP::Report
#
package IMAP::Report;

use Moose;

has opts => (
    is  => 'rw',
    isa => 'HashRef',
);

has imap_options => (
    is         => 'rw',
    isa        => 'HashRef',
    auto_deref => 1,
);

has ssl_socket_close_options => (
    is         => 'rw',
    isa        => 'HashRef',
    auto_deref => 1,
);


# {{{ cache_init
#
sub cache_init {

    my $self = shift;
    my $opts = $self->opts;

    my $irc = IMAP::Report::Cache->new({ file => $opts->{cache_file} });

    my $cache = $irc->init;

    return $cache;

} # }}}

# {{{ header table
#
sub headers {

    my $self = shift;

    my $args = shift;

    # Lazy imap to english translation table
    #
    my %q_header_table = (

                    'DATE'                          => 'INTERNALDATE',
                    'INTERNALDATE'                  => 'DATE',

                    'SUBJECT'                       => 'BODY[HEADER.FIELDS (SUBJECT)]',
                    'BODY[HEADER.FIELDS (SUBJECT)]' => 'SUBJECT',

                    'SIZE'                          => 'RFC822.SIZE',
                    'RFC822.SIZE'                   => 'SIZE',

                    'FULLHEADERS'                   => 'RFC822.HEADER',
                    'RFC822.HEADER'                 => 'FULLHEADERS',

                    'TO'                            => 'BODY[HEADER.FIELDS (TO)]',
                    'BODY[HEADER.FIELDS (TO)]'      => 'TO',

                    'FROM'                          => 'BODY[HEADER.FIELDS (FROM)]',
                    'BODY[HEADER.FIELDS (FROM)]'    => 'FROM',

                    'LISTID'                        => 'BODY[HEADER.FIELDS (LIST-ID)]',
                    'BODY[HEADER.FIELDS (LIST-ID)]' => 'LISTID',

    );

    my %qq_header_table = (

                        'DATE'                            => 'INTERNALDATE',
                        'INTERNALDATE'                    => 'DATE',

                        'SUBJECT'                         => 'BODY[HEADER.FIELDS ("SUBJECT")]',
                        'BODY[HEADER.FIELDS ("SUBJECT")]' => 'SUBJECT',

                        'SIZE'                            => 'RFC822.SIZE',
                        'RFC822.SIZE'                     => 'SIZE',

                        'FULLHEADERS'                     => 'RFC822.HEADER',
                        'RFC822.HEADER'                   => 'FULLHEADERS',

                        'TO'                              => 'BODY[HEADER.FIELDS ("TO")]',
                        'BODY[HEADER.FIELDS ("TO")]'      => 'TO',

                        'FROM'                            => 'BODY[HEADER.FIELDS ("FROM")]',
                        'BODY[HEADER.FIELDS ("FROM")]'    => 'FROM',

                        'LISTID'                          => 'BODY[HEADER.FIELDS ("LIST-ID")]',
                        'BODY[HEADER.FIELDS ("LIST-ID")]' => 'LISTID',

                        );

    return 
        defined $args->{quote_headers} && $args->{quote_headers}
        ? %qq_header_table
        : %q_header_table
        ;

} # }}}

# {{{ imap_folders
#
# Read the raw list of folders directly from the imap server, no cache checking.
#
# Takes no args, returns a plain list of foldernames.
#
sub imap_folders {

    my $self = shift;
    my $args = shift;

    my $opts                     = $self->opts;
    my %imap_options             = $self->imap_options;
    my %ssl_socket_close_options = $self->ssl_socket_close_options;

    my $if_socket = 
        $opts->{Ssl}
        ? create_ssl_socket( 'if_socket' )
        : 0
        ;

    if ( $if_socket ) {
        $imap_options{Socket} = $if_socket;
    }

    my $if_imap = Mail::IMAPClient->new( %imap_options )
        or die "Cannot connect to host : $@";

    verbose( "Fetching folder list from IMAP server..." );

    my $list = $if_imap->folders
        or die_clean( 1, "Error fetching folders: $!\n" . $if_imap->LastError );

    $if_imap->disconnect;

    if ( $opts->{Ssl} ) {
        $if_socket->close( %ssl_socket_close_options );
    }

    return @$list;

}

# }}}

# {{{ types
#
sub types {

    # These are the types of reports that can be run.
    #
    return {
             all_folders_message_count_report       => 'Total count of messages in ALL folders',
             all_folders_message_sizes_report       => 'Total size of messages in ALL folders',
             all_folders_biggest_message_report     => 'Total list of biggest messages in ALL folders',
             all_folders_list_ids_report            => 'Total summary of the message List-ID headers in ALL folders',
             all_folders_messages_by_subject_report => 'Total summary of the message Subject headers in ALL folders',
             messages_by_subject_report             => 'Folder statistics report for message SUBJECT',
             messages_by_list_id_report             => 'Folder statistics report for message LISTID',
             messages_by_from_address_report        => 'Folder statistics report for message FROM addresses',
             messages_by_to_address_report          => 'Folder statistics report for message TO addresses',
             biggest_messages_report                => 'Folder statistics report for message SIZE',
             size_report                            => 'Folder summary report for total size of messages',
             list                                   => 'Display the current list of folders',
           };

} # }}}

# {{{ sub convert_bytes
#
# For pretty printing byte numbers.
#
sub convert_bytes {

    my $bytes = shift;

    return unless $bytes;

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

    return '0 seconds' unless $seconds;

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

# {{{ convert_date_to_epoch
#
sub convert_date_to_epoch {

    my $date = shift;

    my $dm = Date::Manip::Date->new();

    $dm->parse($date);
    my $epoch = $dm->printf('%s');

    return $epoch
        ? $epoch
        : 0;

}    # }}}

# {{{ create_ssl_socket
#
sub create_ssl_socket {

    my $self = shift;

    my $opts = $self->opts;

    my $description = shift;

    # Redundant, just being cautious...
    #
    return unless $opts->{Ssl};

    my $s = IO::Socket::SSL->new(
        Proto                   => 'tcp',
        PeerAddr                => $opts->{server},
        PeerPort                => $opts->{port},
        SSL_create_ctx_callback => sub { my $ctx = shift;
                                        ddump( 'ssl_ctx', $ctx ) if $opts->{debug};
                                        ddump( 'ssl_ctx_callback_description', $description ) if $opts->{debug};
                                        Net::SSLeay::CTX_sess_set_cache_size( $ctx, 128 ); },
    );

    # no EFFING output buffering...
    #
    select ($s);
    $| = 1;
    select (STDOUT);
    $| = 1;

    $s->verify_hostname( $opts->{server},'imap' )
        or warn "Error running verify_hostname: $!\n";

    return $s;

} # }}}

# {{{ tabulator
#
# Pretty print some rows in a dynamic width table...
#
# Use an anon hashref to pass in an arrayref of rows and a corresponding
# arrayref of column names.
#
# returns a list of the rows of the table.
#
sub tabulator {

    my $self = shift;
    my $args = shift;

    my $rows    = $args->{rows};
    my $columns = $args->{columns};

    my $header  =
        defined $args->{header} && $args->{header}
        ? $args->{header}
        : ''
        ;

    my $cols    = scalar(@$columns);
    my $pad     = 2;
    my $widths  = [];

    my @tabbed;

    # Dynamically calculate column widths for our list of messages
    #
    for my $row ( @$rows ) {
        for ( 0..$#$columns ) {
            $widths->[$_] = max( $widths->[$_],  length $row->[$_] );
        }
    }


    # Loop through one more time in case any of our column names are wider than
    # the column data...
    #
    for my $col ( @$columns ) {
        for ( 0..$#$columns ) {
            $widths->[$_] = max( $widths->[$_],  length $columns->[$_] );
        }
    }

    # Create our format string to feed to sprintf...
    #
    my $format = '';
    for ( @$widths ) {
        $format .= "%-${_}s";
        $format .= ' ' x $pad;
    }

    my $dashes = [];

    # Underline the column names.
    #
    for ( 0..$#$columns ) {
        push @$dashes, '-' x $widths->[$_];
    }

    unshift( @$rows, $dashes );
    unshift( @$rows, $columns );

    # Now turn each row into a dynamically constructed tabular report and pass
    # it back...
    #
    for ( @$rows ) {
        my $cur_row = sprintf( $format, @$_ );
        push @tabbed, $cur_row . "\n";
    }

    return @tabbed;

} # }}}

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

    my $self = shift;
    my $opts = $self->opts;

    my $report = shift;

    my $file           = "$ENV{HOME}/imap-report.txt";
    my $running_report = "$ENV{HOME}/imap-running-report.txt";

    open ( RPT, ">" . $file )
        or die_clean( 1, "Unable to write report.\n" );

    open ( RRPT, ">>" . $running_report )
        or die_clean( 1, "Unable to write running report.\n" );

    print RPT  $_ for @$report;
    print RRPT $_ for @$report;

    close RPT;
    close RRPT;

    system( "less -niSRX $file" );

    #system( "cat $file" );

    die_clean( 0, "Quitting" )
        if $opts->{list};

} # }}}



# }}}

# {{{ package IMAP::Report::Headers
#
package IMAP::Report::Headers;

use Moose::Role;

requires 'name';
requires 'imap_name';

1;

package IMAP::Report::Headers::TO;

use Moose;

with 'IMAP::Report::Headers';

sub name      {'TO'}
sub imap_name {'BODY[HEADER.FIELDS (TO)]'}

1;

package IMAP::Report::Headers::FROM;

use Moose;

with 'IMAP::Report::Headers';

sub name      {'FROM'}
sub imap_name {'BODY[HEADER.FIELDS (FROM)]'}

1;

package IMAP::Report::Headers::SUBJECT;

use Moose;

with 'IMAP::Report::Headers';

sub name      {'SUBJECT'}
sub imap_name {'BODY[HEADER.FIELDS (SUBJECT)]'}

1;

package IMAP::Report::Headers::DATE;

use Moose;

with 'IMAP::Report::Headers';

sub name      {'DATE'}
sub imap_name {'INTERNALDATE'}

1;

package IMAP::Report::Headers::SIZE;

use Moose;

with 'IMAP::Report::Headers';

sub name      {'SIZE'}
sub imap_name {'RFC822.SIZE'}

1;

package IMAP::Report::Headers::LISTID;

use Moose;

with 'IMAP::Report::Headers';

sub name      {'LISTID'}
sub imap_name {'BODY[HEADER.FIELDS (LIST-ID)]'}

1;

package IMAP::Report::Headers::FULLHEADERS;

use Moose;

with 'IMAP::Report::Headers';

sub name      {'FULLHEADERS'}
sub imap_name {'RFC822.HEADER'}

1;

# }}}

# {{{ package IMAP::Report::Folders
#
package IMAP::Report::Folders;

use Moose;

extends 'IMAP::Report';

# {{{ fetch_folders
#
# Expects to receive an anon hashref of arguments containing lists for filters
# and excludes.
#
# Returns plain list of folders after filtering and validating.
#
# (Only validated folders are added to the cached list of folders.)
#
sub fetch_folders {

    my $self = shift;
    my $args = shift;

    my $opts = $self->opts;

    my $ff_cache = $args->{cache};

    return unless $ff_cache;

    my $list         = [];
    my $menu_folders = [];

    $list = cache_check({ cache => $ff_cache, content_type => 'folder_list' });

    my @filtered_cached;
    my $cached_count = 0;

    if ( defined $list && ref $list eq 'ARRAY' && scalar(@$list) ) {

        @filtered_cached = filter_folders({ folders => $list });

        $cached_count = scalar(@filtered_cached);

    }

    if ( $opts->{cache_only} ) {

        if ( ! scalar(@filtered_cached) ) {
            die_clean( 1, "No folders..." );
        }

        return @filtered_cached;

    }

    my @imap_list;

    @imap_list = imap_folders();

    if ( ! scalar @imap_list ) {
        die_clean( 1, "No folders..." );
    }

    my @filtered_imap = filter_folders({ folders => \@imap_list });

    if ( ! scalar @filtered_imap ) {
        die_clean( 1, "No folders..." );
    }

    my $imap_count = scalar(@filtered_imap);

    # If the number of folders on the server matches the number of folders in
    # cache, then don't bother to validate and store the list of folders.
    #
    # TODO
    #
    # Obviously just comparing the number of folders isn't going to detect an
    # actual difference between real and cached...
    #
    if ( $cached_count == $imap_count ) {
        return @filtered_imap;
    }

    # Only folders that have been validated will be cached.  A bit of a slow
    # operation, but important.
    #
    print "\n\nValidating list of IMAP folders...\n";

    my @valid_list = validate_folders({ folders => \@filtered_imap });

    for ( @valid_list ) {
        cache_put({ cache => $ff_cache, content_type => 'validated_folder_list', folder => $_ });
    }

    return @valid_list;

} # }}}

# {{{ filter_folders
#
# Takes a list of folders, processes the includes and excludes, and returns the
# pared down list.
#
sub filter_folders {

    my $self = shift;
    my $args = shift;
    my $opts = $self->opts;
    my $list = $args->{folders};

    return unless $list;

    my $filter_list =
        $opts->{filters}
        ? $opts->{filters}
        : []
        ;

    my $excludes_list =
        $opts->{exclude}
        ? $opts->{exclude}
        : []
        ;

    # Some hard coded excludes...
    #
    my @extra_excludes = qw/[Gmail]/;

    # Here's where we filter out what we want from the list of folders.
    #
    my @filtered_list =
                        grep { my $item = $_; not grep { $item eq $_     } @extra_excludes }
                        grep { my $item = $_; not grep { $item =~ m/$_/i } @$excludes_list }
        @$filter_list ? grep { my $item = $_;     grep { $item =~ m/$_/i } @$filter_list   } @$list : @$list;

    return @filtered_list;

} # }}}

# {{{ get_folder_size
#
# TODO
#
# Simplify this mess.
#
sub get_folder_size {

    my $self = shift;
    my $args = shift;

    my $opts = $self->opts;

    my $folder    = $args->{folder};
    my $gfs_cache = $args->{cache};

    print "\n\n\n\nFetching message details for folder '$folder'...\n";

    my $msg_count = fetch_messages({ cache => $gfs_cache, folder => $folder });

    my $fetched_messages = cache_report({ folder      => $folder,
                                          cache       => $gfs_cache,
                                          report_type => 'total_folder_size' });


    return ( 0, 0 ) unless scalar(@$fetched_messages);

    my $totalsize;

    my $counter = 0;

    for ( @$fetched_messages ) {
       #$totalsize += $fetched_messages->{$_}->{$header_table{'Size'}};
        $totalsize += $_->[0];
        $counter++;
    }

    ddump( 'fetched_messages', $fetched_messages ) if $opts->{debug};

    if ( ! $counter ) {
        show_error( "Error: No messages to report for '$folder'..." );
        return ( 0, 0 );
    }

    return ( $totalsize, $counter );

} # }}}

# }}}

# {{{ package IMAP::Report::Folders::Cache
#
package IMAP::Report::Folders::Cache;

use Moose;

extends 'IMAP::Report::Folders';

# {{{ check
#
sub check {

    my $self = shift;

    my $args = shift;

    my $opts = $self->opts;

    my $dbh          = $self->{cache};
    my $content_type = $args->{content_type};
    my $value        = defined $args->{value} ? $args->{value} : '';

    my $cur_time = time;

        # {{{ folder_list cache check

        # Checks the cached list of folders and returns an arrayref list of
        # them.

        cache_prune({ cache        => $dbh,
                      content_type => $content_type });

        # If we're in cache_only mode, instead of grabbing the stored list of
        # folders, we'll create a list of folders from the actual messages
        # stored in the cache.
        #
        my $sql =
            $opts->{cache_only}
            ? q[
                    SELECT DISTINCT
                        folder
                    FROM
                        messages
                    WHERE
                        server = ?
                        AND username = ?
              ]

            : q[

                    SELECT
                        folder
                    FROM
                        folders
                    WHERE
                        server = ?
                        AND username = ?
                        AND validated = 1

               ]
            ;

        my $folderlist = [];

        push @$folderlist, $_->{folder}
            for @{ $dbh->selectall_arrayref( $sql, { Slice => {} },
                                             $opts->{server}, $opts->{user} ) };

        if ( scalar(@$folderlist) ) {
            return $folderlist;
        } else {
            return;
        }

        # }}}

    return;

} # }}}

# }}}

# {{{ package IMAP::Report::Messages
#
package IMAP::Report::Messages;

use Moose;

extends 'IMAP::Report';

# {{{ fetch_messages
#
# Expects to receive an anon hashref of options with at least the folder name
# and the cache object.
#
# Determines whether the folder cache needs to be updated and takes care of it
# as-needed.
#
# Returns the count of messages present in the message cache.
#
sub fetch_messages {

    my $self = shift;
    my $args = shift;
    my $opts = $self->opts;

    my %imap_options             = $self->imap_options;
    my %ssl_socket_close_options = $self->ssl_socket_close_options;

    my %header_table = (
                         IMAP::Report::Headers::TO->name          => IMAP::Report::Headers::TO->imap_name,
                         IMAP::Report::Headers::FROM->name        => IMAP::Report::Headers::FROM->imap_name,
                         IMAP::Report::Headers::SUBJECT->name     => IMAP::Report::Headers::SUBJECT->imap_name,
                         IMAP::Report::Headers::DATE->name        => IMAP::Report::Headers::DATE->imap_name,
                         IMAP::Report::Headers::SIZE->name        => IMAP::Report::Headers::SIZE->imap_name,
                         IMAP::Report::Headers::LISTID->name      => IMAP::Report::Headers::LISTID->imap_name,
                         IMAP::Report::Headers::FULLHEADERS->name => IMAP::Report::Headers::FULLHEADERS->imap_name,
                       );

    my $folder  = $args->{folder};
    my $f_cache = $args->{cache};

    return unless $folder;
    return unless $f_cache;

    my $cached_count = cache_check({ cache => $f_cache, content_type => 'fetched_messages', value => $folder });

    return $cached_count if $opts->{cache_only};

    my $f_imap_socket =
        $opts->{Ssl}
        ? create_ssl_socket( 'f_imap_socket' )
        : 0
        ;

    if ( $f_imap_socket ) {
        $imap_options{Socket} = $f_imap_socket;
    }

    my $f_imap = Mail::IMAPClient->new(%imap_options)
        or die "Cannot connect to host : $@";

    my @headers;

    # TODO
    #
    # Fix this header handling...
    #
    push @headers, $header_table{$_} for qw/DATE SUBJECT SIZE TO FROM LISTID FULLHEADERS/;
   #push @headers, $f_imap->Quote($header_table{$_}) for qw/DATE SUBJECT SIZE TO FROM LISTID/;

#   push @headers, qw(DATE SIZE);
#   push @headers, 'BODY[HEADER.FIELDS (' . $f_imap->Quote($_) . ')]' 
#       for qw/SUBJECT TO FROM LISTID/;

    ddump( 'headers', \@headers ) if $opts->{debug};

    $f_imap->examine( $folder );

    my $num = $f_imap->message_count;

    # If our threshold is a percentage, calculate it now...
    #
    my $threshold =
        $opts->{threshold_percentage}
        ? $num * $opts->{threshold_percentage}
        : $opts->{threshold}
        ;


    # Make sure all the threads are going to have something to do...
    #
    my $skip_threads =
        $num <= $opts->{threads} && $num < $opts->{min_for_threads}
        ? 1
        : 0
        ;

    print "Server says selected folder '$folder' contains $num messages...\n";

    # Compare the count of cached messages against the actual number of messages
    # in the imap mailbox folder.  If the difference is less than our threshold,
    # don't bother with the imap fetch operation.
    #
    if ( $cached_count && abs( $cached_count - $num ) <= $threshold ) {

        print "Cache contains $cached_count messages for this folder...\n"
            . "Using cached messages only for this folder.\n\n";

        return $cached_count;

    } elsif ( $cached_count == 0 && abs( $cached_count - $num ) <= $threshold ) {
        print "Server says this folder is empty, no cached messages present, skipping update.\n\n\n";
        return 0;
    } else {
        print "Cache contains $cached_count messages...\n"
            . "Difference is more than threshold ($threshold), cache needs to be updated.\n";
    }

    print "Updating cache for folder '$folder'\n\n";

    my $fetched = {};

    if ( $f_imap->Folder() ) {

        # This object will hold our message ids.
        #
        my $msgset = Mail::IMAPClient::MessageSet->new( $f_imap->messages );

        my $msg_ids = [];

        if ( $msgset ) {
            $msg_ids = $msgset->unfold;
        }

        my $msg_count =
            defined $msg_ids && $msg_ids && scalar(@$msg_ids) > 0
            ? scalar(@$msg_ids)
            : 0
            ;

        if ( ! $msg_count ) {
            print "No messages found in folder...\n\n\n\n";
            return;
        }

        my $previously_fetched = cache_check({ cache => $f_cache, content_type => 'messages_previously_fetched', folder => $folder });

        my $ids_to_fetch = [];

        if ( $previously_fetched && scalar(@$previously_fetched) > 0 ) {
            for my $cur_id ( @$msg_ids ) {
                push @$ids_to_fetch, $cur_id
                    unless grep $cur_id eq $_, @$previously_fetched;
            }
        } else {
            $ids_to_fetch = $msg_ids;
        }

       #show_error( "IDS TO FETCH: " . Dumper( $ids_to_fetch ) );

       my $ids_to_fetch_count = scalar(@$ids_to_fetch);

        if ( ! $ids_to_fetch && ! scalar(@$ids_to_fetch) ) {
            print "\n\nAll messages for folder '$folder' have been fetched... skipping.\n\n";
            return;
        }

        cache_put(
                   {
                     cache        => $f_cache,
                     folder       => $folder,
                     content_type => 'unfetched_message_ids_from_folder',
                     values       => $ids_to_fetch
                   }
                 );

        ddump( 'ids_to_fetch', $ids_to_fetch ) if $opts->{debug};

        # TODO
        #
        # FIX.
        #
        my $use_threaded_mode = 0;
        my $skip_threads      = 1;

        if ( $use_threaded_mode && ! $skip_threads ) {

            # {{{ Threaded mode
            #

            $fetched =
                threaded_fetch_msgs(
                                     {
                                       cache         => $f_cache,
                                       folder        => $folder,
                                       message_count => $ids_to_fetch_count
                                     }
                                   );

            ddump( 'fetched', $fetched ) if $opts->{debug};

            if ( $fetched && ! ref $fetched eq 'HASH' ) {
                show_error( "No messages fetched by threads!\n" );
                return;
            }

            my $fetched_ids = [];

            for ( sort keys %$fetched ) {
                push @$fetched_ids, $_;
            }

            cache_put(
                       {
                         cache        => $f_cache,
                         content_type => 'fetched_messages',
                         folder       => $folder,
                         values       => $fetched
                       }
                     );

            cache_put(
                       {
                         cache        => $f_cache,
                         content_type => 'update_message_fetch_status',
                         folder       => $folder,
                         values       => $fetched_ids
                       }
                     );

            # }}}

        } else {

            # {{{ Non-threaded mode
            #

            print "Fetching messages for folder: '$folder'\n\n\n";


            my $sbar = IMAP::Report::Progress->new( length => 10,
                                            max    => $num,
                                            show_rotation => 1 );

            my $scounter = 0;

            #$sbar->text('fetching messages');

            my $total_msgs_added = 0;

            my $done = 0;
            my $offset = 0;

            # Keep looping while we get message id's from the cache.
            #
            while ( $ids_to_fetch_count > 0 ) {

                my $cur_block =
                    cache_check(
                                 {
                                   cache        => $f_cache,
                                   content_type => 'messages_to_be_fetched',
                                   folder       => $folder,
                                   limit        => $opts->{max_fetch},
                                   offset       => $offset,
                                 }
                               );

                ddump( 'cur_block', $cur_block ) if $opts->{debug};

                unless ( $cur_block && scalar(@$cur_block) > 0 ) {
                    print "\n\n\n\n\nNo messages remaining to be fetched...\n\n\n\n";
                    last;
                }

                my $cur_block_count = scalar(@$cur_block);

                # Make a M::I::MS object to get a clean range of message ids.
                #
                my $cur_msgset  = Mail::IMAPClient::MessageSet->new(@$cur_block);
                my $cur_msg_ids = $cur_msgset->unfold;

                $sbar->info( 'Fetching next ' . $cur_block_count . ' messages...' );
                $sbar->write;

                unless ( $f_imap->noop or $f_imap->reconnect ) {
                    show_error( "reconnect failed: $@\n" . $f_imap->LastError );
                    next;
                }

                ddump( 'headers_being_fetched', \@headers );

               #my $test_fetch = $f_imap->parse_headers( $cur_msg_ids, 'DATE', 'RFC822.SIZE', 'SUBJECT', 'TO', 'FROM' );
               #ddump( 'test_fetch_parse_headers', $test_fetch ) if $opts->{debug};

                $fetched = $f_imap->fetch_hash( $cur_msg_ids, @headers );


                ddump( 'fetched',   $fetched )           if $opts->{debug};
                ddump( 'LastError', $f_imap->LastError ) if $opts->{debug};

                my $cur_count = ( scalar( keys %$fetched ) );

                # Shouldn't be necessary, just being safe....
                #
                next unless $cur_count;

                $sbar->info( 'Processing message headers...' );
                $sbar->write;

                my $fetched_ids = [];
                my $counter = 0;

                # Ugly.
                #
                # Iterate the whole list of fetched messages and fix each value returned.
                #
                for my $cur_id ( keys %$fetched ) {

                    for my $cur_header (@headers) {

                        $fetched->{$cur_id}->{$cur_header} =
                            stripper( $header_table{$cur_header},
                                    $fetched->{$cur_id}->{$cur_header} );

                    }

                   # if ( $opts->{search} ) {

                   #     my $search =
                   #         generate_search_string({ folder => $folder,
                   #                                  date   => $DATE,
                   #                                  header => 'Subject',
                   #                                  value  => $fetched->{$cur_id}->{$header_table{Subject}} });

                   #     $fetched->{$cur_id}->{Search} = $search;

                   # }


                    if ( ( ( $counter++ % 10 ) + 1 ) == 10 ) {

                        $sbar->update( $total_msgs_added );
                        $sbar->write;

                    }

                    push @$fetched_ids, $cur_id;

                }

                # Store our results in the cache then return the results.
                #
                verbose( "\n\n\n\n\nStoring $cur_count messages in cache...\n" );

                $sbar->info( "Storing $cur_count messages in cache..." );
                $sbar->write;

                cache_put({ cache => $f_cache, content_type => 'fetched_messages', folder => $folder, values => $fetched });
                cache_put({ cache => $f_cache, content_type => 'update_message_fetch_status', folder => $folder, values => $fetched_ids });

                $total_msgs_added += $cur_count;

                $sbar->info( 'Finished caching current sequence...' );
                $sbar->update( $total_msgs_added );
                $sbar->write;

                $ids_to_fetch_count -= $cur_count;
                sleep 1;

            }

            # }}}

        }

    } else {

        die_clean( 1, "Error checking current folder selection: $! " . $f_imap->LastError );

    }


    $f_imap->disconnect;

    if ( $opts->{Ssl} ) {
        $f_imap_socket->close( %ssl_socket_close_options );
    }

    return $num;


} # }}}

# {{{ stripper
#
# Oddly, the chomp function behaves in an unexpected way on subjects and other
# headers returned from the imap server.  I know it's an issue of LF vs. CR, but
# I still couldn't get it to behave cleanly, so I did it this way rather than
# any chomp chop chomp monkey business...
#
sub stripper {

    my $self = shift;
    my $opts = $self->opts;

    my $name  = shift;
    my $field = shift;

    return unless $name;

    ddump( 'before_stripping_name',  $name )  if $opts->{debug};
    ddump( 'before_stripping_field', $field ) if $opts->{debug};

    if ( $name eq 'FULLHEADERS' ) {

        # For the full headers column, just do a small amount of sanitizing.
        #
        my @fheaders = split( "\r\n", $field );

        my @h;

        for ( @fheaders ) {

            s/[\r\n\t]+/ /g;      # CRLF and tabs
            s/\R+/ /g;            #
            s/\s+$//;             # trailing spaces
            s/\s+/ /;             # multi spaces
            s/\\//g;              # escapes

            push @h, $_;

        }

        return join( "\n", @h );

    }

    $field =~ s/[\r\n\t]+/ /g;      # CRLF and tabs
    $field =~ s/\R+/ /g;            #

    $field =~ s/^\s+//;             # leading spaces
    $field =~ s/\s+$//;             # trailing spaces
    $field =~ s/\s+/ /;             # multi spaces

    $field =~ s/\\//g;              # escapes

    # And one more for good measure...
    #
    $field =~ s/[[:^print:]]//g;    # Non-printables

    if ( ! $field ) {
        if ( $name eq 'LISTID' ) {
            $field = '(no list)';
        } else {
            $field = ']]EmptyField[[';
        }
    }

    # Strip off the name of the envelope attribute
    #
    if ( $field =~ m/^$name:\s+(.*)$/i ) {
        $field = $1;
    }

    if ( $name eq 'LISTID' && $field =~ m/^List-Id:\s+(.*)$/i ) {
        $field = $1;
    }

    # For from addresses, just grab the address.
    #
    if ( $name eq 'FROM' or $name eq 'TO' ) {

        $field = lc($field);

        my @addrs       = Mail::Address->parse($field);
        my $addr_obj    = $addrs[0];
        my $parsed_addr = $addr_obj->address;

        # Just wanted a little visual distinction for which test the field was
        # failing...
        #
        if ( ! $parsed_addr ) {
            $parsed_addr = ']]Empty-Address-Field[[';
        }

        if ( length $parsed_addr <= 3 ) {
            $parsed_addr = ']]Empty_Address_Field[[';
        }

        $field = $parsed_addr;

    }

    # For Dates, convert to epoch
    #
    if ( $name eq 'DATE' ) {

        my $epoch = convert_date_to_epoch($field);

        if ( ! $epoch ) {
            $epoch = 1;
        }

        $field = $epoch;
    }

    # Strip off the subject cruft...
    #
    if ( $name eq 'SUBJECT' ) {
        $field =~ s/^Re:\s+//gi;
        $field =~ s/^Fwd:\s+//gi;
        $field =~ s/\s+Re:\s+//gi;
        $field =~ s/\s+Fwd:\s+//gi;
    }


    if ( ! $field ) {
        $field = ']]EMPTY[[';
    }

    ddump( 'after_stripping_field', $field ) if $opts->{debug};

    return $field;

} # }}}

# {{{ max
#
sub max {

    my ( $a, $b ) = @_;

    $a = 0 unless $a;

    return $a > $b
        ? $a
        : $b;

} # }}}

# {{{ get_list_ids
#
# TODO
#
# Simplify this mess.
#
sub get_list_ids {

    my $self = shift;
    my $args = shift;

    my $opts = $self->opts;

    my $folder    = $args->{folder};
    my $gli_cache = $args->{cache};

    return unless $folder;
    return unless $gli_cache;

    my $raw_report         = [];
    my $total_num_messages = 0;

    my $update_count = fetch_messages({ folder => $folder, cache => $gli_cache });

    my $fetched_details = cache_report({ folder      => $folder,
                                         cache       => $gli_cache,
                                         report_type => 'all_list_ids' });

    ddump( 'fetched_details',     $fetched_details )     if $opts->{debug};
    ddump( 'ref fetched_details', ref $fetched_details ) if $opts->{debug};

    for my $cur_msg ( @$fetched_details ) {
        push @$raw_report, [ $cur_msg->[0], $folder, $cur_msg->[1] ];
    }

    return @$raw_report;

} # }}}

# }}}

# {{{ package IMAP::Report::Messages::Cache
#
package IMAP::Report::Messages::Cache;

use Moose;

extends 'IMAP::Report::Messages';

# {{{ check
#
sub check {

    my $self = shift;

    my $args = shift;

    my $opts = $self->opts;

    my $dbh          = $self->{cache};
    my $content_type = $args->{content_type};
    my $value        = defined $args->{value} ? $args->{value} : '';

    my $cur_time = time;

        # {{{ fetched messages cache check

        return unless defined $value && $value;

        cache_prune({ cache        => $dbh,
                      content_type => $content_type });

        my $sql = q[
            SELECT
                count(msg_id)
            FROM
                messages
            WHERE
                server = ?
                AND username = ?
                AND folder = ?
        ];

        my $sth = $dbh->prepare( $sql );

        $sth->execute( $opts->{server}, $opts->{user}, $value );

        my $count = $sth->fetch;

        return
            $count->[0]
            ? $count->[0]
            : 0
            ;

        # }}}

    return;

} # }}}

# {{{ put
#
# Handle inserting the various types of information we want to cache.
#
# Sticks in the current time value so for cache aging purposes later.
#
sub put {

    my $self = shift;
    my $args = shift;

    my $opts = $self->opts;

    my $dbh          = $args->{cache};
    my $content_type = $args->{content_type};
    my $values       = $args->{values};
    my $folder       = $args->{folder};

    my %header_table = (
                         IMAP::Report::Headers::TO->name          => IMAP::Report::Headers::TO->imap_name,
                         IMAP::Report::Headers::FROM->name        => IMAP::Report::Headers::FROM->imap_name,
                         IMAP::Report::Headers::SUBJECT->name     => IMAP::Report::Headers::SUBJECT->imap_name,
                         IMAP::Report::Headers::DATE->name        => IMAP::Report::Headers::DATE->imap_name,
                         IMAP::Report::Headers::SIZE->name        => IMAP::Report::Headers::SIZE->imap_name,
                         IMAP::Report::Headers::LISTID->name      => IMAP::Report::Headers::LISTID->imap_name,
                         IMAP::Report::Headers::FULLHEADERS->name => IMAP::Report::Headers::FULLHEADERS->imap_name,
                       );

    return unless $dbh;
    return unless $content_type;

        # {{{ fetched message cache population

        return unless defined $values && ref $values eq 'HASH';
        return unless $folder;


        #show_error( "PUTTING CACHE FOR FOLDER: $folder " . Dumper( $values ) );

        ddump( 'cache_put_values', $values ) if $opts->{debug};

        $dbh->begin_work;

        my $sql = q[
            INSERT OR REPLACE INTO messages (
                server,
                username,
                msg_id,
                folder,
                "TO",
                "FROM",
                SUBJECT,
                DATE,
                SIZE,
                LISTID,
                FULLHEADERS,
                last_update
            ) VALUES (
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?
            )
        ];

       #my $mcount = scalar( keys %$values );

       #my $cbar = IMAP::Report::Progress->new( max    => $mcount,
       #                                length => 10 );

       #$cbar->text('Caching:');
       #$cbar->info('messages');

       #my $counter = 0;

       #$cbar->update( $counter++ );
       #$cbar->write;

        my $sth = $dbh->prepare($sql);

        my $in_time = time;

        for ( keys %$values ) {
            my $result = $sth->execute(
                $opts->{server},
                $opts->{user},
                $_,
                $folder,
                $values->{$_}->{ $header_table{TO} },
                $values->{$_}->{ $header_table{FROM} },
                $values->{$_}->{ $header_table{SUBJECT} },
                $values->{$_}->{ $header_table{DATE} },
                $values->{$_}->{ $header_table{SIZE} },
                $values->{$_}->{ $header_table{LISTID} },
                $values->{$_}->{ $header_table{FULLHEADERS} },
                $in_time
            );

            if ( $dbh->errstr ) {
                show_error( 'Message cache insert error: ' . $dbh->errstr );
                $dbh->rollback;
                return;
            }

            #if ( ( ( $counter % 10 ) + 1 ) == 10 ) {
            #   $cbar->update( $counter++ );
            #   $cbar->write;
            #}

        }

        $dbh->commit;

        # }}}

    return;

} # }}}

# {{{ update_status
#
sub update_status {

    my $self = shift;
    my $args = shift;
    my $opts = $self->opts;

    my $dbh          = $args->{cache};
    my $content_type = $args->{content_type};
    my $values       = $args->{values};
    my $folder       = $args->{folder};

    return unless $dbh;
    return unless $content_type;

        # {{{ update the message id's status to indicate that it has been fetched

        return unless defined $values && ref $values eq 'ARRAY';
        return unless $folder;

        my $cur_time = time;

        ddump( 'cache_put_values->update_message_fetch_status', $values ) if $opts->{debug};

        $dbh->begin_work;

        my $sql = q[
            INSERT OR REPLACE INTO fetchlist (
                server,
                username,
                folder,
                msg_id,
                last_update
            ) VALUES (
                ?,
                ?,
                ?,
                ?,
                ?
            )
        ];

       #my $mcount = scalar( keys %$values );

       #my $cbar = IMAP::Report::Progress->new( max    => $mcount,
       #                                length => 10 );

       #$cbar->text('Caching:');
       #$cbar->info('messages');

       #my $counter = 0;

       #$cbar->update( $counter++ );
       #$cbar->write;

        my $sth = $dbh->prepare($sql);

        for ( @$values ) {

            my $result = $sth->execute(
                $opts->{server},
                $opts->{user},
                $folder,
                $_,
                $cur_time
            );

            if ( $dbh->errstr ) {
                show_error( 'Message fetchlist cache insert error: ' . $dbh->errstr );
                $dbh->rollback;
                return;
            }

            #if ( ( ( $counter % 10 ) + 1 ) == 10 ) {
            #   $cbar->update( $counter++ );
            #   $cbar->write;
            #}

        }

        $dbh->commit;

        # }}}

    return;

} # }}}

# }}}

# {{{ package IMAP::Report::Messages::ToFetch
#
package IMAP::Report::Messages::ToFetch;

use Moose;

extends 'IMAP::Report::Messages';

# {{{ cache_check
#
# My crude method of a caching mechanism.
#
# Feed this function the type of content and value for which to look.  Returns
# cached elements based on the different types of information, arrayrefs,
# hashrefs, bools, etc.
#
# TODO
#
# Implement cache aging
#
sub cache_check {

    my $self = shift;

    my $args = shift;

    my $opts = $self->opts;

    my $dbh          = $self->{cache};
    my $content_type = $args->{content_type};
    my $value        = defined $args->{value} ? $args->{value} : '';

    my $cur_time = time;

        # {{{ return the list of message id's that need to be fetched.

        my $folder = $args->{folder};

        my $limit =
            $args->{limit}
            ? $args->{limit}
            : $opts->{max_fetch}
            ;

        my $offset = $args->{offset};

        ddump( 'folder', $folder ) if $opts->{debug};
        ddump( 'limit',  $limit  ) if $opts->{debug};
        ddump( 'offset', $offset ) if $opts->{debug};

        # TODO
        #
        # Add some locking here...

        return unless $folder;
        return unless $limit;
        return unless defined $offset;

        #$dbh->begin_work;

        my $sql = q[
            SELECT
                msg_id
            FROM
                fetchlist
            WHERE
                server = ?
                AND username = ?
                AND folder = ?
                AND last_update IS NULL
            ORDER BY
                msg_id
            LIMIT ?
            OFFSET ?
        ];

       #my $mcount = scalar( keys %$values );

       #my $cbar = IMAP::Report::Progress->new( max    => $mcount,
       #                                length => 10 );

       #$cbar->text('Caching:');
       #$cbar->info('messages');

       #my $counter = 0;

       #$cbar->update( $counter++ );
       #$cbar->write;

        my $folderlist = [];

        push @$folderlist, $_->{msg_id}
            for @{ $dbh->selectall_arrayref( $sql, { Slice => {} }, $opts->{server}, $opts->{user}, $folder, $limit, $offset ) };

        if ( $dbh->errstr ) {
            show_error( 'Messages fetched from fetchlist cache error: ' . $dbh->errstr );
            return;
        }

        #show_error( "messages_to_be_fetched_list: " . Dumper( $folderlist ) );

        if ( scalar(@$folderlist) ) {
            return $folderlist;
        } else {
            return;
        }

        #$dbh->commit;

        # }}}

    return;

} # }}}

# {{{ unfetched_messages
#
sub unfetched_messages {

    my $self = shift;
    my $args = shift;
    my $opts = $self->opts;

    my $dbh          = $args->{cache};
    my $content_type = $args->{content_type};
    my $values       = $args->{values};
    my $folder       = $args->{folder};

    return unless $dbh;
    return unless $content_type;

        # {{{ message id's present in the selected folder

        return unless defined $values && ref $values eq 'ARRAY';
        return unless $folder;

        #show_error( "PUTTING CACHE FOR FOLDER: $folder " . Dumper( $values ) );

        ddump( 'cache_put_values', $values ) if $opts->{debug};

        $dbh->begin_work;

        my $sql = q[
            INSERT OR REPLACE INTO fetchlist (
                server,
                username,
                folder,
                msg_id
            ) VALUES (
                ?,
                ?,
                ?,
                ?
            )
        ];

       #my $mcount = scalar( keys %$values );

       #my $cbar = IMAP::Report::Progress->new( max    => $mcount,
       #                                length => 10 );

       #$cbar->text('Caching:');
       #$cbar->info('messages');

       #my $counter = 0;

       #$cbar->update( $counter++ );
       #$cbar->write;

        my $sth = $dbh->prepare($sql);

        for ( @$values ) {

            my $result = $sth->execute(
                $opts->{server},
                $opts->{user},
                $folder,
                $_
            );

            if ( $dbh->errstr ) {
                show_error( 'Message fetchlist cache insert error: ' . $dbh->errstr );
                $dbh->rollback;
                return;
            }

            #if ( ( ( $counter % 10 ) + 1 ) == 10 ) {
            #   $cbar->update( $counter++ );
            #   $cbar->write;
            #}

        }

        $dbh->commit;

        # }}}

    return;

} # }}}

# }}}

# {{{ package IMAP::Report::Type
#
package IMAP::Report::Type;

use Moose::Role;

requires 'description';

requires 'report';

# {{{ show_report_types
#
sub show_report_types {

my $types = report_types();

my $rpts = [];

for ( sort { $types->{$a} cmp $types->{$b} } keys %$types ) {
    push @$rpts, [ $_, $types->{$_} ];
}

my @cur_report;

for ( tabulator({ rows    => $rpts, columns => [qw/Type Description/] }) ) {
    push @cur_report, $_;
}

print for @cur_report;

die_clean( 0, 'Quitting...' );

} # }}}

# {{{ package IMAP::Report::Type::all_folders_message_count_report
#
package IMAP::Report::Type::all_folders_message_count_report;

use Moose;

with 'IMAP::Report::Type';

sub description { 'Total count of messages in ALL folders' }

# {{{ report
#
# Displays a count of the number of messages in each folder.  While it can use
# cached values, the message counts themselves are not a cached value.
#
sub report {

    my $self                     = shift;
    my $args                     = shift;
    my $opts                     = $self->opts;
    my %imap_options             = $self->imap_options;
    my %ssl_socket_close_options = $self->ssl_socket_close_options;
    my $reports = $self->types;


    my $fmcr_cache = $args->{cache};

    my $report_type = $reports->{all_folders_message_count_report};

    my $total_message_count = 0;

    my @fsize_report = ();

    my $total_num_messages = 0;
    my $folders_counted    = 0;

    my @imap_folders = fetch_folders({ cache => $fmcr_cache });

    my $num_folders = scalar(@imap_folders);

    my $raw_report = {};

    my @uncached_folders;

    for ( @imap_folders ) {

        next if $_ eq '[Gmail]/All Mail';

        my $count = cache_check({ cache => $fmcr_cache, content_type => 'fetched_messages', value => $_ });

        if ( $count ) {

            # Add a couple of asterisks to easily see which folders are cached.
            #
            $raw_report->{ '**' . $_ } = $count;

            # We don't want the 'All Mail' gmail label to skew our total since
            # it represents ALL messages in a gmailbox.
            #
            $total_num_messages += $count;

        } else {

            push @uncached_folders, $_ unless $opts->{cache_only};

        }

    }

    print "\n\nFound $total_num_messages cached messages.\n\n"
        . "Now asking the imap server directly...\n\n";

    if ( scalar(@uncached_folders) ) {

        my $bar = IMAP::Report::Progress->new( max => $num_folders, length => 10 );

        my $fmcr_socket = 
            $opts->{Ssl}
            ? create_ssl_socket( 'fmcr_socket' )
            : 0
            ;

        if ( $fmcr_socket ) {
            $imap_options{Socket} = $fmcr_socket;
        }


        my $fmcr_imap = Mail::IMAPClient->new( %imap_options )
            or die "Cannot connect to host : $@";

        for ( @uncached_folders ) {

            next if $_ eq '[Gmail]/All Mail';

            $fmcr_imap->examine($_)
                or show_error( "Error selecting $_: $@\n" );

            $bar->info( "Sent: " . $fmcr_imap->Transaction . " STATUS \"$_\"" );

            my $count = $fmcr_imap->message_count();

            $raw_report->{$_} = $count if $count;

            $bar->update( ++$folders_counted );

            $bar->write;

            $total_num_messages += $count;

        }

        print "\n" x 5;

        $fmcr_imap->disconnect;

        if ( $opts->{Ssl} ) {
            $fmcr_socket->close( %ssl_socket_close_options );
        }

    }




    my @rows;

    # Iterate the hashref of raw data and populate a list of arrayrefs to feed
    # to the tabulator
    #
    for my $fc ( reverse sort { $raw_report->{$a} <=> $raw_report->{$b} } keys %$raw_report ) {
        push @rows, [ $raw_report->{$fc}, $fc ];
    }

    push @fsize_report, $_
        for tabulator({ rows    => \@rows,
                        columns => [qw/Folder Messages/] });


    push @fsize_report, "\n\n\n\n\n" . '-' x 60 . "\n\n";
    push @fsize_report, "Total messages found: $total_num_messages\n\n";

#   if ( ! $opts->{cache_only} && ref $fmcr_imap ) {
#       $fmcr_imap->disconnect;
#       $fmcr_socket->close( %ssl_socket_close_options );
#   }

    return @fsize_report;


} # }}}

1;

# }}}

# {{{ package IMAP::Report::Type::all_folders_message_sizes_report
#
package IMAP::Report::Type::all_folders_message_sizes_report;

use Moose;

with 'IMAP::Report::Type';

sub description { 'Total size of messages in ALL folders' }

# {{{ report
#
sub report {

    my $self = shift;
    my $args = shift;
    my $opts = $self->opts;

    my $reports = $self->types;

    my $folder     = $args->{folder};
    my $fmsr_cache = $args->{cache};

    my $report_type = $reports->{all_folders_message_sizes_report};

    my @imap_folders = fetch_folders({ cache => $fmsr_cache });

    # Extra press-enter-to-continue prompts to make damn sure that we understand
    # that this is long, heavy operation on a mailbox with a ton of messages.
    #
    if ( ! $opts->{cache_only} ) {

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
            . '-' x 60 . "\n"
            . join( "\n", @imap_folders )
            . "\n" . '-' x 60 . "\n"
            . "\n\nAbove is the list of folders that will be iterated for the report.\n\n"
            . "Last chance to cancel!\n\n"
        );

    }

    my @msize_report = ();
    my $raw_report   = {};


    push @msize_report, "SIZE\t\t\tCount\t\t\tFolder\n";
    push @msize_report, '-' x 60 . "\n";

    my $total_num_messages  = 0;
    my $total_message_bytes = 0;

    for ( @imap_folders ) {

        # Always skip the 'All Mail' gmail label.  This
        # represents every message, so it will just skew the
        # results.
        #
        if ( $_ eq '[Gmail]/All Mail' ) {
            print "...Skipping the $_ Folder...\n";
            next;
        }

        my ( $cur_size, $message_count ) = get_folder_size({ folder => $_,
                                                             cache  => $fmsr_cache });

        if ($cur_size) {
            $raw_report->{$_}->{SIZE}  = $cur_size;
            $raw_report->{$_}->{count} = $message_count;
            $total_num_messages += $message_count;
            $total_message_bytes += $cur_size;
        }

    }

    for my $cur_folder ( reverse sort { $raw_report->{$a}->{SIZE} <=> $raw_report->{$b}->{SIZE} } keys %$raw_report ) {

        push @msize_report,
              convert_bytes( $raw_report->{$cur_folder}->{SIZE} )
            . "\t" x 3
            . $raw_report->{$cur_folder}->{count}
            . "\t" x 3
            . $cur_folder
            . "\n";

    }

    push @msize_report, '-' x 60 . "\n\n";
    push @msize_report, "Total messages found: $total_num_messages\n\n";
    push @msize_report, 'Total size of all messages: ' . convert_bytes( $total_message_bytes ) . "\n\n";

    return @msize_report;

} # }}}

1;

# }}}

# {{{ package IMAP::Report::Type::all_folders_biggest_message_report
#
package IMAP::Report::Type::all_folders_biggest_message_report;

use Moose;

with 'IMAP::Report::Type';

sub description { 'Total list of biggest messages in ALL folders' }

# {{{ report
#
# Report sorted by whatever header...
#
sub report {

    my $self = shift;
    my $args = shift;
    my $opts = $self->opts;

    my $reports = $self->types;

    my $afbmr_cache = $args->{cache};

    my @cur_report;

    my $report_type = $reports->{all_folders_biggest_message_report};

    my @imap_folders = fetch_folders({ cache => $afbmr_cache });

    # Iterate our folders to make sure the cache is current.
    #
    for ( @imap_folders ) {
        my $msg_count = fetch_messages({ cache => $afbmr_cache, folder  => $_ });
    }

    push @cur_report, "\n" x 5;

    my $msgs = cache_report({ cache       => $afbmr_cache,
                              report_type => 'all_folders_biggest_message_report' });


    ddump( 'msgs', $msgs ) if $opts->{debug};

    push @cur_report, $_
        for tabulator({ rows    => $msgs,
                        columns => [qw/FOLDER DATE SIZE TO FROM SUBJECT/] });

    push @cur_report, "\n\n\n";

    return @cur_report;

} # }}}

1;

# }}}

# {{{ package IMAP::Report::Type::all_folders_list_ids_report
#
package IMAP::Report::Type::all_folders_list_ids_report;

use Moose;

with 'IMAP::Report::Type';

sub description { 'Total summary of the message List-ID headers in ALL folders' }

# {{{ report
#
sub report {

    my $self = shift;
    my $args = shift;

    my $opts = $self->opts;

    my $reports = $self->types;

    my $alir_cache = $args->{cache};

    my $report_type = $reports->{all_folders_list_ids_report};

    my @imap_folders = fetch_folders({ cache => $alir_cache });

    # Extra press-enter-to-continue prompts to make damn sure that we understand
    # that this is long, heavy operation on a mailbox with a ton of messages.
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
        . join( "\n", @imap_folders )
        . "\n\nAbove is the list of folders that will be iterated for the report.\n\n"
        . "Last chance to cancel!\n\n"
    );

    my $total_message_count = 0;

    my @listid_report = ();

    my $raw_report         = [];

    for ( @imap_folders ) {
        for my $cur_msg ( get_list_ids({ cache => $alir_cache, folder => $_ }) ) {
            push @$raw_report, [ $cur_msg->[0], $cur_msg->[1], $cur_msg->[2] ];
        }
    }

    if ( ! scalar(@$raw_report) ) {
        show_error( "No message details found..." );
        return;
    }

    push @listid_report, $_
        for tabulator({ rows    => $raw_report,
                        columns => [qw/COUNT FOLDER LISTID/] });


    push @listid_report, "\n\n\n" . '-' x 60 . "\n\n\n\n";

    return @listid_report;

} # }}}

1;

# }}}

# {{{ package IMAP::Report::Type::all_folders_messages_by_subject_report
#
package IMAP::Report::Type::all_folders_messages_by_subject_report;

use Moose;

with 'IMAP::Report::Type';

sub description { 'Total summary of the message Subject headers in ALL folders' }

# {{{ report
#
sub report {

    return;

} # }}}

1;

# }}}

# {{{ package IMAP::Report::Type::messages_by_subject_report
#
package IMAP::Report::Type::messages_by_subject_report;

use Moose;

with 'IMAP::Report::Type';

sub description { 'Folder statistics report for message SUBJECT' }

# {{{ report
#
sub report {

    return;

} # }}}

1;

# }}}

# {{{ package IMAP::Report::Type::messages_by_list_id_report
#
package IMAP::Report::Type::messages_by_list_id_report;

use Moose;

with 'IMAP::Report::Type';

sub description { 'Folder statistics report for message LISTID' }

# {{{ report
#
sub report {

    my $self = shift;
    my $args = shift;

    my $opts = $self->opts;

    my $reports = $self->types;

    my $mblir_cache = $args->{cache};

    return unless $mblir_cache;

    my $report_type = $reports->{messages_by_list_id_report};

    my @imap_folders = fetch_folders({ cache => $mblir_cache });

    my $folder = folder_choice({ folders => \@imap_folders, report_type => $report_type });

    my @listid_report = ();

    my $raw_report         = [];

    for my $cur_msg ( get_list_ids({ cache => $mblir_cache, folder => $folder }) ) {
        push @$raw_report, [ $cur_msg->[0], $cur_msg->[1], $cur_msg->[2] ];
    }

    if ( ! scalar(@$raw_report) ) {
        show_error( "No message details found..." );
        return;
    }

    push @listid_report, $_
        for tabulator({ rows    => $raw_report,
                        columns => [qw/COUNT FOLDER LISTID/] });

    return @listid_report;

} # }}}

1;

# }}}

# {{{ package IMAP::Report::Type::messages_by_from_address_report
#
package IMAP::Report::Type::messages_by_from_address_report;

use Moose;

with 'IMAP::Report::Type';

sub description { 'Folder statistics report for message FROM addresses' }

# {{{ report
#
sub report {

    return;

} # }}}

1;

# }}}

# {{{ package IMAP::Report::Type::messages_by_to_address_report
#
package IMAP::Report::Type::messages_by_to_address_report;

use Moose;

with 'IMAP::Report::Type';

sub description { 'Folder statistics report for message TO addresses' }

# {{{ report
#
sub report {

    return;

} # }}}

1;

# }}}

# {{{ package IMAP::Report::Type::biggest_messages_report
#
package IMAP::Report::Type::biggest_messages_report;

use Moose;

with 'IMAP::Report::Type';

sub description { 'Folder statistics report for message SIZE' }

# {{{ report
#
# This report is to give us a top-ten style report to see the largest messages
# in a folder.
#
# Expects to receive a cache object.
#
# Returns a report in the form of an array
#
sub report {

    my $self = shift;
    my $args = shift;

    my $opts = $self->opts;

    my $reports = $self->types;

    my $bmr_cache = $args->{cache};

    my $report_type = $reports->{biggest_messages_report};

    my @imap_folders = fetch_folders({ cache => $bmr_cache });

    # Now that we know what type of report to run, we need to choose the folder
    # on which we wish to report.
    #
    my $folder = folder_choice({ folders => \@imap_folders, report_type => $report_type });

    return unless $folder;

    my @breport;

    my $stime = time;

    my $fetch_count = fetch_messages({ folder => $folder,
                                       cache  => $bmr_cache });

    if ( ! $fetch_count ) {
        push @breport, "\n\n** No messages on which to report for folder: $folder\n\n";
        return @breport;
    }

    push @breport, "\n\nReporting on $fetch_count messages from folder: $folder\n\n";

    my $fetched_messages = cache_report({ folder      => $folder,
                                          cache       => $bmr_cache,
                                          report_type => 'report_by_size' });

    ddump( 'fetched_messages', $fetched_messages ) if $opts->{debug};

    my $ftime = time;

    my $elapsed = $ftime - $stime;

    push @breport, "\nTotal time to fetch all messages: $elapsed seconds\n";
    push @breport, "\nReporting on the top " . $opts->{top} . " messages.\n";

    my $totalsize;

    # quick calculation on the total size of the messages so
    # it can appear at the top of the report.
    #

    #for ( keys %$fetched_messages ) {
    #    $totalsize += $fetched_messages->{$_}->{$header_table{'Size'}};
    #}

    my @msglist;

    my $reportsize;

    push @breport,
        "\n\n\n"
        . "Top messages, sorted by size...\n"
        . '-' x 60 . "\n\n"
        . "DATE\t\t\t\tSIZE\t\tSUBJECT\n"
        . '-' x 60 . "\n\n";


    for ( @$fetched_messages ) {

        my $folder       = $_->[0];
        my $msg_id       = $_->[1];
        my $TO           = $_->[2];
        my $FROM         = $_->[3];
        my $DATE         = $_->[4];
        my $SUBJECT      = $_->[5];
        my $SIZE         = $_->[6];
        my $LISTID       = $_->[7];


        push @breport, "$DATE\t$SIZE\t\t$SUBJECT\n";

        if ( $opts->{search} ) {

            my $search =
                generate_search_string({ folder => $folder,
                                         date   => $DATE,
                                         header => 'Subject',
                                         value  => $SUBJECT });

            push @breport, $search . "\n\n";

        }

    }


    return @breport;

} # }}}

1;

# }}}

# {{{ package IMAP::Report::Type::size_report
#
package IMAP::Report::Type::size_report;

use Moose;

with 'IMAP::Report::Type';

sub description { 'Folder summary report for total size of messages' }

# {{{ report
#
sub report {

    my $self = shift;
    my $args = shift;

    my $opts = $self->opts;

    my $reports = $self->types;

    my $sr_cache = $args->{cache};

    my $report_type = $reports->{size_report};

    my @sreport;

    my @imap_folders = fetch_folders({ cache => $sr_cache });

    my $folder = folder_choice({ folders     => \@imap_folders,
                                 report_type => $report_type });

    return unless $folder;

    print "\n\nFetching message details for folder '$folder'...\n\n";

    my $stime = time;

    my $msg_count = fetch_messages({ cache => $sr_cache, folder  => $folder });

    if ( ! $msg_count ) {
        push @sreport, "\n\n** No messages on which to report for folder: $folder\n\n";
        return @sreport;
    }

    push @sreport, "\n\nReporting on $msg_count messages from folder: $folder\n\n";

    my $fetched_messages = cache_report({ folder      => $folder,
                                          cache       => $sr_cache,
                                          report_type => 'total_folder_size' });

    my $totalsize = 0;

    $totalsize += $_->[0] for @$fetched_messages;

    my $ftime = time;

    my $elapsed = $ftime - $stime;

    push @sreport, "\nTime to fetch: $elapsed seconds\n";

    push @sreport, '-' x 60 . "\n";
    push @sreport, "\n\nTotal size of all messages in '$folder' = " . convert_bytes($totalsize) . "\n\n";
    push @sreport, '-' x 60 . "\n";

    return @sreport;

} # }}}

1;

# }}}

# {{{ package IMAP::Report::Type::list
#
package IMAP::Report::Type::list;

use Moose;

with 'IMAP::Report::Type';

sub description { 'Display the current list of folders' }

# {{{ report
#
# List the folders directly from the imap server.
#
sub report {

    my $self = shift;

    my $opts = $self->opts;

    my @r;

    push @r, "IMAP server '" . $opts->{server} . "' shows the following folders:\n\n";

    my @imap_folders = imap_folders();

    my @filtered = filter_folders({ folders => \@imap_folders });

    push @r, "$_\n" for @filtered;

    return @r;

} # }}}

1;

# }}}

# }}}

# {{{ package IMAP::Report::Cache
#
package IMAP::Report::Cache;

use Moose;

extends 'IMAP::Report';

has file => (
   is       => 'rw',
   isa      => 'Str',
   required => 1,
);

# {{{ init
#
# Crude implementation of a cache using SQLite
#
sub init {

    my $self = shift;
    my $opts = $self->opts;
    my $file = $self->file;

    main::ddump( 'self_in_init', $self );
    main::ddump( 'opts_in_init', $opts );

    my $cfile = $file;

    unless ( eval 'require DBI; import DBI; require DBD::SQLite; import DBD::SQLite; 1;' ) {
        die_clean( 1, "DBI and DBD::SQLite perl modules required." );
        return;
    }

    my $is_new_db =
        -s $cfile
        ? 0
        : 1
        ;

    my $dsn = "dbi:SQLite:dbname=$cfile";

    my $c;

    my $err;

    my $dbh = DBI->connect( $dsn, '', '' )
        or $err = 1;

    if ( $err ) {
        warn "Unable to init cache db: $!\n";
        return;
    }

    $dbh->do('PRAGMA cache_size = 16384;');
    $dbh->do('PRAGMA page_size = 2048;');
    $dbh->do('PRAGMA temp_store = 2;');
    $dbh->do('PRAGMA synchronous = 0;');

    if ( $is_new_db ) {

        my $sql = q[

            CREATE TABLE folders (
                id              INTEGER NOT NULL PRIMARY KEY,
                server          TEXT NOT NULL,
                username        TEXT NOT NULL,
                folder          TEXT NOT NULL,
                msg_count       INTEGER,
                validated       BOOLEAN,
                last_update     INTEGER,
                    UNIQUE (
                        server,
                        username,
                        folder
                    )
            );
        ];

        my $sth = $dbh->prepare( $sql );

        my $err = $sth->execute;

        $sql = q[

            CREATE TABLE messages (
                id              INTEGER NOT NULL PRIMARY KEY,
                server          TEXT NOT NULL,
                username        TEXT NOT NULL,
                folder          TEXT NOT NULL,
                msg_id          INTEGER NOT NULL,
                "TO"            TEXT,
                "FROM"          TEXT,
                SUBJECT         TEXT,
                LISTID          TEXT,
                DATE            INTEGER NOT NULL,
                SIZE            INTEGER NOT NULL,
                FULLHEADERS     TEXT,
                last_update     INTEGER,
                    UNIQUE (
                        server,
                        username,
                        folder,
                        msg_id
                    )
            );

        ];

        $sth = $dbh->prepare( $sql );

        $err = $sth->execute();

        $sql = q[

            CREATE TABLE fetchlist (
                id              INTEGER NOT NULL PRIMARY KEY,
                server          TEXT NOT NULL,
                username        TEXT NOT NULL,
                folder          TEXT NOT NULL,
                msg_id          INTEGER UNIQUE NOT NULL,
                last_update     INTEGER,
                    UNIQUE (
                        server,
                        username,
                        folder,
                        msg_id
                    )
            );

        ];

        $sth = $dbh->prepare( $sql );

        $err = $sth->execute();

    }

    $dbh->{AutoCommit} = 1;

    return $dbh;

} # }}}

# {{{ check
#
# My crude method of a caching mechanism.
#
# Feed this function the type of content and value for which to look.  Returns
# cached elements based on the different types of information, arrayrefs,
# hashrefs, bools, etc.
#
# TODO
#
# Implement cache aging
#
sub check {

    my $self = shift;

    my $args = shift;

    my $opts = $self->opts;

    my $dbh          = $self->{cache};
    my $content_type = $args->{content_type};
    my $value        = defined $args->{value} ? $args->{value} : '';

    my $cur_time = time;

    if ( $content_type eq 'folder_list' ) {

        # {{{ folder_list cache check

        # Checks the cached list of folders and returns an arrayref list of
        # them.

        cache_prune({ cache        => $dbh,
                      content_type => $content_type });

        # If we're in cache_only mode, instead of grabbing the stored list of
        # folders, we'll create a list of folders from the actual messages
        # stored in the cache.
        #
        my $sql =
            $opts->{cache_only}
            ? q[
                    SELECT DISTINCT
                        folder
                    FROM
                        messages
                    WHERE
                        server = ?
                        AND username = ?
              ]

            : q[

                    SELECT
                        folder
                    FROM
                        folders
                    WHERE
                        server = ?
                        AND username = ?
                        AND validated = 1

               ]
            ;

        my $folderlist = [];

        push @$folderlist, $_->{folder}
            for @{ $dbh->selectall_arrayref( $sql, { Slice => {} },
                                             $opts->{server}, $opts->{user} ) };

        if ( scalar(@$folderlist) ) {
            return $folderlist;
        } else {
            return;
        }

        # }}}

    } elsif ( $content_type eq 'validated_folder_list' ) {

        # {{{ validated folder cache check

        # The list of VALIDATED folders is cached separately.  This just returns
        # a true/false if a folder appears in the list of validated folders.
        # (Validated folders have passed a test using the 'exists' method.)
        #
        return unless defined $value && $value;


        my $sql = q[
            SELECT
                folder
            FROM
                folders
            WHERE
                server = ?
                AND username = ?
                AND validated = 1

        ];

        my $results = [];

        push @$results, $_->{folder}
            for @{ $dbh->selectall_arrayref( $sql, { Slice => {} },
                                             $opts->{server}, $opts->{user} ) };
        return $results->[0];

                # }}}

    } elsif ( $content_type eq 'fetched_messages' ) {

        # {{{ fetched messages cache check

        return unless defined $value && $value;

        cache_prune({ cache        => $dbh,
                      content_type => $content_type });

        my $sql = q[
            SELECT
                count(msg_id)
            FROM
                messages
            WHERE
                server = ?
                AND username = ?
                AND folder = ?
        ];

        my $sth = $dbh->prepare( $sql );

        $sth->execute( $opts->{server}, $opts->{user}, $value );

        my $count = $sth->fetch;

        return
            $count->[0]
            ? $count->[0]
            : 0
            ;

        # }}}

    } elsif ( $content_type eq 'message_count' ) {

        # {{{ cached message count check

        # This looks into the cache of messages and if messages have been
        # cached, returns the count of the number of messages stored there.
        # Folder message counts themselves are not actually cached.
        #

        return unless defined $value && $value;

       #if ( defined $cache->{ $opts->{server} }->{imap_folders}->{fetched_messages}->{$value}->{messages}
       #     && ref $cache->{ $opts->{server} }->{imap_folders}->{fetched_messages}->{$value}->{messages} eq 'HASH' ) {

       #    my $result = scalar( keys %{ $cache->{ $opts->{server} }->{imap_folders}->{fetched_messages}->{$value}->{messages} } );

       #    return $result;

       #}

        # }}}

    } elsif ( $content_type eq 'messages_to_be_fetched' ) {

        # {{{ return the list of message id's that need to be fetched.

        my $folder = $args->{folder};

        my $limit =
            $args->{limit}
            ? $args->{limit}
            : $opts->{max_fetch}
            ;

        my $offset = $args->{offset};

        ddump( 'folder', $folder ) if $opts->{debug};
        ddump( 'limit',  $limit  ) if $opts->{debug};
        ddump( 'offset', $offset ) if $opts->{debug};

        # TODO
        #
        # Add some locking here...

        return unless $folder;
        return unless $limit;
        return unless defined $offset;

        #$dbh->begin_work;

        my $sql = q[
            SELECT
                msg_id
            FROM
                fetchlist
            WHERE
                server = ?
                AND username = ?
                AND folder = ?
                AND last_update IS NULL
            ORDER BY
                msg_id
            LIMIT ?
            OFFSET ?
        ];

       #my $mcount = scalar( keys %$values );

       #my $cbar = IMAP::Report::Progress->new( max    => $mcount,
       #                                length => 10 );

       #$cbar->text('Caching:');
       #$cbar->info('messages');

       #my $counter = 0;

       #$cbar->update( $counter++ );
       #$cbar->write;

        my $folderlist = [];

        push @$folderlist, $_->{msg_id}
            for @{ $dbh->selectall_arrayref( $sql, { Slice => {} }, $opts->{server}, $opts->{user}, $folder, $limit, $offset ) };

        if ( $dbh->errstr ) {
            show_error( 'Messages fetched from fetchlist cache error: ' . $dbh->errstr );
            return;
        }

        #show_error( "messages_to_be_fetched_list: " . Dumper( $folderlist ) );

        if ( scalar(@$folderlist) ) {
            return $folderlist;
        } else {
            return;
        }

        #$dbh->commit;

        # }}}

    } elsif ( $content_type eq 'messages_previously_fetched' ) {

        # {{{ return the list of message id's that have been fetched.

        my $folder = $args->{folder};

        return unless $folder;


        #$dbh->begin_work;

        my $sql = q[
            SELECT
                msg_id
            FROM
                fetchlist
            WHERE
                server = ?
                AND username = ?
                AND folder = ?
                AND last_update NOT NULL
        ];

       #my $mcount = scalar( keys %$values );

       #my $cbar = IMAP::Report::Progress->new( max    => $mcount,
       #                                length => 10 );

       #$cbar->text('Caching:');
       #$cbar->info('messages');

       #my $counter = 0;

       #$cbar->update( $counter++ );
       #$cbar->write;

        my $folderlist = [];

        push @$folderlist, $_
            for @{ $dbh->selectall_arrayref( $sql, {}, $opts->{server}, $opts->{user}, $folder ) };

        if ( $dbh->errstr ) {
            show_error( 'Messages fetched from fetchlist cache error: ' . $dbh->errstr );
            return;
        }

        if ( scalar(@$folderlist) ) {
            return $folderlist;
        } else {
            return;
        }

        #$dbh->commit;

        # }}}

    }

    return;

} # }}}

# {{{ put
#
# Handle inserting the various types of information we want to cache.
#
# Sticks in the current time value so for cache aging purposes later.
#
sub put {

    my $self = shift;
    my $args = shift;

    my $opts = $self->opts;

    my $reports = $self->types;

    my $dbh          = $args->{cache};
    my $content_type = $args->{content_type};
    my $values       = $args->{values};
    my $folder       = $args->{folder};

    my %header_table = (
                         IMAP::Report::Headers::TO->name          => IMAP::Report::Headers::TO->imap_name,
                         IMAP::Report::Headers::FROM->name        => IMAP::Report::Headers::FROM->imap_name,
                         IMAP::Report::Headers::SUBJECT->name     => IMAP::Report::Headers::SUBJECT->imap_name,
                         IMAP::Report::Headers::DATE->name        => IMAP::Report::Headers::DATE->imap_name,
                         IMAP::Report::Headers::SIZE->name        => IMAP::Report::Headers::SIZE->imap_name,
                         IMAP::Report::Headers::LISTID->name      => IMAP::Report::Headers::LISTID->imap_name,
                         IMAP::Report::Headers::FULLHEADERS->name => IMAP::Report::Headers::FULLHEADERS->imap_name,
                       );


    return unless $dbh;
    return unless $content_type;

    if ( $content_type eq 'folder_list' ) {

        # {{{ folder_list cache population

        return unless ref $values eq 'ARRAY';

        $dbh->begin_work;

        my $sql = q[

            INSERT INTO folders (
                server,
                username,
                folder,
                last_update
            ) VALUES (
                ?,
                ?,
                ?,
                ?
            )

        ];

        my $cur_time = time;

        for my $cur_folder (@$values) {

            my $sth = $dbh->prepare($sql);

            $sth->execute( $opts->{server}, $opts->{user}, $cur_folder, $cur_time );

            verbose( "Inserted into DB: " . $opts->{server} . ' ' . $opts->{user} . ' ' . $cur_folder . ' ' . $cur_time . "\n" );

        }

        $dbh->commit;

        # }}}

    } elsif ( $content_type eq 'validated_folder_list' ) {

        # {{{ validated folder list cache population

        return unless $folder;
        return if ref $folder;

        $dbh->begin_work;

        my $sql = q[
            INSERT OR REPLACE INTO folders (
                server,
                username,
                folder,
                validated,
                last_update
            ) VALUES (
                ?,
                ?,
                ?,
                ?,
                ?
            )
        ];

        my $sth = $dbh->prepare($sql);

        my $err;

        $sth->execute( $opts->{server}, $opts->{user}, $folder, 1, time )
            or $err = 1;

        if ( $err ) {
            warn "Error caching validated folder: $!\n";
            $dbh->rollback;
            return;
        }

        $dbh->commit;

        # }}}

    } elsif ( $content_type eq 'fetched_messages' ) {

        # {{{ fetched message cache population

        return unless defined $values && ref $values eq 'HASH';
        return unless $folder;


        #show_error( "PUTTING CACHE FOR FOLDER: $folder " . Dumper( $values ) );

        ddump( 'cache_put_values', $values ) if $opts->{debug};

        $dbh->begin_work;

        my $sql = q[
            INSERT OR REPLACE INTO messages (
                server,
                username,
                msg_id,
                folder,
                "TO",
                "FROM",
                SUBJECT,
                DATE,
                SIZE,
                LISTID,
                FULLHEADERS,
                last_update
            ) VALUES (
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?
            )
        ];

       #my $mcount = scalar( keys %$values );

       #my $cbar = IMAP::Report::Progress->new( max    => $mcount,
       #                                length => 10 );

       #$cbar->text('Caching:');
       #$cbar->info('messages');

       #my $counter = 0;

       #$cbar->update( $counter++ );
       #$cbar->write;

        my $sth = $dbh->prepare($sql);

        my $in_time = time;

        for ( keys %$values ) {
            my $result = $sth->execute(
                $opts->{server},
                $opts->{user},
                $_,
                $folder,
                $values->{$_}->{ $header_table{TO} },
                $values->{$_}->{ $header_table{FROM} },
                $values->{$_}->{ $header_table{SUBJECT} },
                $values->{$_}->{ $header_table{DATE} },
                $values->{$_}->{ $header_table{SIZE} },
                $values->{$_}->{ $header_table{LISTID} },
                $values->{$_}->{ $header_table{FULLHEADERS} },
                $in_time
            );

            if ( $dbh->errstr ) {
                show_error( 'Message cache insert error: ' . $dbh->errstr );
                $dbh->rollback;
                return;
            }

            #if ( ( ( $counter % 10 ) + 1 ) == 10 ) {
            #   $cbar->update( $counter++ );
            #   $cbar->write;
            #}

        }

        $dbh->commit;

        # }}}

    } elsif ( $content_type eq 'unfetched_message_ids_from_folder' ) {

        # {{{ message id's present in the selected folder

        return unless defined $values && ref $values eq 'ARRAY';
        return unless $folder;

        #show_error( "PUTTING CACHE FOR FOLDER: $folder " . Dumper( $values ) );

        ddump( 'cache_put_values', $values ) if $opts->{debug};

        $dbh->begin_work;

        my $sql = q[
            INSERT OR REPLACE INTO fetchlist (
                server,
                username,
                folder,
                msg_id
            ) VALUES (
                ?,
                ?,
                ?,
                ?
            )
        ];

       #my $mcount = scalar( keys %$values );

       #my $cbar = IMAP::Report::Progress->new( max    => $mcount,
       #                                length => 10 );

       #$cbar->text('Caching:');
       #$cbar->info('messages');

       #my $counter = 0;

       #$cbar->update( $counter++ );
       #$cbar->write;

        my $sth = $dbh->prepare($sql);

        for ( @$values ) {

            my $result = $sth->execute(
                $opts->{server},
                $opts->{user},
                $folder,
                $_
            );

            if ( $dbh->errstr ) {
                show_error( 'Message fetchlist cache insert error: ' . $dbh->errstr );
                $dbh->rollback;
                return;
            }

            #if ( ( ( $counter % 10 ) + 1 ) == 10 ) {
            #   $cbar->update( $counter++ );
            #   $cbar->write;
            #}

        }

        $dbh->commit;

        # }}}

    } elsif ( $content_type eq 'update_message_fetch_status' ) {

        # {{{ update the message id's status to indicate that it has been fetched

        return unless defined $values && ref $values eq 'ARRAY';
        return unless $folder;

        my $cur_time = time;

        ddump( 'cache_put_values->update_message_fetch_status', $values ) if $opts->{debug};

        $dbh->begin_work;

        my $sql = q[
            INSERT OR REPLACE INTO fetchlist (
                server,
                username,
                folder,
                msg_id,
                last_update
            ) VALUES (
                ?,
                ?,
                ?,
                ?,
                ?
            )
        ];

       #my $mcount = scalar( keys %$values );

       #my $cbar = IMAP::Report::Progress->new( max    => $mcount,
       #                                length => 10 );

       #$cbar->text('Caching:');
       #$cbar->info('messages');

       #my $counter = 0;

       #$cbar->update( $counter++ );
       #$cbar->write;

        my $sth = $dbh->prepare($sql);

        for ( @$values ) {

            my $result = $sth->execute(
                $opts->{server},
                $opts->{user},
                $folder,
                $_,
                $cur_time
            );

            if ( $dbh->errstr ) {
                show_error( 'Message fetchlist cache insert error: ' . $dbh->errstr );
                $dbh->rollback;
                return;
            }

            #if ( ( ( $counter % 10 ) + 1 ) == 10 ) {
            #   $cbar->update( $counter++ );
            #   $cbar->write;
            #}

        }

        $dbh->commit;

        # }}}

    }

    return;

} # }}}

# {{{ prune
#
sub prune {

    my $self = shift;

    my $opts = $self->opts;

    # TODO
    #
    # FIX.
    #
    my $cache_age = 7 * ( 24 * 60 * 60 );

    # Don't prune in cache_only mode...
    #
    return if $opts->{cache_only};

    # Don't prune if cache pruning is disabled...
    #
    return unless $opts->{cache_prune};

    my $args = shift;

    my $content_type = $args->{content_type};
    my $dbh          = $args->{cache};

    return unless $content_type;
    return unless $dbh;

    my $cur_time = time;
    my $max_age  = $cur_time - $cache_age;

    if ( $content_type eq 'folder_list' ) {

        my $sql = q[
            DELETE FROM
                folders
            WHERE
                server = ?
                AND username = ?
                AND last_update < ?
        ];

        my $sth = $dbh->prepare( $sql );

        my $result = $sth->execute( $opts->{server}, $opts->{user}, $max_age );

    } elsif ( $content_type eq 'fetched_messages' ) {

        my $sql = q[
            DELETE FROM
                messages
            WHERE
                server = ?
                AND username = ?
                AND last_update < ?
        ];

        my $sth = $dbh->prepare( $sql );

        my $result = $sth->execute( $opts->{server}, $opts->{user}, $max_age );

    } elsif ( $content_type eq 'messages_to_be_fetched' ) {

        my $sql = q[
            DELETE FROM
                fetchlist
            WHERE
                server = ?
                AND username = ?
                AND last_update < ?
        ];

        my $sth = $dbh->prepare( $sql );

        my $result = $sth->execute( $opts->{server}, $opts->{user}, $max_age );

    }

    $opts->{cache_prune} = 0;

    return;

}

# }}}

# {{{ report
#
# Here's where we're going to start generating our reports.
#
# Expects to receive an anon hashref contain the cache (dbh) object, type of
# report we want to run, the name of the folder, and the header to be used for
# sorting operations in the reports.
#
sub report {

    my $self = shift;
    my $args = shift;

    my $opts = $self->opts;

    my $reports = $self->types;

    my $dbh          = $args->{cache};
    my $folder       = $args->{folder};
    my $report_type  = $args->{report_type};

    if ( $report_type eq 'report_by_size' ) {

        # {{{ report by size

        return unless $folder;

        my $sql = q[
            SELECT
                folder,
                msg_id,
                "TO",
                "FROM",
                DATE,
                SUBJECT,
                SIZE,
                LISTID
            FROM
                messages
            WHERE
                server = ?
                AND username = ?
                AND folder = ?
            ORDER BY SIZE DESC
            LIMIT ?
        ];



        my @results = @{
            $dbh->selectall_arrayref(
                                      $sql,
                                      {},
                                      $opts->{server},
                                      $opts->{user},
                                      $folder,
                                      $opts->{top} )};


        ddump( 'selectall_results', @results ) if $opts->{debug};

        my $messages = [];

        push @$messages, [
                           $_->[0],                     # folder
                           $_->[1],                     # msg_id
                           $_->[2],                     # to_address
                           $_->[3],                     # from_address
                           scalar localtime $_->[4],    # date
                           $_->[5],                     # subject
                           convert_bytes( $_->[6] )     # size
                         ] for @results;

        ddump( 'header report of collected messages', $messages ) if $opts->{debug};

        return $messages;

        # }}}

    } elsif ( $report_type eq 'report_by_header' ) {

        # {{{ report by header

        return unless $folder;

        my $header = $args->{header};

        if ( ! $header ) {
            show_error( "Invalid header: $header" );
            return;
        }

        if ( ! $folder ) {
            show_error( "Invalid folder: $folder" );
            return;
        }

        # TODO
        #
        # Fix this ridiculousness...
        #
       #my $header_column;

       #if ( $header eq 'TO' ) {
       #    $header_column = 'TO';
       #} elsif ( $header eq 'FROM' ) {
       #    $header_column = 'FROM';
       #} elsif ( $header eq 'DATE' ) {
       #    $header_column = 'DATE';
       #} elsif ( $header eq 'SUBJECT' ) {
       #    $header_column = 'SUBJECT';
       #} elsif ( $header eq 'SIZE' ) {
       #    $header_column = 'SIZE';
       #} else {
       #    $header_column = 'SUBJECT';
       #}

       #ddump( 'header_column', $header_column ) if $opts->{debug};


        my $sql = qq[
            SELECT
                count( $header ) AS count_column,
                "TO",
                "FROM",
                DATE,
                SUBJECT,
                SIZE,
                LISTID
            FROM
                messages
            WHERE
                server = ?
                AND username = ?
                AND folder = ?
            GROUP BY $header
                HAVING count_column >= 1
            ORDER BY count_column DESC
            LIMIT ?
        ];

        ddump( 'report_by_header_sql',    $sql )            if $opts->{debug};
        ddump( 'report_by_header_header', $header )         if $opts->{debug};
        ddump( 'report_by_header_server', $opts->{server} ) if $opts->{debug};
        ddump( 'report_by_header_folder', $folder )         if $opts->{debug};
        ddump( 'report_by_header_top',    $opts->{top} )    if $opts->{debug};

        my @results = @{
            $dbh->selectall_arrayref(
                                      $sql,
                                      {},
                                      $opts->{server},
                                      $opts->{user},
                                      $folder,
                                      $opts->{top}
                                    )
            };

        ddump( 'selectall_results', @results ) if $opts->{debug};

        my $messages = [];

        push @$messages, [
                           $_->[0],                     # count_column
                           $_->[1],                     # to_address
                           $_->[2],                     # from_address
                           scalar localtime $_->[3],    # date
                           $_->[4],                     # subject
                           convert_bytes( $_->[5] ),    # size
                         ] for @results;

        ddump( 'header report of collected messages', $messages ) if $opts->{debug};

        return $messages;

        # }}}

    } elsif ( $report_type eq 'all_folders_report_by_header' ) {

        # {{{ all folders report by header

        my $header = $args->{header};

        if ( ! $header ) {
            show_error( "Invalid header: $header" );
            return;
        }

        # TODO
        #
        # Fix this ridiculousness...
        #
       #my $header_column;

       #if ( $header eq 'TO' ) {
       #    $header_column = 'TO';
       #} elsif ( $header eq 'FROM' ) {
       #    $header_column = 'FROM';
       #} elsif ( $header eq 'DATE' ) {
       #    $header_column = 'DATE';
       #} elsif ( $header eq 'SUBJECT' ) {
       #    $header_column = 'SUBJECT';
       #} elsif ( $header eq 'SIZE' ) {
       #    $header_column = 'SIZE';
       #} else {
       #    $header_column = 'SUBJECT';
       #}

       #ddump( 'header_column', $header_column ) if $opts->{debug};


        my $sql = qq[
            SELECT
                count( $header ) AS count_column,
                folder,
                DATE,
                SIZE
                "TO",
                "FROM",
                SUBJECT
            FROM
                messages
            WHERE
                server = ?
                AND username = ?
            GROUP BY $header
                HAVING count_column >= 1
            ORDER BY count_column DESC
            LIMIT ?
        ];

        ddump( 'report_by_header_sql',    $sql )            if $opts->{debug};
        ddump( 'report_by_header_header', $header )         if $opts->{debug};
        ddump( 'report_by_header_server', $opts->{server} ) if $opts->{debug};
        ddump( 'report_by_header_folder', $folder )         if $opts->{debug};
        ddump( 'report_by_header_top',    $opts->{top} )    if $opts->{debug};

        my @results = @{
            $dbh->selectall_arrayref(
                                      $sql,
                                      {},
                                      $opts->{server},
                                      $opts->{user},
                                      $opts->{top}
                                    )
            };

        ddump( 'selectall_results', @results ) if $opts->{debug};

        my $messages = [];

        push @$messages, [
                           $_->[0],                     # count_column
                           $_->[1],                     # folder
                           scalar localtime $_->[2],    # date
                           convert_bytes( $_->[3] ),    # size
                           $_->[4],                     # to_address
                           $_->[5],                     # from_address
                           $_->[6],                     # subject
                         ] for @results;

        ddump( 'header report of collected messages', $messages ) if $opts->{debug};

        return $messages;

        # }}}

    } elsif ( $report_type eq 'total_folder_size' ) {

        # {{{ Total size of messages in folder

        if ( ! $folder ) {
            show_error( "Invalid folder: $folder" );
            return;
        }

        my $sql = qq[
            SELECT
                SIZE
            FROM
                messages
            WHERE
                server = ?
                AND username = ?
                AND folder = ?
        ];

        my @results = @{
            $dbh->selectall_arrayref(
                                      $sql,
                                      {},
                                      $opts->{server},
                                      $opts->{user},
                                      $folder
                                    )
            };

        ddump( 'selectall_results', @results ) if $opts->{debug};

        my $messages = [];

        push @$messages, [ $_->[0] ] for @results;

        ddump( 'size report of collected messages', $messages ) if $opts->{debug};

        return $messages;

        # }}}

    } elsif ( $report_type eq 'all_folders_biggest_message_report' ) {

        # {{{ Biggest messages in all folders

        my $sql = q[
            SELECT
                folder,
                DATE,
                SIZE,
                "TO",
                "FROM",
                SUBJECT
            FROM
                messages
            WHERE
                server = ?
                AND username = ?
            ORDER BY
                SIZE DESC
            LIMIT ?
        ];

        my @results = @{
            $dbh->selectall_arrayref(
                                      $sql,
                                      {},
                                      $opts->{server},
                                      $opts->{user},
                                      $opts->{top}
                                    )
            };

        ddump( 'selectall_results', @results ) if $opts->{debug};

        my $messages = [];

        push @$messages, [
                           $_->[0],                     # folder
                           scalar localtime $_->[1],    # date
                           convert_bytes($_->[2]),      # size
                           $_->[3],                     # to
                           $_->[4],                     # from
                           $_->[5],                     # subject
                         ] for @results;

        return $messages;

        # }}}

    } elsif ( $report_type eq 'all_list_ids' ) {

        # {{{ Total summary of list IDs of all messages in a folder

        if ( ! $folder ) {
            show_error( "Invalid folder: $folder" );
            return;
        }

        my $sql = qq[
            SELECT
                count( LISTID ) AS count_column,
                LISTID
            FROM
                messages
            WHERE
                server = ?
                AND username = ?
                AND folder = ?
            GROUP BY LISTID
                HAVING count_column >= 1
            ORDER BY count_column DESC

        ];

        my @results = @{
            $dbh->selectall_arrayref(
                                      $sql,
                                      {},
                                      $opts->{server},
                                      $opts->{user},
                                      $folder
                                    )
            };

        ddump( 'all_list_ids_selectall_results', @results ) if $opts->{debug};

        my $messages = [];

        push @$messages, [ $_->[0], $_->[1] ] for @results;

        ddump( 'listid report of collected messages', $messages ) if $opts->{debug};

        return $messages;

        # }}}

    } elsif ( $report_type eq 'folder_list' ) {

        # {{{ folder_list cache check

        # Checks the cached list of folders and returns an arrayref list of
        # them.


        my $sql = q[

            SELECT
                folder
            FROM
                folders
            WHERE
                server = ?
                AND username = ?

        ];

        my $folderlist = [];

        push @$folderlist, $_->{folder}
            for @{ $dbh->selectall_arrayref( $sql, { Slice => {} },
                                             $opts->{server}, $opts->{user} ) };

        if ( scalar(@$folderlist) ) {
            return $folderlist;
        } else {
            return;
        }

        # }}}

    } elsif ( $report_type eq 'validated_folder_list' ) {

        # {{{ validated folder cache check

        # The list of VALIDATED folders is cached separately.  This just returns
        # a true/false if a folder appears in the list of validated folders.
        # (Validated folders have passed a test using the 'exists' method.)
        #

        my $sql = q[
            SELECT
                folder
            FROM
                folders
            WHERE
                server = ?
                AND username = ?
                AND validated = 1

        ];

        my $results = [];

        push @$results, $_->{folder}
            for @{ $dbh->selectall_arrayref( $sql, { Slice => {} },
                                             $opts->{server}, $opts->{user} ) };

        return $results->[0];

        # }}}

    } elsif ( $report_type eq 'fetched_messages' ) {

        # {{{ fetched messages cache check

        return unless $folder;

        my $sql = q[
            SELECT
                msg_id,
                "TO",
                "FROM",
                SUBJECT,
                DATE,
                SIZE,
                LISTID
            FROM
                messages
            WHERE
                folder = ?
        ];

        my $msgs = {};

       #for ( @{ $dbh->selectall_arrayref( $sql, { Slice => {} }, $value ) } ) {

       #    $msgs->{$_->{msg_id}}->{$header_table{'FROM'}}    = $_->{from_address};
       #    $msgs->{$_->{msg_id}}->{$header_table{'DATE'}}    = $_->{date};
       #    $msgs->{$_->{msg_id}}->{$header_table{'SUBJECT'}} = $_->{subject};
       #    $msgs->{$_->{msg_id}}->{$header_table{'TO'}}      = $_->{to_address};
       #    $msgs->{$_->{msg_id}}->{$header_table{'SIZE'}}    = $_->{size};

       #}

        return $msgs;

        # }}}

    } elsif ( $report_type eq 'message_count' ) {

        # {{{ cached message count check

        # This looks into the cache of messages and if messages have been
        # cached, returns the count of the number of messages stored there.
        # Folder message counts themselves are not actually cached.
        #

        return unless $folder;


        # }}}

    }

    return;

} # }}}

# }}}

# {{{ package Mail::Address

# Copyrights 1995-2011 by Mark Overmeer <perl@overmeer.net>.
#  For other contributors see ChangeLog.
# See the manual pages for details on the licensing terms.
# Pod stripped from pm file by OODoc 2.00.
package Mail::Address;
use vars '$VERSION';
$VERSION = '2.08';

use strict;

use Carp;

# use locale;   removed in version 1.78, because it causes taint problems

sub Version { our $VERSION }



# given a comment, attempt to extract a person's name
sub _extract_name
{   # This function can be called as method as well
    my $self = @_ && ref $_[0] ? shift : undef;

    local $_ = shift
        or return '';

    # Using encodings, too hard. See Mail::Message::Field::Full.
    return '' if m/\=\?.*?\?\=/;

    # trim whitespace
    s/^\s+//;
    s/\s+$//;
    s/\s+/ /;

    # Disregard numeric names (e.g. 123456.1234@compuserve.com)
    return "" if /^[\d ]+$/;

    s/^\((.*)\)$/$1/; # remove outermost parenthesis
    s/^"(.*)"$/$1/;   # remove outer quotation marks
    s/\(.*?\)//g;     # remove minimal embedded comments
    s/\\//g;          # remove all escapes
    s/^"(.*)"$/$1/;   # remove internal quotation marks
    s/^([^\s]+) ?, ?(.*)$/$2 $1/; # reverse "Last, First M." if applicable
    s/,.*//;

    # Change casing only when the name contains only upper or only
    # lower cased characters.
    unless( m/[A-Z]/ && m/[a-z]/ )
    {   # Set the case of the name to first char upper rest lower
        s/\b(\w+)/\L\u$1/igo;  # Upcase first letter on name
        s/\bMc(\w)/Mc\u$1/igo; # Scottish names such as 'McLeod'
        s/\bo'(\w)/O'\u$1/igo; # Irish names such as 'O'Malley, O'Reilly'
        s/\b(x*(ix)?v*(iv)?i*)\b/\U$1/igo; # Roman numerals, eg 'Level III Support'
    }

    # some cleanup
    s/\[[^\]]*\]//g;
    s/(^[\s'"]+|[\s'"]+$)//g;
    s/\s{2,}/ /g;

    $_;
}

sub _tokenise
{   local $_ = join ',', @_;
    my (@words,$snippet,$field);

    s/\A\s+//;
    s/[\r\n]+/ /g;

    while ($_ ne '')
    {   $field = '';
        if(s/^\s*\(/(/ )    # (...)
        {   my $depth = 0;

     PAREN: while(s/^(\(([^\(\)\\]|\\.)*)//)
            {   $field .= $1;
                $depth++;
                while(s/^(([^\(\)\\]|\\.)*\)\s*)//)
                {   $field .= $1;
                    last PAREN unless --$depth;
	            $field .= $1 if s/^(([^\(\)\\]|\\.)+)//;
                }
            }

            carp "Unmatched () '$field' '$_'"
                if $depth;

            $field =~ s/\s+\Z//;
            push @words, $field;

            next;
        }

        if( s/^("(?:[^"\\]+|\\.)*")\s*//       # "..."
         || s/^(\[(?:[^\]\\]+|\\.)*\])\s*//    # [...]
         || s/^([^\s()<>\@,;:\\".[\]]+)\s*//
         || s/^([()<>\@,;:\\".[\]])\s*//
          )
        {   push @words, $1;
            next;
        }

        croak "Unrecognised line: $_";
    }

    push @words, ",";
    \@words;
}

sub _find_next
{   my ($idx, $tokens, $len) = @_;

    while($idx < $len)
    {   my $c = $tokens->[$idx];
        return $c if $c eq ',' || $c eq ';' || $c eq '<';
        $idx++;
    }

    "";
}

sub _complete
{   my ($class, $phrase, $address, $comment) = @_;

    @$phrase || @$comment || @$address
       or return undef;

    my $o = $class->new(join(" ",@$phrase), join("",@$address), join(" ",@$comment));
    @$phrase = @$address = @$comment = ();
    $o;
}


sub new(@)
{   my $class = shift;
    bless [@_], $class;
}


sub parse(@)
{   my $class = shift;
    my @line  = grep {defined} @_;
    my $line  = join '', @line;

    my (@phrase, @comment, @address, @objs);
    my ($depth, $idx) = (0, 0);

    my $tokens  = _tokenise @line;
    my $len     = @$tokens;
    my $next    = _find_next $idx, $tokens, $len;

    local $_;
    for(my $idx = 0; $idx < $len; $idx++)
    {   $_ = $tokens->[$idx];

        if(substr($_,0,1) eq '(') { push @comment, $_ }
        elsif($_ eq '<')    { $depth++ }
        elsif($_ eq '>')    { $depth-- if $depth }
        elsif($_ eq ',' || $_ eq ';')
        {   warn "Unmatched '<>' in $line" if($depth);
            my $o = $class->_complete(\@phrase, \@address, \@comment);
            push @objs, $o if defined $o;
            $depth = 0;
            $next = _find_next $idx+1, $tokens, $len;
        }
        elsif($depth)       { push @address, $_ }
        elsif($next eq "<") { push @phrase,  $_ }
        elsif( /^[.\@:;]$/ || !@address || $address[-1] =~ /^[.\@:;]$/ )
        {   push @address, $_ }
        else
        {   warn "Unmatched '<>' in $line" if $depth;
            my $o = $class->_complete(\@phrase, \@address, \@comment);
            push @objs, $o if defined $o;
            $depth = 0;
            push @address, $_;
        }
    }
    @objs;
}


sub phrase  { shift->set_or_get(0, @_) }
sub address { shift->set_or_get(1, @_) }
sub comment { shift->set_or_get(2, @_) }

sub set_or_get($)
{   my ($self, $i) = (shift, shift);
    @_ or return $self->[$i];

    my $val = $self->[$i];
    $self->[$i] = shift if @_;
    $val;
}


my $atext = '[\-\w !#$%&\'*+/=?^`{|}~]';
sub format
{   my @addrs;

    foreach (@_)
    {   my ($phrase, $email, $comment) = @$_;
        my @addr;

        if(defined $phrase && length $phrase)
        {   push @addr
              , $phrase =~ /^(?:\s*$atext\s*)+$/o ? $phrase
              : $phrase =~ /(?<!\\)"/             ? $phrase
              :                                    qq("$phrase");

            push @addr, "<$email>"
                if defined $email && length $email;
        }
        elsif(defined $email && length $email)
        {   push @addr, $email;
        }

        if(defined $comment && $comment =~ /\S/)
        {   $comment =~ s/^\s*\(?/(/;
            $comment =~ s/\)?\s*$/)/;
        }

        push @addr, $comment
            if defined $comment && length $comment;

        push @addrs, join(" ", @addr)
            if @addr;
    }

    join ", ", @addrs;
}


sub name
{   my $self   = shift;
    my $phrase = $self->phrase;
    my $addr   = $self->address;

    $phrase    = $self->comment
        unless defined $phrase && length $phrase;

    my $name   = $self->_extract_name($phrase);

    # first.last@domain address
    if($name eq '' && $addr =~ /([^\%\.\@_]+([\._][^\%\.\@_]+)+)[\@\%]/)
    {   ($name  = $1) =~ s/[\._]+/ /g;
	$name   = _extract_name $name;
    }

    if($name eq '' && $addr =~ m#/g=#i)    # X400 style address
    {   my ($f) = $addr =~ m#g=([^/]*)#i;
	my ($l) = $addr =~ m#s=([^/]*)#i;
	$name   = _extract_name "$f $l";
    }

    length $name ? $name : undef;
}


sub host
{   my $addr = shift->address || '';
    my $i    = rindex $addr, '@';
    $i >= 0 ? substr($addr, $i+1) : undef;
}


sub user
{   my $addr = shift->address || '';
    my $i    = index $addr, '@';
    $i >= 0 ? substr($addr,0,$i) : $addr;
}

1;

# }}}

# {{{ package main
#
package main;

use strict;
use warnings;

use Data::Dumper;
use Getopt::Long qw/:config auto_help auto_version/;

use Mail::IMAPClient;
use Term::ReadKey qw/GetTerminalSize/;
use Term::Menus;

our $VERSION = sprintf "%d.%d", q$Revision: 1.1 $ =~ /(\d+)/g;

$|=1;

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
$opts->{Maxcommandlength}  = 1_000;   # size of a single fetch operation
$opts->{Keepalive}         = 0;
$opts->{Fast_io}           = 1;
$opts->{Reconnectretry}    = 3;
$opts->{Ssl}               = 1;
$opts->{Uid}               = 0;
$opts->{list}              = 0;
$opts->{types}             = 0;
$opts->{threads}           = 0;
$opts->{use_threaded_mode} = 0;
$opts->{min_for_threads}   = 200;     # Minimum number of messages before threaded mode allowed
$opts->{threshold}         = 20;      # Message count threshold before flushing message cache
$opts->{conf}              = "$ENV{HOME}/.imapreportrc";
$opts->{cache_file}        = "$ENV{HOME}/.imap-report.cache";
$opts->{cache_age}         = 7;
$opts->{cache_only}        = 0;
$opts->{cache_prune}       = 1;
$opts->{max_fetch}         = 500;     # Max num messages in a single fetch operation.
$opts->{search}            = 0;       # Generate gmail search strings for top lists.
$opts->{quote_headers}     = 0;

GetOptions(

    $opts,

        'server=s',
        'port=i',
        'user=s',
        'password=s',
        'top=i',
        'filters|folder|search=s@{,}',
        'exclude=s@{,}',
        'report=s',
        'types!',
        'list!',
        'search!',
        'log=s',
        'conf=s',
        'Keepalive!',
        'Fast_io!',
        'Maxcommandlength=i',
        'Reconnectretry=i',
        'Ssl!',
        'Uid!',
        'cache_file=s',
        'cache_age=i',
        'cache_only!',
        'cache_prune!',
        'max_fetch=i',
        'threshold=s',
        'threads=i',
        'use_threaded_mode!',
        'min_for_threads=i',
        'pager=s',
        'quote_headers!',
        'debug!',
        'verbose!',

);

read_config_file();

our $cache_age =
    $opts->{cache_age}
    ? $opts->{cache_age} * 24 * 60 * 60
    : 24 * 60 * 60
    ;

if ( $opts->{threshold} =~ m/^(\d+)%/ ) {
    $opts->{threshold_percentage} = $1 * 0.10;
}

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


die "--cache_file option required.\n" unless $opts->{cache_file};
die "--server option required.\n"     unless $opts->{server};
die "--port option required.\n"       unless $opts->{port};

# Default to unthreaded. Only turn on threading if our conditions are met...
#
my $use_threaded_mode = 0;

if ( $opts->{use_threaded_mode} && $opts->{threads} >= 2 ) {

    if (
        eval
            q[
                require threads;
                import threads ( 'exit' => 'threads_only' );
                require threads::shared;
                import threads::shared qw(shared);
                require Thread::Queue;
                import Thread::Queue ();
                1;
            ]) {

        verbose( "Threads detected..." );

        $use_threaded_mode = 1;

    } else {

        verbose( "No threads detected..." );

    }
}

# Date::Manip is now mandatory...
#
unless ( eval 'require Date::Manip; import Date::Manip::Date; 1;' ) {
    die_clean(  1, "Date::Manip module needed..." );
}


# Prompt for user/passwd if necessary
#
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

              top: $opts->{top}
 Maxcommandlength: $opts->{Maxcommandlength}
          threads: $opts->{threads}
use_threaded_mode: $opts->{use_threaded_mode}

       cache_file: $opts->{cache_file}
        cache_age: $opts->{cache_age}

          verbose: $opts->{verbose}

]);


# Set our global imap options here, so we can append to them individually later
# as needed.
#
my %global_imap_options = (
    Server           => $opts->{server},
    Port             => $opts->{port},
    User             => $opts->{user},
    Password         => $opts->{password},
    Keepalive        => $opts->{Keepalive},
    Fast_io          => $opts->{Fast_io},
    Reconnectretry   => $opts->{Reconnectretry},
    Maxcommandlength => $opts->{Maxcommandlength},
    Uid              => $opts->{Uid},
    Clear            => 100,
    Buffer           => 16384,
    Peek             => 1,
);

# Taking a more manual approach to socket creation and corresponding M::I object
# creation....
#
my %ssl_socket_close_options = (
    SSL_no_shutdown => 1,
    SSL_ctx_free    => 1,
);

my %imap_options = %global_imap_options;

if ( $opts->{debug} ) {

    open( DBG, '>>' . $opts->{log} )
        or die_clean( 1, "Unable to open debuglog: $!\n" );

    $imap_options{Debug}    = $opts->{debug} . '.main';
    $imap_options{Debug_fh} = *DBG;

}


# {{{ Perform a quick authentication test
#
if ( ! $opts->{cache_only} ) {

    print "Using IMAP Server: "
        . $opts->{server}
        . ':'
        . $opts->{port}
        . "\n"
        . "Connecting...\n"
        ;

        # Do a quick imap login to make sure we have good credentials.
        #
        my $imap_socket =
            $opts->{Ssl}
            ? create_ssl_socket( 'first_imap_connection_socket' )
            : 0
            ;

        if ( $imap_socket ) {
            $imap_options{Socket} = $imap_socket;
        }

        my $imap = Mail::IMAPClient->new(%imap_options)
            or die "Cannot connect to host : $@";

        if ( $imap->IsAuthenticated ) {
            print "Login successful.\n";
        } else {
            die_clean( 1, "Login failed: $!" );
        }

        $imap->disconnect;

        if ( $opts->{Ssl} ) {
            $imap_socket->close( %ssl_socket_close_options );
        }


} # }}}


# }}}

my $ir = IMAP::Report->new({ opts                     => $opts,
                             imap_options             => \%imap_options,
                             ssl_socket_close_options => \%ssl_socket_close_options });

ddump( 'ir', $ir );

my %header_table = $ir->headers();

ddump( 'header_table', \%header_table );


# Some global queues used in threaded fetching mode...
#
our ( $progress_queue, $fetched_queue, $sequence_queue,
      $sequences_finished_queue, $thread_errors_queue,
      $cache_put_status_queue, $cached_msg_count_queue,
      $idle_thread_queue );

my %thread_pumpkins;

my $break = 0;

# Handle signals gracefully
#
#$SIG{'INT'}  = 'die_signal';
$SIG{'QUIT'} = 'die_signal';
$SIG{'USR1'} = 'die_signal';
#$SIG{'CHLD'} = 'IGNORE';
$SIG{'ABRT'} = 'die_signal';
#$SIG{'SEGV'} = 'die_signal';
#$SIG{'CHLD'} = sub { print "\n\n!!CHLD SIG!!\n\n"; };


# Gracefully terminate application on ^C or command line 'kill'
$SIG{'INT'} = $SIG{'TERM'} =
    sub {
        print(">>> Terminating <<<\n");
       #$TERM = 1;
        # Add -1 to head of idle queue to signal termination
       #$IDLE_QUEUE->insert(0, -1);

        if ( scalar( keys %thread_pumpkins ) ) {
            for ( keys %thread_pumpkins ) {
                $thread_pumpkins{$_}->enqueue(-1);
            }
        }

        die_clean( 1, 'Die on signal...' );

    };



my $reports = $ir->types();

ddump( 'reports', $reports );

ddump( 'ir_opts', $ir->opts );

my $cache = $ir->cache_init;

ddump( 'ir', $ir );
ddump( 'cache', $cache );


# {{{ The main loop...
#
# Keep going until 'quit'
#
while (1) {

    if ($use_threaded_mode) {
        $fetched_queue            = Thread::Queue->new();
        $progress_queue           = Thread::Queue->new();
        $sequence_queue           = Thread::Queue->new();
        $sequences_finished_queue = Thread::Queue->new();
        $thread_errors_queue      = Thread::Queue->new();
        $idle_thread_queue        = Thread::Queue->new();
        $cache_put_status_queue   = Thread::Queue->new();
        $cached_msg_count_queue   = Thread::Queue->new();
    }

    my $banner = "Choose your report type: \n\n\n";


    # Choose what type of report we want to run.
    #
    my $action =
        $opts->{report} && $reports->{$opts->{report}}
        ? $reports->{$opts->{report}}
        : choose_action({ banner  => $banner,
                          reports => $reports });

    die_clean( 0, 'Quitting...' ) if $action eq ']quit[';

    print "\n\n Report selected: $action\n\n";

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

        @report = biggest_messages_report({ cache => $cache });

    } elsif ( $action eq $reports->{size_report} ) {

        @report = size_report({ cache => $cache });

    } elsif ( $action eq $reports->{all_folders_message_count_report} ) {

        @report = all_folders_message_count_report({ cache => $cache });

    } elsif ( $action eq $reports->{all_folders_biggest_message_report} ) {

        @report = all_folders_biggest_message_report({ cache => $cache });

    } elsif ( $action eq $reports->{all_folders_message_sizes_report} ) {

        @report = all_folders_message_sizes_report({ cache => $cache });

    } elsif ( $action eq $reports->{all_folders_list_ids_report} ) {

        @report = all_folders_list_ids_report({ cache => $cache });

    } elsif ( $action eq $reports->{all_folders_messages_by_subject_report} ) {

        @report = all_folders_report_by_header({ cache => $cache, header => 'SUBJECT' });

    } elsif ( $action eq $reports->{messages_by_subject_report} ) {

        @report = messages_by_header_report({ cache => $cache, header => 'SUBJECT' });

    } elsif ( $action eq $reports->{messages_by_from_address_report} ) {

        #@report = messages_by_from_address_report({ cache => $cache });

        @report = messages_by_header_report({ cache => $cache, header => 'FROM' });

    } elsif ( $action eq $reports->{messages_by_to_address_report} ) {

        #@report = messages_by_to_address_report();

        @report = messages_by_header_report({ cache => $cache, header => 'TO' });

    } elsif ( $action eq $reports->{messages_by_list_id_report} ) {

        @report = messages_by_list_id_report({ cache => $cache });

    } elsif ( $action eq $reports->{list} ) {

        @report = list_folders({ cache => $cache });

    } else {

        die_clean( 1, "Invalid report type selected: $action" );

    }

    next unless scalar(@report);

    my $ts = scalar localtime time;

    unshift @report,
                "\n\n\n\n\n"
                . '-' x 60 . "\n"
                . "Report type: $action\n"
                . '-' x 60 . "\n"
                . "Report timestamp: $ts\n"
                . '-' x 60
                . "\n\n\n\n\n";


    push @report,
        "\n\n\n\n\n"
        . '-' x 60 . "\n"
        . "\n\n\n\n\n";


    print_report(\@report);


    # If we manually specified the report type, don't bother going back to the
    # menu.
    #
    die_clean( 0, "Quitting..." )  if $opts->{report};

} # }}}

# {{{ subs
#

# {{{ Report generators

# {{{ all_folders_report_by_header
#
# Report sorted by whatever header...
#
sub all_folders_report_by_header {

    my $args = shift;

    my $header      = $args->{header};
    my $afrbh_cache = $args->{cache};

    return unless $header;

    return unless grep $header eq $_, keys %header_table;

    my @cur_report;

    my $report_type;

    # TODO
    #
    # YUCK.
    #
    if ( $header eq 'SUBJECT' ) {
        $report_type = $reports->{all_folders_messages_by_subject_report};
    } elsif ( $header eq 'TO' ) {
        $report_type = $reports->{messages_by_to_address_report};
    } else {

        # This will be the default header report type...
        #
        $report_type = $reports->{messages_by_from_address_report};

    }

    my @imap_folders = fetch_folders({ cache => $afrbh_cache });

    my $folder = folder_choice({ folders => \@imap_folders, report_type => $report_type });

    return unless $folder;

    my $start_time = time;

    my $num = fetch_messages({ folder => $folder,
                               cache  => $afrbh_cache });


    my $finish_time = time;

    my $elapsed = convert_seconds( $finish_time - $start_time );

    if ( ! $num ) {
        push @cur_report, "\n\n** No messages on which to report for folder: $folder\n\n";
        return @cur_report;
    }


    push @cur_report, "\n" x 10;

    push @cur_report, "Finished fetching messages for folder '$folder'\n\n";
    push @cur_report, "Total time to fetch messages: $elapsed\n\n";
    push @cur_report, "Total messages processed: $num\n\n";
    push @cur_report, 'Top ' . $opts->{top} . " $header instances\n\n\n";


    my $msgs = cache_report({ folder      => $folder,
                              header      => $header,
                              cache       => $afrbh_cache,
                              report_type => 'all_folders_report_by_header' });


    ddump( 'msgs', $msgs ) if $opts->{debug};

    push @cur_report, $_
        for tabulator({ rows    => $msgs,
                        header  => $header,
                        columns => [qw/COUNT FOLDER DATE SIZE TO FROM SUBJECT/] });

    push @cur_report, "\n\n\n";

    return @cur_report;

} # }}}

# {{{ messages_by_header_report
#
# Report sorted by whatever header...
#
sub messages_by_header_report {

    my $args = shift;

    my $header     = $args->{header};
    my $mbhr_cache = $args->{cache};

    return unless $header;

    return unless grep $header eq $_, keys %header_table;

    my @cur_report;

    my $report_type;

    # TODO
    #
    # YUCK.
    #
    if ( $header eq 'SUBJECT' ) {
        $report_type = $reports->{messages_by_subject_report};
    } elsif ( $header eq 'TO' ) {
        $report_type = $reports->{messages_by_to_address_report};
    } else {

        # This will be the default header report type...
        #
        $report_type = $reports->{messages_by_from_address_report};

    }

    my @imap_folders = fetch_folders({ cache => $mbhr_cache });

    my $folder = folder_choice({ folders => \@imap_folders, report_type => $report_type });

    return unless $folder;

    my $start_time = time;

    my $num = fetch_messages({ folder => $folder,
                               cache  => $mbhr_cache });


    my $finish_time = time;

    my $elapsed = convert_seconds( $finish_time - $start_time );

    if ( ! $num ) {
        push @cur_report, "\n\n** No messages on which to report for folder: $folder\n\n";
        return @cur_report;
    }


    push @cur_report, "\n" x 10;

    push @cur_report, "Finished fetching messages for folder '$folder'\n\n";
    push @cur_report, "Total time to fetch messages: $elapsed\n\n";
    push @cur_report, "Total messages processed: $num\n\n";
    push @cur_report, 'Top ' . $opts->{top} . " $header instances\n\n\n";


    my $msgs = cache_report({ folder      => $folder,
                              header      => $header,
                              cache       => $mbhr_cache,
                              report_type => 'report_by_header' });


    ddump( 'msgs', $msgs ) if $opts->{debug};

    push @cur_report, $_
        for tabulator({ rows    => $msgs,
                        header  => $header,
                        columns => [qw/COUNT TO FROM DATE SUBJECT SIZE LISTID/] });

    push @cur_report, "\n\n\n";

    return @cur_report;

} # }}}

# }}}

# {{{ IMAP operations

# {{{ validate_folders
#
# Receives an arrayref of folder names.  Runs the M::I->exists() method on them
# to verify they really exist. Important because of the curious way gmail
# handles nested labels.
#
sub validate_folders {

    my $args = shift;

    my $folders  = $args->{folders};

    return unless $folders;

    my $vf_socket = 
        $opts->{Ssl}
        ? create_ssl_socket( 'vf_socket' )
        : 0
        ;

    if ( $vf_socket ) {
        $imap_options{Socket} = $vf_socket;
    }

    my $vf_imap = Mail::IMAPClient->new( %imap_options )
        or die "Cannot connect to host : $@";


    # Run the exists method on the folder to compensate for the odd behavior
    # that can come from using nested gmail labels.
    #
    my $max = scalar(@$folders);

    my $bar = IMAP::Report::Progress->new( max => $max, length => 10 );
    my $counter= 1;

    my @validated_folders;

    for my $folder ( @$folders ) {

        $bar->info( "Sent: " . $vf_imap->Transaction . " STATUS \"$folder\"" );

        if ( $vf_imap->exists($folder) ) {
            push @validated_folders, $folder;
        }

        $bar->update( $counter++ );
        $bar->write;

    }

    $vf_imap->disconnect;

    if ( $opts->{Ssl} ) {
        $vf_socket->close( %ssl_socket_close_options );
    }


    return @validated_folders;

}

# }}}

# {{{ threaded_fetch_msgs
#
# Expects to receive a list of items representing the message attributes we want
# to fetch.
#
# Returns a hashref of message id's as the keys, and the values for each key are
# hashrefs of the message attributes on which we want to report.
#
sub threaded_fetch_msgs {

    my $args = shift;

    my $folder         = $args->{folder};
    my $tfm_cache      = $args->{cache};
    my $imap_msg_count = $args->{message_count};

    return unless $folder;
    return unless $tfm_cache;

    my $fetcher = {};

    my $total_sequences     = 0;
    my $sequences_completed = 0;

    # Take the list of message id's and break them up into smaller chunks in the
    # form of an array of MessageSet objects.
    #
   #my $threaded_sequences = threaded_sequence_chunker( $imap_msg_ids );

    # Call the function that takes our message id list and loads up our
    # sequences queue with M::I::MessageSet objects.
    #
   #threaded_sequence_chunker( $imap_msg_ids );





    my $offset = 0;

    # Iterate through the message ids creating blocks of --max_fetch in size.
    # Turn each block of message ids into a M::I::MessageSet object and stuff it
    # into the queue of sequences.
    #
    while ( $offset <= $imap_msg_count ) {

        my $cur_block = cache_check(
                                    {
                                      cache        => $tfm_cache,
                                      content_type => 'messages_to_be_fetched',
                                      folder       => $folder,
                                      limit        => $opts->{max_fetch},
                                      offset       => $offset
                                    }
                                   );

        ddump( 'cur_block_from_cache', $cur_block ) if $opts->{debug};

        unless ( $cur_block && scalar(@$cur_block) > 0 ) {
            verbose("\n\n\n\n\nNo messages remaining to be fetched...\n\n\n\n");
            last;
        }

        # Make a M::I::MS object to get a clean range of message ids.
        #
        my $cur_msgset  = Mail::IMAPClient::MessageSet->new(@$cur_block);

        {
            lock($sequence_queue);
            $sequence_queue->enqueue( $cur_msgset );
        }

        $offset += $opts->{max_fetch};

    }

    # Establish how many message sequence objects our threads need to
    # process....
    #
    {
        lock($sequence_queue);
        $total_sequences = $sequence_queue->pending();
    }


    # Start the caching thread.
    #
    {

        my $cache_put_thread =
            threads->create(
                             {
                                'context' => 'void',
                                'exit'    => 'thread_only'
                             },
                             \&cache_put_thread,
                                $folder,
                                $total_sequences,
                                $imap_msg_count
                           );

        print "Spawned caching thread.\n\n\n";

        $cache_put_thread->detach();

    }




    my $threads_available = 0;

    # Start our worker threads up first.
    #
    for ( 1 .. $opts->{threads} ) {
        my $work_q = Thread::Queue->new();
        my $thr = threads->create( \&imap_thread, $folder, $work_q );
        $thread_pumpkins{ $thr->tid() } = $work_q;
        $thr->detach();
    }

    ddump( 'thread_pumpkins', \%thread_pumpkins ) if $opts->{debug};

    ddump( 'sequences_completed', $sequences_completed ? $sequences_completed : 'zero' ) if $opts->{debug};
    ddump( 'total_sequences',     $total_sequences     ? $total_sequences     : 'zero' ) if $opts->{debug};


    THREADLOOP:
    while ( $sequences_completed < $total_sequences ) {

        # Give our threads a chance to start up and log in...
        #
        sleep 6;

        my $available_tid;
        my $cur_seq;

        ddump( 'sequences_complete', $sequences_completed ) if $opts->{debug};

        {

            ddump( 'step', 1 ) if $opts->{debug};

            # Check and see if there's any idle threads...
            #
            lock($idle_thread_queue);
            lock($sequence_queue);
           #lock($thread_pumpkins{$available_tid});

            ddump( 'step', 2 ) if $opts->{debug};

            $available_tid = $idle_thread_queue->dequeue();

            ddump( 'step', 3 ) if $opts->{debug};
            ddump( 'available_tid', $available_tid ) if $opts->{debug};
            ddump( 'available_tid', $available_tid ? $available_tid : 'zero' ) if $opts->{debug};

            # If no threads are available, go back and wait.
            #
            next THREADLOOP unless $available_tid;

            # Otherwise, there's a thread available. Look for a sequence of
            # message id's to give to our worker..
            #
            $cur_seq = $sequence_queue->extract();

            # Loop if we didn't get a M::I::MS object.
            #
            next THREADLOOP unless $cur_seq && ref $cur_seq;

            ddump( 'cur_seq3', $cur_seq ) if $opts->{debug};

            ddump( 'step', 4 ) if $opts->{debug};
            # We have messages to fetch, pass that sequence to the available
            # worker thread.
            #
            $thread_pumpkins{$available_tid}->enqueue($cur_seq);

        }


            ddump( 'step', 5 ) if $opts->{debug};

    }

            ddump( 'done_fetching message sequences', 1 ) if $opts->{debug};

    # Here, we've completed fetching all of our message sequences.
    #
    for ( keys %thread_pumpkins ) {
        # Tell all the threads they can go away...
        #
        $thread_pumpkins{$_}->enqueue(-1);
    }

    ddump( 'done_telling_threads_to_go_away', 1 ) if $opts->{debug};

    my $caching_done = 0;

    while ( ! $caching_done ) {

        sleep 4;

        ddump( 'waiting_for_caching_thread_to_finish', 1 ) if $opts->{debug};

        # Wait for our caching thread to finish up.
        #
        {
            lock($cache_put_status_queue);
            $caching_done = $cache_put_status_queue->extract();
        }

    }

    print "\n\nThreads complete...\n\n";


} # }}}

# {{{ imap_thread
#
# Expects to receive the folder name, current thread id, and the
# Mail::IMAPClient::MessageSet object.
#
# Returns nothing, simply adds the results to the thread-safe queue.
#
sub imap_thread {

    # Each thread gets an appropriate list of M::I::MessageSet objects
    #
    my $folder       = shift;
    my $work_q       = shift;

    my $thr = threads->self();
    $thr->set_thread_exit_only(1);

    my $tid = $thr->tid();


    my @headers;

    # TODO
    #
    # Fix this header handling...
    #
    push @headers, $header_table{$_} for qw/DATE SUBJECT SIZE TO FROM LISTID/;

    my %imap_options = %global_imap_options;

    if ( $opts->{debug} ) {

        open( CURDBG, '>>' . $opts->{log} . ".$tid" )
            or die_clean( 1, "Unable to open debuglog $tid: $!\n" );

        $imap_options{Debug}    = $opts->{debug} . '.' . $tid;
        $imap_options{Debug_fh} = *CURDBG;

    }

    my $imap_thread_socket =
        $opts->{Ssl}
        ? create_ssl_socket( 'Socket for thread: ' . $tid )
        : 0
        ;

    if ( $imap_thread_socket ) {
        $imap_options{Socket} = $imap_thread_socket;
    }

    my $imap_error;

    # Each thread gets its own imap object...
    #
    my $imap_thread = Mail::IMAPClient->new(%imap_options)
        or $imap_error = "Cannot connect to host : $@";

    if ( $imap_error ) {

        {
            lock($thread_errors_queue);
            $thread_errors_queue->enqueue($imap_error);
        }

        sleep 15;

        exit (1);

    }

    my $done = 0;

    while ( ! $done ) {

        # If we got this far, it's time to wait for some work...
        #
        {
            lock($idle_thread_queue);
            $idle_thread_queue->enqueue($tid);
        }

        sleep 2;

        my $cur_msgset;

        {
            lock($work_q);
            $cur_msgset = $work_q->extract();

        }

        if ( defined $cur_msgset ) {

            if ( ! ref $cur_msgset && $cur_msgset < 0 ) {
                $done = 1;
                last;
            }

        }

        if ( defined $cur_msgset && $cur_msgset && ! ref $cur_msgset ) {
            next;
        }

        if ( ! defined $cur_msgset ) {
            next;
        }

        # We got some work.  The parent has taken us out of the idle queue at
        # this point.


        unless ( $imap_thread->noop or $imap_thread->reconnect ) {
            show_error( "reconnect failed: $@\n" . $imap_thread->LastError );
            next;
        }



        # Reselect the folder so this thread is in the right place.  No
        # validation or exists check necessary at this stage.
        #
        if ( ! $imap_thread->examine($folder) ) {

            my $error = "\n\n\nERROR: Problem selecting folder $folder in thread $tid: $@\n\n\n\n";

            ddump( 'error', $error ) if $opts->{debug};

            {
                lock($thread_errors_queue);
                $thread_errors_queue->enqueue($error);
            }

            # We failed to fetch this sequence, so stick it back in the queue
            # and let another thread handle it.
            {
                lock($sequence_queue);
                $sequence_queue->enqueue( $cur_msgset );

            }

            # Introducing some delay to allow some time to pass before another
            # threads starts up to handle the messages that we put back into the
            # queue.
            #
            sleep 15;

            last;
        }


        while ( ! $done ) {

            # Make a M::I::MS object to get a clean range of message ids.
            #
        #my $cur_msgset  = Mail::IMAPClient::MessageSet->new(@$cur_block);

        #{
        #    lock($sequence_queue);
        #    $sequence_queue->enqueue( $cur_msgset );
        #}


            my @cur_msg_id_list = $cur_msgset->unfold;

            my $cur_msg_id_list_count = scalar(@cur_msg_id_list);

            next unless $cur_msg_id_list_count;

            unless ( $imap_thread->noop or $imap_thread->reconnect ) {
                show_error( "reconnect failed: $@\n" . $imap_thread->LastError );
                next;
            }


            my $cur_fetcher = $imap_thread->fetch_hash( \@cur_msg_id_list, @headers);

            if ( ! ref $cur_fetcher eq 'HASH' ) {

                my $error = "ERROR: fetching messages in folder $folder in thread $tid: $@\n";

                ddump( 'error', $error ) if $opts->{debug};

                # We failed to fetch this sequence, stick it back in the queue and
                # let another thread handle it.
                #
                {
                    lock($sequence_queue);
                    $sequence_queue->enqueue( $cur_msgset );

                }

                {
                    lock($thread_errors_queue);
                    $thread_errors_queue->enqueue($error);
                }

                # Introducing some delay to allow some time to pass before another
                # threads starts up to handle the messages that we put back into the
                # queue.
                #
                sleep 15;

            #   {
            #       lock($thread_pumpkins);
            #       my $foo = $thread_pumpkins->extract();
            #   }

                last;

            }


            {

                lock( $fetched_queue );

                ddump( 'cur_fetcher_before_stripping', $cur_fetcher ) if $opts->{debug};

                for my $m_id ( keys %$cur_fetcher ) {

                    next unless $m_id;

                    # Iterate the list of fetched messages and fix each value returned.
                    # The header information has some issues like CRLF and such.
                    #
                    for my $cur_header (@headers) {

                        $cur_fetcher->{$m_id}->{$cur_header} =
                            stripper( $header_table{$cur_header},
                                    $cur_fetcher->{$m_id}->{$cur_header} );

                    #$cur_fetcher->{$m_id}->{$cur_header} = '[EmptyValue]'
                    #    unless $cur_fetcher->{$m_id}->{$cur_header};

                    }

                    $fetched_queue->enqueue( $m_id => $cur_fetcher->{$m_id} );

                }

            }


            {
                lock($sequences_finished_queue);
                $sequences_finished_queue->enqueue( $cur_msgset );
            }


        }

        sleep 3;

    };


    if ( $imap_thread ) {

        my $rc = $imap_thread->_imap_command( 'LOGOUT', 'BYE' );

        #ddump( 'response code from LOGOUT', $rc );

        sleep 5;

        # This thread is done, tear down the imap connection.
        #
        $imap_thread->disconnect;

        if ( $opts->{Ssl} ) {
            $imap_thread_socket->close( %ssl_socket_close_options );
        }

        sleep 5;

    }


    exit(0);

} # }}}

# {{{ my_parse_headers
#
sub my_parse_headers {

    my ( $self, $msgspec, @fields ) = @_;
    my $fields = join ' ', @fields;
    my $msg = ref $msgspec eq 'ARRAY' ? $self->Range($msgspec) : $msgspec;
    my $peek = !defined $self->Peek || $self->Peek ? '.PEEK' : '';

    my $string = "$msg BODY$peek"
      . ( $fields eq 'ALL' ? '[HEADER]' : "[HEADER.FIELDS ($fields)]" );

    my $raw = $self->fetch($string) or return undef;
    my $cmd = shift @$raw;

    my %headers;    # message ids to headers
    my $h;          # fields for current msgid
    my $field;      # previous field name, for unfolding
    my %fieldmap = map { ( lc($_) => $_ ) } @fields;
    my $msgid;
    my $CR = \015;
    my $LF = \012;

    # BUG: parsing this way is prone to be buggy but works most of the time
    # some example responses:
    # * OK Message 1 no longer exists
    # * 1 FETCH (UID 26535 BODY[HEADER] "")
    # * 5 FETCH (UID 30699 BODY[HEADER] {1711}
    # header: value...
    foreach my $header ( map { split /$CR?$LF/o } @$raw ) {

        # Windows2003/Maillennium/others? have UID after headers
        if (
            $header =~ s/^\* \s+ (\d+) \s+ FETCH \s+
                        \( (.*?) BODY\[HEADER (?:\.FIELDS)? .*? \]\s*//ix
          )
        {    # start new message header
            ( $msgid, my $msgattrs ) = ( $1, $2 );
            $h = {};
            if ( $self->Uid )    # undef when win2003
            {
                $msgid = $msgattrs =~ m/\b UID \s+ (\d+)/x ? $1 : undef;
            }
            $headers{$msgid} = $h if $msgid;
        }
        $header =~ /\S/ or next;    # skip empty lines.

        # ( for vi
        if ( $header =~ /^\)/ ) {    # end of this message
            undef $h;                # inbetween headers
            next;
        }
        elsif ( !$msgid && $header =~ /^\s*UID\s+(\d+).*\)$/ ) {
            $headers{$1} = $h;       # found UID win2003/Maillennium

            undef $h;
            next;
        }

        unless ( defined $h ) {
            $self->_debug("found data between fetch headers: $header");
            next;
        }

        if ( $header and $header =~ s/^(\S+)\:\s*// ) {
            $field = $fieldmap{ lc $1 } || $1;
            push @{ $h->{$field} }, $header;
        }
        elsif ( $field and ref $h->{$field} eq 'ARRAY' ) {    # folded header
            $h->{$field}[-1] .= $header;
        }
        else {

            # show data if it is not like  '"")' or '{123}'
            $self->_debug("non-header data between fetch headers: $header")
              if ( $header !~ /^(?:\s*\"\"\)|\{\d+\})$CR?$LF$/o );
        }
    }

    # if we asked for one message, just return its hash,
    # otherwise, return hash of numbers => header hash
    ref $msgspec eq 'ARRAY' ? \%headers : $headers{$msgspec};
} # }}}


# }}}

# {{{ Cache operations

# {{{ cache_put_thread
#
sub cache_put_thread {

    my $folder          = shift;
    my $total_sequences = shift;
    my $imap_msg_count  = shift;

    my $thr = threads->self();
    $thr->set_thread_exit_only(1);

    my $cpt_cache = cache_init( $opts->{cache_file} );

    my $threads_active      = 0;
    my $sequences_completed = 0;
    my $cached_msg_count    = 0;


    # I suspected Term::Menus might be mucking with output buffering...
    #
    $|=1;

    my $tbar =
        IMAP::Report::Progress->new(
                                max           => $imap_msg_count,
                                length        => 10,
                                show_rotation => 1
                           );

    $tbar->text( 'Elapsed time: 0 seconds / 0 threads running' );
    $tbar->info( 'Thread stats: ' );
    $tbar->update( 0 );
    $tbar->write;

    my $start_time = time;

    my $done = 0;

    # Keep iterating while there are sequences in queue.
    #
    while ( ! $done ) {

        sleep 3;

        my $cur_time = time;

        my $elapsed_time = convert_seconds( $cur_time - $start_time );

        $tbar->text(
            "Elapsed time: $elapsed_time / $threads_active threads running"
        );

        my %extracted_messages;
        my $pending;

        {

            lock($fetched_queue);

            $pending = $fetched_queue->pending();

            if ( $pending ) {

                verbose("\n\nTotal messages pending for processing: $pending");
                %extracted_messages = $fetched_queue->extract( 0, $pending );

            }

        }

        my $ex_count = scalar( keys %extracted_messages );
        next unless $ex_count;

        ddump( 'extracted_messages', \%extracted_messages ) if $opts->{debug};

        cache_put({ cache => $cpt_cache, content_type => 'fetched_messages', folder => $folder, values => \%extracted_messages });

        $cached_msg_count += scalar( keys %extracted_messages );

        my $tstat = '';

        {

            # Check if any of our threads have produced errors and display
            # them in our progress bar.
            #
            # TODO
            #
            # End result is that the error only gets displayed for 2
            # seconds...   need better handling...
            #
            lock($thread_errors_queue);

            my $msg = $thread_errors_queue->extract();

            $tstat =
                defined $msg && $msg
                ? 'Thread status: ' . $msg
                : 'Thread status: ok '
                ;

            ddump( 'msg', $msg ) if defined $msg && $msg;

        }


        $tbar->info( $tstat );
        #$tbar->update($sequences_completed);
        $tbar->update($cached_msg_count);
        $tbar->write;

        {

           #lock($cached_msg_count_queue);
           #$cached_msg_count_queue->enqueue($cur_msg_count);

            lock($sequences_finished_queue);
            $sequences_completed = $sequences_finished_queue->pending();


            # If all of our expected sequences are complete and there are no
            # more threads active, we can go away...
            #
           #if ( $sequences_completed == $total_sequences && ! $threads_active >= 1 )
            if ( $sequences_completed == $total_sequences ) {
                $done = 1;
            }

        }

    }

    $cpt_cache->disconnect;

    # Before we go away, stick a message in the cache status queue to indicate
    # our that we're done.
    #
    $cache_put_status_queue->enqueue( 1 );

   #$thr->exit(0);
    exit(0);


} # }}}

# }}}

# {{{ Message processing

# {{{ threaded_sequence_chunker
#
# This is where we take our big list of message id's and split it into chunks.
# Expects to receive one param, a Mail::IMAPClient::MessageSet object containing
# our message id's.
#
# Returns an array of M::I::MessageSet objects, carefully sorted into different
# buckets so that no threads will be attempting to fetch the same messages.
# Divides the number of messages evenly between our number threads.
#
# The reason for the back-and-forth switching from lists to MessageSet objects
# is because M::I::MessageSet makes the effort to express the list of message
# ids in optimal RFC2060 representation.
#
sub threaded_sequence_chunker {

    my $all_msg_ids = shift;

    show_error( 'Error processing sequences' ) unless $all_msg_ids;

    my $message_count = scalar(@$all_msg_ids);

    #my $max = int( $message_count / $opts->{threads} );

    my $max = 500;

    my @cur_block;

   #my $msg_id_buckets  = {};
   #my $msg_set_buckets = {};

    my $counter = 0;

    while ( my @cur_block = splice @$all_msg_ids, 0, $max ) {

       #my $bucket_id = $counter % $opts->{threads};

        my $cur_set = Mail::IMAPClient::MessageSet->new( @cur_block );
        $sequence_queue->enqueue( $cur_set );

       #push @{$msg_set_buckets->{$bucket_id}}, $cur_set;
       #$counter++;

    }

#   ddump( 'msg_set_buckets', $msg_set_buckets ) if $opts->{debug};

#   return $msg_set_buckets;

} # }}}

# {{{ old_threaded_sequence_chunker
#
# This is where we take our big list of message id's and split it into chunks.
# Expects to receive one param, a Mail::IMAPClient::MessageSet object containing
# our message id's.
#
# Returns an array of M::I::MessageSet objects, carefully sorted into different
# buckets so that no threads will be attempting to fetch the same messages.
# Divides the number of messages evenly between our number threads.
#
# The reason for the back-and-forth switching from lists to MessageSet objects
# is because M::I::MessageSet makes the effort to express the list of message
# ids in optimal RFC2060 representation.
#
sub old_threaded_sequence_chunker {

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

    ddump( 'msg_set_buckets', $msg_set_buckets ) if $opts->{debug};

    return $msg_set_buckets;

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
# To display a helpful string you can cut and paste directly into the gmail
# search pane to search for the reported message.
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
              "SUBJECT:\"$subject\"",
              $adate,
              $bdate,
              $label );

} # }}}

# {{{ generate_search_string
#
# To display a helpful string you can cut and paste directly into the gmail
# search pane to search for the reported message.
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

# }}}

# {{{ Converters

# }}}

# {{{ Menuing

# {{{ choose_action
#
# Expects to receive a string to use as the banner to display at the top of the
# menu screen.
#
# Returns the choice corresponding with the menu item selected from Term::Menus,
# which is the verbatim description of that particular menu option.
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
    my $choice = show_folder_picker({ imap_folders => $folders,
                                      banner       => "Choose folder for report type: $type" });

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

    ddump( 'folders', $folder_list ) if $opts->{debug};

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

    return $choice;


} # }}}

# }}}

# {{{ Progress info

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

        # The statistics are fully recomputed from scratch with every iteration.
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
# Pretty print the stats of the current fetch operation.  Will replace with a
# progressbar.
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
#sub estimate_completion_time {
#
#    my $stats                  = shift;
#    my $total_number_of_blocks = shift;
#
#    my $number_of_blocks_completed = scalar(@$stats);
#
#    my $elapsed_seconds_completed = 0;
#
#    for (@$stats) {
#        $elapsed_seconds_completed += $_->[0];
#    }
#
#    my $rate          = $elapsed_seconds_completed / $number_of_blocks_completed;
#    my $total_eta     = $rate * $total_number_of_blocks;
#    my $remaining_eta = $total_eta - $elapsed_seconds_completed;
#
#    return convert_seconds($remaining_eta);
#
#} # }}}

# }}}

# {{{ Execution and configuration

# {{{ show_error
#
sub show_error {

    my $error = shift;

    print "\n\n\n$error\n";

    enter();

} # }}}

# {{{ read_config_file
#
# Allow a home directory config file containing your password and other options
# to be set.
#
# Simple foo=bar syntax.  Ignore comments, strip leading and trailing spaces.
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

    if ( $opts->{debug} ) {
        close DBG;
    }

    print "\n$msg\n";

    if ( $err ) {
        verbose( "Exiting with status: $err" );
        exit $err;
    } else {
        verbose( "Exiting with clean status." );
        exit;
    }

} # }}}

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
        chomp( $user = <> );
        print "\n";
    }
    return $user;

}

# }}}

# {{{ debugging output
#
sub fddump {

    $Data::Dumper::Varname = shift;

    open( DDBG, '>>' . $opts->{log} . 'dumperlog' )
        or die_clean( 1, "Error opening dumperlog: $!\n" );

    print DDBG Dumper( @_ );

    close DDBG;

}

sub ddump {

    $Data::Dumper::Varname = shift;

    print Dumper( @_ );

}

sub verbose {

    return unless $opts->{verbose};

    my $v = shift;

    print "\n$v\n";

    # Only pause for user input if we're in debug mode...
    #
    enter() if $opts->{debug};

}

sub enter {

    print "\nPress [Enter] to continue: ";

    chomp( my $input = <> );

    return $input;

}

# }}}

# }}}

# }}}

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

=item B<--filters> I<string>

Folder filters.  Restrict all operations to folders matching the specified string.  This option can be specified multiple times.

=item B<--exclude> I<string>

Folder exclusions.  The list of folders will be pruned of the ones matching the specified string.  This option can be specified multiple times.  Perl compatible regex should work as long as you take care not to allow your shell to swallow up the expression.

=item B<--use_threaded_mode>

Use multiple threads to fetch messages simultaneously.  Speeds things up dramatically, but the underlying Net::SSLeay isn't very thread safe.  Tried to make it as stable as possible keep the threads from aggressively tearing down the imap connection when they're done, but this option not may not work well for you.  You can also fairly easily hit Google's bandwidth limit for your imap connection.

(Redundant, ugly commandline option name is so you'll only use it very intentionally.)

(default: false)

=item B<--threads> (valid values: 2, 3, 5, 7, 11, 13, 17)

Specify the number of threads.  Higher is faster, but too high will engage the temporary bandwidth usage ban from Google.

(default: 0)

=item B<--cache> I<cache_filename>

Name of the file used to store cached information.

(default: $HOME/.imap-report.cache)

=item B<--cache_age> I<integer>

Maximum age of cached information in days.

(Pruning of the cache can be suppressed with --no-cache_prune).

(default: 7)

=item B<--cache_only>

Supresses almost all IMAP operations and works from cache only.

(default: false)

=item B<--threshold>

The message count threshold where the number of messages in cache differs from the number of messages on the server so that a folder isn't fully refetched just because a single new message was received.

(default: 20)

=item B<--conf> I<config_filename>

Name of the file from which to read configuration options.

All of these configuration options can be stored in the specified file using the same names listed here.  Must only be readable by the user.

(default: $HOME/.imapreportrc)

=item B<--list>

Just show the list of folders.

=item B<--types>

Just show the types of reports available for use with the --report option.

=item B<--report>

Specify a specific type of report shown by the --types option.

=item B<--pager>

The pager to use for displaying the report.

(default: /usr/bin/less)

=item B<--debug>

Lots of ugly debugging output to a logfile (--log)

=item B<--verbose>

A bit more output than usual

=back

=head3 Mail::IMAPClient pass-through options

=over 15

=item B<--Fast_io>

Corresponds to the Mail::IMAPClient Fast_io option to allow buffered I/O.

(default: true)

=item B<--Keepalive>

Corresponds to the Mail::IMAPClient Keepalive option.

(default: false)

=item B<--Maxcommandlength>

Corresponds to the Mail::IMAPClient Maxcommandlength option to limit the size of individual fetches.

(default: 1000)

=item B<--Reconnectretry>

Corresponds to the Mail::IMAPClient Reconnectretry option to try and re-establish lost connections.

(default: 3)

=item B<--Ssl>

Corresponds to the Mail::IMAPClient Ssl option.

(default: true)

=back

=head2 EXAMPLE

C<./imap-report.pl>

(No options are necessary to run this script.  See the description of options below for how to override the default settings.  Run perldoc imap-report.pl to for further instruction.)

=head1 ACKNOWLEDGEMENTS

Built largely using Mail::IMAPClient currently maintained by E<lt>L<PLOBBES|http://search.cpan.org/~plobbes/>E<gt>,
the Mail::Address module by E<LT>L<MARKOV|http://search.cpan.org/~markov/>E<gt> (also a former maintainer
of Mail::IMAPClient), the Term::Menus module by E<lt>L<REEDFISH|http://search.cpan.org/~reedfish/>E<gt>, along
with String::ProgressBar from E<lt>L<AHERNIT|http://search.cpan.org/~ahernit/>E<gt>.

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

