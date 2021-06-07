#!/usr/bin/perl

use LWP::Simple;
use URI::Escape;
use Data::Dumper;
use strict;
use warnings;

my $where = shift @ARGV
    or die "Usage: $0 \"111 Main St, Anytown, KS\"\n";

my $addr = uri_escape($where);

my @result = get("http://rpc.geocoder.us/service/csv?address=$addr" );

print Dumper \@result;

