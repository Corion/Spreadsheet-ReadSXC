#!perl
use strict;
use Test::More tests => 5;
use File::Basename 'dirname';
use Spreadsheet::ReadSXC;

my $d = dirname($0);
my $sxc_file = "$d/t.sxc";

sub dies_ok {
    my( $code, $error_msg, $name ) = @_;
    $name ||= $error_msg;

    my $old_handler = Archive::Zip::setErrorHandler(sub {});

    my $died = eval {
        $code->();
        1
    };
    my $err = $@;
    is $died, undef, $name;
    like $err, $error_msg, $name;

    Archive::Zip::setErrorHandler($old_handler);
};

is Spreadsheet::ReadSXC::read_sxc('no-such-file.sxc'), undef, "Default silent API";

dies_ok sub { Spreadsheet::ReadSXC::read_sxc('no-such-file.sxc', { StrictErrors => 1 }) }, qr/Couldn't open 'no-such-file.sxc':/,"Non-existent file";
dies_ok sub { Spreadsheet::ReadSXC::read_sxc_fh(undef) }, qr//, "undef filehandle";
