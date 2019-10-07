use strict;
use Test::More tests => 2;
use File::Basename 'dirname';
use Spreadsheet::ParseODS;
use Data::Dumper;

my $d = dirname($0);

my $workbook = Spreadsheet::ParseODS->new()->parse("$d/t.sxc");

my @sheets = sort keys %$workbook;

is_deeply \@sheets, [qw[
    Sheet1 Sheet2 Sheet3
]], "Correct spreadsheet names"
or diag Dumper \@sheets;

my @sheet1_raw = (['-$1,500.99', '17', undef],[undef, undef, undef],['one', 'more', 'cell']);
my @sheet1_curr = ([-1500.99, 17, undef],[undef, undef, undef],['one', 'more', 'cell']);

is_deeply $workbook->{Sheet1}, \@sheet1_raw, "Raw cell values (to be changed)"
    or diag Dumper $workbook->{Sheet1};

my @sheet1_curr_date_multiline = (
    [-1500.99, 17, undef],
    [undef, undef, undef],
    ['one', 'more', 'cell'],
    [undef,undef,undef],
    ['Date','1980-11-21', undef],
    ["A cell value\nThat contains\nMultiple lines",undef,undef],
    ["\nA cell that starts\nWith an empty line\nAnd ends with an empty\nLine as well\n",undef,undef],
);
