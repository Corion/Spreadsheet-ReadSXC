use strict;
use Test::More tests => 3;
use File::Basename 'dirname';
use Spreadsheet::ParseODS;
use Data::Dumper;

my $d = dirname($0);

my $workbook = Spreadsheet::ParseODS->new()->parse("$d/print-area.ods");

my $areas = $workbook->get_print_areas;
is_deeply $areas, [[1,1,4,1],undef], "Retrieving all print areas works"
    or diag Dumper $areas;

my $area1 = $workbook->worksheet('printarea')->get_print_area;
is_deeply $area1, [1,1,4,1], "Retrieving all print areas works"
    or diag Dumper $area1;

my $area2 = $workbook->worksheet('no printarea')->get_print_area;
is $area2, undef, "A sheet without a print area has undef"
    or diag Dumper $area2;

