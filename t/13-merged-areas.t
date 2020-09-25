use strict;
use Test::More;
use File::Basename 'dirname';
use Spreadsheet::ParseODS;
use Data::Dumper;

my $d = dirname($0);

plan tests => 1;

my $workbook = Spreadsheet::ParseODS->new()->parse("$d/merged.ods");

my $merged_areas = [$workbook->worksheets()]->[0]->merged_areas();
is_deeply $merged_areas, [
              [0,1,1,2], # B1:C2
              [1,0,2,0], # A2:A3
          ],
          "We read the proper merged areas"
    or diag Dumper $merged_areas;
