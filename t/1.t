use strict;
use Test::More tests => 8;
BEGIN { use_ok('Spreadsheet::ReadSXC') };
BEGIN { use_ok('Archive::Zip') };
BEGIN { use_ok('XML::Parser') };

my $zip = Archive::Zip->new();
ok(( $zip->read("t.sxc") == 0 ), 'Unzipping .sxc file');

my $workbook_ref = Spreadsheet::ReadSXC::read_sxc("t.sxc");

my @sheets = sort keys %$workbook_ref;

ok((($sheets[0] eq "Sheet1") and ($sheets[1] eq "Sheet2") and ($sheets[2] eq "Sheet3")), 'Comparing spreadsheet names');

my @sheet1_data = (['-$1,500.99', '17', undef],[undef, undef, undef],['one', 'more', 'cell']);
my @sheet3_data = (['Both alike', 'Both alike', undef], [undef, undef, undef], [undef, undef, undef], [undef, undef, undef], [undef, undef, undef], [undef, undef, undef], [undef, undef, undef], [undef, undef, undef], [undef, undef, undef], [undef, undef, undef], [undef, undef, undef], [undef, undef, undef], [undef, undef, undef], [undef, undef, 'Cell C14']);

my @sheet1 = @{$$workbook_ref{"Sheet1"}};
is_deeply \@sheet1, \@sheet1_data, 'Verifying Sheet1';

is_deeply $workbook_ref->{"Sheet2"}, [], 'Verifying Sheet2';

my @sheet3 = @{$$workbook_ref{"Sheet3"}};
is_deeply \@sheet3, \@sheet3_data, 'Verifying Sheet3';
