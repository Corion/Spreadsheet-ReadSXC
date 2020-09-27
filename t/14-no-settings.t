use strict;
use Test::More;
use File::Basename 'dirname';
use Spreadsheet::ParseODS;

my $d = dirname($0);

plan tests => 2;

my $workbook;
my $ok = eval {
    $workbook = Spreadsheet::ParseODS->new(
        #readonly => 1,
    )->parse("$d/20200617_Testnummers_inclusief_omnummertabel_GBA-V.ods",
        readonly => 1
    );
    1;
};

is $ok, 1, "We don't crash when parsing the workbook"
    or diag $@;
ok $workbook->worksheet('Toelichting'), 'We find the worksheet "Toelichting"';
