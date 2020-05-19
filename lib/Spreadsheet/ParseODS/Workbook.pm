package Spreadsheet::ParseODS::Workbook;
use Moo 2;
use Filter::signatures;
use feature 'signatures';
no warnings 'experimental::signatures';
use PerlX::Maybe;

our $VERSION = '0.24';

=head1 NAME

Spreadsheet::ParseODS::Workbook - a workbook

=cut

=head2 C<< ->filename >>

  print $workbook->filename;

The name of the file if applicable.

=cut

has 'filename' => (
    is => 'rw',
);

has '_sheets' => (
    is => 'lazy',
    default => sub { {} },
);

has '_worksheets' => (
    is => 'lazy',
    default => sub { {} },
);

has '_styles' => (
    is => 'lazy',
    default => sub { {} },
);

=head2 C<< ->table_styles >>

The styles that identify whether a table is hidden, and other styles

=cut

has 'table_styles' => (
    is      => 'lazy',
    default => sub { {} },
);

=head2 C<< get_print_areas() >>

    my $print_areas = $workbook->get_print_areas();
    # [[ [$start_row, $start_col, $end_row, $end_col], ... ]]

The C<< ->get_print_areas() >> method returns the print areas
of each sheet as an arrayref of arrayrefs. If a sheet has no
print area, C<undef> is returned for its print area.

=cut

sub get_print_areas( $self ) {
    [ map { $_->get_print_areas } $self->worksheets ]
}

sub get_filename( $self ) {
    $self->filename
}

sub worksheets( $self ) {
    @{ $self->_sheets }
};

sub worksheet( $self, $name ) {
    $self->_worksheets->{ $name }
}

1;
