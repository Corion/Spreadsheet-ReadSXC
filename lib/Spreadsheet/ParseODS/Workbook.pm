package Spreadsheet::ParseODS::Workbook;
use Moo 2;
use Filter::signatures;
use feature 'signatures';
no warnings 'experimental::signatures';

our $VERSION = '0.26';

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

has '_settings' => (
    is => 'rw',
    handles => [ 'active_sheet_name' ],
);

# The worksheets themselves
has '_sheets' => (
    is => 'lazy',
    default => sub { [] },
);

# Mapping of names to sheet objects
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

# <config:config-item config:name="ActiveTable" config:type="string">Sheet3</config:config-item>
sub get_active_sheet($self) {
    $self->worksheet( $self->active_sheet_name );
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
