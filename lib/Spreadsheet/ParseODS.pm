package Spreadsheet::ParseODS;
use strict;
use warnings;

use Archive::Zip ':ERROR_CODES';
use Moo 2;
use XML::Parser;
use XML::Twig::XPath;
use Carp qw(croak);
use List::Util 'max';
use Storable 'dclone';

our $VERSION = '0.23';
our @CARP_NOT = (qw(XML::Twig));

use Filter::signatures;
use feature 'signatures';
no warnings 'experimental::signatures';

=head1 NAME

Spreadsheet::ParseODS - read SXC and ODS files

=head1 WARNING

This module is not yet API-compatible with Spreadsheet::ParseXLSX
and Spreadsheet::ParseXLS

=head1 METHODS

=head2 C<< ->new >>

=cut

has 'line_separator'       => ( is => 'ro', default => "\n", );
has 'IncludeCoveredCells'  => ( is => 'ro', default => 0,  );
has 'DropHiddenRows'       => ( is => 'ro', default => 0,  );
has 'DropHiddenColumns'    => ( is => 'ro', default => 0,  );
has 'NoTruncate'           => ( is => 'ro', default => 0,  );

has 'twig' => (
    is => 'lazy',
    default => sub {
        XML::Twig::XPath->new(
            no_xxe => 1,
            keep_spaces => 1,
        )
    },
);

=head2 C<< ->parse >>

=cut

sub parse {
    my( $self, $source, $formatter ) = @_;
    my $p = $self->twig;

    # Convert to ref, later
    my %workbook = ();
    my @worksheets = ();
    my @sheet_order = ();
    my @cell = ();
    my $row_hidden = 0;
    my $max_datarow = -1;
    my $max_datacol = -1;
    my $col_count = -1;
    my @hidden_cols = ();

    my %handlers;

#    $handlers{ 'text:p' } = sub {
#        my( $element ) = @_;
#
#        if( $element->parent->nodeType ne 'office:annotation' ) {
#            # Add to current cell
#            push @cell, $element->value;
#        };
#    };
#
#    $handlers{ 'table:table-cell' }
#    = $handlers{ 'table:covered-table-cell' } = sub {
#        my( $element ) = @_;
#
## increase cell count
#        $col++;
## if number-columns-repeated is set, set $repeat_cells value accordingly for later use
#        my $repeat = $element->att('table:number-columns-repeated');
#        if( defined $repeat ) {
#            $repeat_cells = $repeat;
#        }
## save the currency value (if available)
#        if (exists $attributes{'table:value'} or exists $attributes{'office:value'} ) {
#            $currency_value = $attributes{'table:value'} || $attributes{'office:value'};
#        }
## if cell contains date or time values, set boolean variable for later use
#        elsif (exists $attributes{'table:date-value'} or exists $attributes{'office:date-value'}) {
#            $date_value = $attributes{'table:date-value'} || $attributes{'office:date-value'};
#        }
#        elsif (exists $attributes{'table:time-value'} or exists $attributes{'office:time-value'}) {
#            $time_value = $attributes{'table:time-value'} || $attributes{'office:time-value'};
#        }
#    };

    $handlers{ "table:table" } = sub {
        my( $twig, $table ) = @_;

        my $max_datarow = -1;
        my $max_datacol = -1;
        @hidden_cols = ();

        my $tablename = $table->att('table:name');
        my $tableref = $workbook{ $tablename } = [];
        my $table_hidden = $table->att( 'table:visibility' );

        # Look at table:column and decide other stuff
#    $handlers{ "table:table-column" } = sub {
#        my( $element ) = @_;
## increase column count
#        $col_count++;
## if columns is hidden, add column number to @hidden_cols array for later use
#        my $hidden = defined $element->att('table:visibility');
#        if (  $hidden ) {
#            push @hidden_cols, $col_count;
#        };
#
## if number-columns-repeated is set and column is hidden, add affected columns to @hidden_cols
#        if ( my $repeat = $element->att('table:number-columns-repeated') ) {
#            $col_count++;
#            if ( $hidden ) {
#                for (2..$repeat ) {
#                    push @hidden_cols, $hidden_cols[$#hidden_cols] + 1;
#                }
#            }
#        }
#    };

        for my $row ($table->findnodes('.//table:table-row')) {
            my $row_hidden = $table->att( 'table:visibility' );
# if row is hidden, set $row_hidden for later use
# if number-rows-repeated is set, set $repeat_rows value accordingly for later use

            my $rowref = [];

            my $repeat_row = 0;
            my $repeat = $row->att('table:number-rows-repeated');
            if( defined $repeat ) {
                $repeat_row = $repeat -1;
            };

            #my $row_has_content = 1;

            # Do we really only want to add a cell if it contains text?!
            my $colnum = -1;
            for my $cell ($row->findnodes("./table:table-cell | ./table:covered-table-cell")) {
                my $repeat = $cell->att('table:number-columns-repeated') || 1;
                for my $i (1..$repeat) {
                    # Yes, this is somewhat inefficient, but it saves us
                    # from later programming errors if we create/store
                    # references. We can always later turn this inside-out.
                    $colnum++;
                    if( $cell->is_empty ) {
                        push @$rowref, Spreadsheet::ParseODS::Cell->new({
                            type        => undef,
                            unformatted => undef,
                            value       => undef,
                        });
                    } else {
                        my ($text,$type,$unformatted);

                        $type =     $cell->{att}->{"office:value-type"} # ODS
                                 || $cell->{att}->{"table:value-type"}  # SXC
                                 || '' ;
                        ($unformatted) = grep { defined($_) }
                                       $cell->{att}->{"office:value"}, # ODS
                                       $cell->{att}->{"table:value"},  # SXC
                                       $cell->{att}->{"office:date-value"}, # ODS
                                       $cell->{att}->{"table:date-value"},  # SXC
                                       ;

                        if( $type ) {
                            # $row_has_content = 1;
                        };

                        my @text = $cell->findnodes('text:p');
                        if( @text ) {
                            $text = join $self->line_separator, map { $_->text } @text;
                            $max_datacol = max( $max_datacol, $colnum );
                        } else {
                            $text = $unformatted;
                        };

                        my $cell = Spreadsheet::ParseODS::Cell->new({
                            value       => $text,
                            unformatted => $unformatted,
                            type        => $type,
                        });

                        push @$rowref, $cell;
                    };
                };
            };

            push @$tableref, $rowref;
            $max_datarow++;

            for my $r (1..$repeat_row) {
                push @$tableref, dclone( $rowref );
                $max_datarow++;
            };
        }

        # decrease $max_datacol if hidden columns within range
        if ( ( ! $self->NoTruncate ) and ( $self->DropHiddenColumns ) ) {
            for ( 1..scalar grep { $_ <= $max_datacol } @hidden_cols ) {
                $max_datacol--;
            }
        }

        # truncate/expand table to $max_datarow and $max_datacol
        if ( ! $self->NoTruncate ) {
            $#{$tableref} = $max_datarow;
            foreach ( @{$tableref} ) {
                $#{$_} = $max_datacol;
            }
        }

        @$tableref = ()
            if $max_datacol == 0;

        my $ws = Spreadsheet::ParseODS::Worksheet->new({
                label => $tablename,
                data  => \@{$workbook{$tablename}},
                col_min => 0,
                col_max => $max_datacol,
                row_min => 0,
                row_max => $max_datarow,
        });

        # set up alternative data structure
        push @worksheets, $ws;
        $workbook{ $tablename } = $ws;
    };

    $p->setTwigHandlers( \%handlers );

    my ($method, $xml) = $self->_open_xml_thing( $source );
    $p->$method( $xml );

    return Spreadsheet::ParseODS::Workbook->new(
        _worksheets => \%workbook,
        _sheets => \@worksheets
    );
};

sub _open_xml_thing( $self, $source ) {
    my $ref = ref($source);
    my $xml;
    my $method = 'parse';

    if( ! $ref ) {
        # Specified by filename .
        # $workbook{File} = $source;

        croak "Undef ODS source given"
            unless defined $source;

        if( $source =~ m!(\.xml|\.fods)!i ) {
            # XXX also handle some option that specifies that we want to
            #     parse raw XML here
            $method = 'parsefile';
            $xml = $source;

        } else {
            $xml = $self->_open_sxc( $source )
        };

    } else {
        if ( $ref eq 'SCALAR' ) {
            # Specified by a scalar buffer.
            # XXX We create a copy here. Maybe we should be able to feed
            #     this to XML::Twig without creating (another) copy here?
            #     Or will CoW save us here anyway?
            $xml = $$source;
            #$workbook{File} = undef;

        } elsif ( $ref eq 'ARRAY' ) {
            # Specified by file content
            #$workbook{File} = undef;
            $xml = join( '', @$source );

        } else {
             # Assume filehandle
             # Kick off XML::Twig from Filehandle
             $xml = $self->_open_sxc_fh( $source );
         }
    }

    return ($method, $xml)
}

sub _open_sxc {
    my ($self, $sxc_file, $options_ref) = @_;
    if( !$options_ref->{StrictErrors}) {
        -f $sxc_file && -s _ or return undef;
    };
    open my $fh, '<', $sxc_file
        or croak "Couldn't open '$sxc_file': $!";
    return $self->_open_sxc_fh( $fh, $options_ref );
}

sub _open_sxc_fh {
    my ($self, $fh, $options_ref) = @_;
    my $zip = Archive::Zip->new();
    my $status = $zip->readFromFileHandle($fh);
    $status == AZ_OK
        or croak "Read error from zip";
    my $content = $zip->memberNamed('content.xml');
    $content->rewindData();
    my $stream = $content->fh;
    binmode $stream => ':gzip(none)';
    $stream
}

package Spreadsheet::ParseODS::Workbook;
use Moo 2;
use Filter::signatures;
use feature 'signatures';
no warnings 'experimental::signatures';

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

sub get_filename( $self ) {
    $self->filename
}

sub worksheets( $self ) {
    @{ $self->_sheets }
};

sub worksheet( $self, $name ) {
    $self->_worksheets->{ $name }
}

package Spreadsheet::ParseODS::Worksheet;
use Moo 2;
use Filter::signatures;
use feature 'signatures';
no warnings 'experimental::signatures';

has 'label' => (
    is => 'rw'
);

has 'data' => (
    is => 'rw'
);

has 'sheet_hidden' => (
    is => 'rw',
);

has 'row_min' => (
    is => 'rw',
);

has 'row_max' => (
    is => 'rw',
);

has 'col_min' => (
    is => 'rw',
);

has 'col_max' => (
    is => 'rw',
);

sub get_cell( $self, $row, $col ) {
    return undef if $row > $self->row_max;
    return undef if $col > $self->col_max;
    $self->data->[ $row ]->[ $col ]
}

sub get_name( $self ) {
    $self->name
}

sub is_sheet_hidden( $self ) {
    $self->sheet_hidden
}

sub row_range( $self ) {
    return ($self->row_min, $self->row_max)
}

sub col_range( $self ) {
    return ($self->col_min, $self->col_max)
}

package Spreadsheet::ParseODS::Cell;
use Moo 2;
use Filter::signatures;
use feature 'signatures';
no warnings 'experimental::signatures';

has 'value' => (
    is => 'rw',
);

has 'unformatted' => (
    is => 'rw',
);

has 'type' => (
    is => 'rw',
);

1;
