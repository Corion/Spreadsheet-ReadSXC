package Spreadsheet::ParseODS;
use strict;
use warnings;

use Archive::Zip ':ERROR_CODES';
use Moo 2;
use XML::Parser;
use XML::Twig::XPath;
use Carp qw(croak);
use List::Util 'max';

our $VERSION = '0.23';

=head1 NAME

Spreadsheet::ParseODS - read SXC and ODS files

=head1 WARNING

This module is not yet API-compatible with Spreadsheet::ParseXLSX
and Spreadsheet::ParseXLS

=head1 METHODS

=head2 C<< ->new >>

=cut

has 'ReplaceNewlineWith'   => ( is => 'ro', default => "", );
has 'IncludeCoveredCells'  => ( is => 'ro', default => 0,  );
has 'DropHiddenRows'       => ( is => 'ro', default => 0,  );
has 'DropHiddenColumns'    => ( is => 'ro', default => 0,  );
has 'NoTruncate'           => ( is => 'ro', default => 0,  );
has 'StandardCurrency'     => ( is => 'ro', default => 0,  );
has 'StandardDate'         => ( is => 'ro', default => 0,  );
has 'StandardTime'         => ( is => 'ro', default => 0,  );
has 'OrderBySheet'         => ( is => 'ro', default => 0,  );
has 'StrictErrors'         => ( is => 'ro', default => 0,  );

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
    my $repeat_cells = 1;
    my $repeat_rows = 1;
    my $row_hidden = 0;
    my $date_value = '';
    my $time_value = '';
    my $currency_value = '';
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

        my $tablename = $table->{att}->{'table:name'};
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

            my $repeat = $row->att('table:number-rows-repeated');
            if( defined $repeat ) {
                $repeat_rows = $repeat;
            };

            my $row_has_content = 0;

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
                        push @$rowref, undef
                    } else {
                        my ($text,$type, $value);

                        $type =     $cell->{att}->{"office:value-type"} # ODS
                                 || $cell->{att}->{"table:value-type"}  # SXC
                                 || '' ;
                        ($value) = grep { defined($_) }
                                       $cell->{att}->{"office:value"}, # ODS
                                       $cell->{att}->{"table:value"},  # SXC
                                       $cell->{att}->{"office:date-value"}, # ODS
                                       $cell->{att}->{"table:date-value"},  # SXC
                                       ;

                        if( $type ) {
                            $row_has_content = 1
                        };

                        if( $type eq 'currency' and $self->StandardCurrency ) {
                            $text = $value

                        } elsif( $type eq 'date' and $self->StandardDate ) {
                            $text = $value

                        } elsif( $type =~ qr/^(float|percentage)/ ) {
                            $text = $value

                        } else {
                            my @text = $cell->findnodes('text:p');
                            if( @text ) {
                                for my $line (@text) {
                                    #$row_has_content = 1;
                                    $max_datacol = max( $max_datacol, $colnum );
                                    $text = '' if ! defined $text;
                                    $text .= ''.$line->text;
                                };
                            } else {
                                $text = $value;
                            };
                            $row_has_content = $row_has_content || defined $text;
                        };
                        push @$rowref, $text;
                    };
                };
            };
            if( $row_has_content ) {
                push @$tableref, $rowref;
                $max_datarow++;
                $max_datacol = max( $max_datacol, $#$rowref );
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

# set up alternative data structure
        if ( $self->OrderBySheet ) {
            push @worksheets, (
                {
                    label   => $table,
                    data    => \@{$workbook{$table}},
                }
            );
        };
    };

    $p->setTwigHandlers( \%handlers );

#    my $ref = ref($source);
#    my $xml;
#    my $method = 'parse';
#
#    if( ! $ref ) {
#        # Specified by filename .
#        $workbook{File} = $source;
#
#        if( $source =~ m!\.xml$i ) {
#            # XXX also handle some option that specifies that we want to
#            #     parse raw XML here
#            $method = 'parsefile';
#            $xml = $source;
#
#        } else {
#            $xml = $self->_open_sxc( $source )
#        };
#
#    } else {
#        if ( $ref eq 'SCALAR' ) {
#            # Specified by a scalar buffer.
#            # XXX We create a copy here. Maybe we should be able to feed
#            #     this to XML::Twig without creating (another) copy here?
#            #     Or will CoW save us here anyway?
#            $xml = $$source;
#            $workbook{File} = undef;
#
#        } elsif ( $ref eq 'ARRAY' ) {
#            # Specified by file content
#            $workbook{File} = undef;
#            $xml = join( '', @$source );
#
#        } else {
#             # Assume filehandle
#             # Kick off XML::Twig from Filehandle
#             $xml = $self->_open_sxc_fh( $source );
#         }
#    }

    my $xml = $self->_open_sxc( $source );
    $p->parse( $xml );

    \%workbook
};

sub _fetch_cell_value {
    my( $self, $element ) = @_;
    my $className = $element->tag;

    $element->value()

#=for later
#    if ( ( $tag eq "table:table-cell" ) or ( $tag eq "table:covered-table-cell" ) ) {
## assign currency, date or time value to current workbook cell if requested
#        if ( ( $self->StandardCurrency ) and ( length( $currency_value ) ) ) {
#            $workbook{$table}[$row][$col] = $currency_value;
#            $currency_value = '';
#
#        }
#        elsif ( ( $options{StandardDate} ) and ( $date_value ) ) {
#            $workbook{$table}[$row][$col] = $date_value;
#            $date_value = '';
#        }
#        elsif ( ( $options{StandardTime} ) and ( $time_value ) ) {
#            $workbook{$table}[$row][$col] = $time_value;
#            $time_value = '';
#        }
## join cell contents and assign to current workbook cell
#        else {
#            $workbook{$table}[$row][$col] = @cell ? join $options{ReplaceNewlineWith} || "",
#                map { defined($_) ? $_ : '' } @cell : undef;
#        }
## repeat current cell, if necessary
#        for (2..$repeat_cells) {
#            $col++;
#            $workbook{$table}[$row][$col] = $workbook{$table}[$row][$col - 1];
#        }
## reset cell and paragraph values to default for next cell
#        @cell = ();
#        $repeat_cells = 1;
#        $text_p = -1;
#    }
#
#=cut

};

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

1;
