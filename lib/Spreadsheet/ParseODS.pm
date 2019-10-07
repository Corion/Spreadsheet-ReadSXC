package Spreadsheet::ParseODS;
use strict;
use warnings;

use Archive::Zip ':ERROR_CODES';
use Moo 2;
use XML::Parser;
use XML::Twig;
use Carp qw(croak);

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
        XML::Twig->new(
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
    my $table = "";
    my $row = -1;
    my $col = -1;
    my $text_p = -1;
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
    my %options = ();

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

    $handlers{ "table:table-row" } = sub {
        my( $element ) = @_;
# reset table, column, and row values to default for this table
        $row = -1;
        $max_datarow = -1;
        $max_datacol = -1;
        $table = "";
        $col_count = -1;
        @hidden_cols = ();

        $table = $element->att('table:name');
        $row_hidden = $element->att( 'table:visibility';

        for my $row ($table->find('
        # XXX Collect table rows here
            # XXX Collect cell and cell values here

# increase row count
        $row++;
# if row is hidden, set $row_hidden for later use
# if number-rows-repeated is set, set $repeat_rows value accordingly for later use
        my $repeat = $element->att('table:number-rows-repeated');
        if( defined $repeat ) {
            $repeat_rows = $repeat;
        };

# decrease $max_datacol if hidden columns within range
        if ( ( ! $options{NoTruncate} ) and ( $options{DropHiddenColumns} ) ) {
            for ( 1..scalar grep { $_ <= $max_datacol } @hidden_cols ) {
                $max_datacol--;
            }
        }
# truncate table to $max_datarow and $max_datacol
        if ( ! $options{NoTruncate} ) {
            $#{$workbook{$table}} = $max_datarow;
            foreach ( @{$workbook{$table}} ) {
                $#{$_} = $max_datacol;
            }
        }
# set up alternative data structure
        if ( $options{OrderBySheet} ) {
            push @worksheets, (
                {
                    label   => $table,
                    data    => \@{$workbook{$table}},
                }
            );
        };
    };

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
#
#    $handlers{ "table:table" } = sub {
#        my( $element ) = @_;
## get name of current table
#    }
}

    $p->setTwigHandlers( \%handlers );

    my $ref = ref($source);
    my $xml;
    my $method = 'parse';

    if( ! $ref ) {
        # Specified by filename .
        $workbook{File} = $source;

        if( $source =~ m!\.xml$i ) {
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
            $workbook{File} = undef;

        } elsif ( $ref eq 'ARRAY' ) {
            # Specified by file content
            $workbook{File} = undef;
            $xml = join( '', @$source );

        } else {
             # Assume filehandle
             # Kick off XML::Twig from Filehandle
             $xml = $self->_open_sxc_fh( $source );
         }
    }

    $p->parse( $xml );
};

sub _fetch_cell_value {
    my( $self, $element ) = @_;
    my $className = $element->tag;

    $element->value()

=begin later
    if ( ( $tag eq "table:table-cell" ) or ( $tag eq "table:covered-table-cell" ) ) {
# assign currency, date or time value to current workbook cell if requested
        if ( ( $self->StandardCurrency ) and ( length( $currency_value ) ) ) {
            $workbook{$table}[$row][$col] = $currency_value;
            $currency_value = '';

        }
        elsif ( ( $options{StandardDate} ) and ( $date_value ) ) {
            $workbook{$table}[$row][$col] = $date_value;
            $date_value = '';
        }
        elsif ( ( $options{StandardTime} ) and ( $time_value ) ) {
            $workbook{$table}[$row][$col] = $time_value;
            $time_value = '';
        }
# join cell contents and assign to current workbook cell
        else {
            $workbook{$table}[$row][$col] = @cell ? join $options{ReplaceNewlineWith} || "",
                map { defined($_) ? $_ : '' } @cell : undef;
        }
# repeat current cell, if necessary
        for (2..$repeat_cells) {
            $col++;
            $workbook{$table}[$row][$col] = $workbook{$table}[$row][$col - 1];
        }
# reset cell and paragraph values to default for next cell
        @cell = ();
        $repeat_cells = 1;
        $text_p = -1;
    }

=cut
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
    my ($fh, $options_ref) = @_;
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
