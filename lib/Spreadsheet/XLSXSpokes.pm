package Spreadsheet::XLSXSpokes;
use Exporter;
use Carp;
use Archive::Zip qw/:ERROR_CODES/;
use Data::Dumper;
use XML::Bare qw/xval forcearray/;
use Time::JulianDay;
@ISA = qw(Exporter);
@EXPORT_OK = qw();
use strict;
use warnings;
$Spreadsheet::XLSXSpokes::VERSION = '0.01';
sub new { my $pkg = shift; return bless { @_ }, $pkg; }

sub read {
  my $self = shift;
  my %args = ( @_ );
  
  my ( $zip, $files ) = $self->read_zip( $args{'file'} );
  
  # Ensure that all the basic files we need exist in the XLSX file
  my @ensure = qw|xl/worksheets/sheet1.xml xl/workbook.xml xl/sharedStrings.xml xl/styles.xml|;
  for my $file ( @ensure ) {
    if( !$files->{ $file } ) {
      die "$file is missing";
    }
  }
  
  my $data = $self->load_xml_files( $zip, $files );
  my $sheets = $self->{'sheets'} = $self->parse_data( $data );
  #print Dumper( $sheets );
}

sub write_xml {
  my $self = shift;
  my %args = ( @_ );
  open( my $file, ">".$args{'file'} );
  my $xml = $self->convert_sheets_to_xml( $self->{'sheets'} );
  print Dumper( $xml );
  print $file XML::Bare::Object::xml( 0, $xml );
  close( $file );
}

sub convert_sheets_to_xml {
  my ( $self, $sheets ) = @_;
  my $xml = {};
  for my $sheet ( @$sheets ) {
    my $name = $sheet->{'name'};
    $xml->{ $name } = { row => $self->convert_sheet_to_xml( $sheet ) };
  }
  return $xml;
}

sub convert_sheet_to_xml {
  my ( $self, $sheet ) = @_;
  my $rows = $sheet->{'rows'};
  my $titles = shift @$rows;
  $titles = $titles->{'cells'};
  my @xrows;
  for my $rawrow ( @$rows ) {
    my $cells = $rawrow->{'cells'};
    my $i = 0;
    my $xrow = {};
    for my $cell ( @$cells ) {
      my $cellname = $titles->[ $i ];
      $xrow->{ $cellname } = { value => $cell };
      $i++;
    }
    push( @xrows, $xrow );
  }
  return \@xrows;
}


sub read_zip {
  my ( $self, $file ) = @_;
  my $zip = Archive::Zip->new();
  if( $zip->read( $file ) != AZ_OK ) { die 'read error'; }
  
  my @files = $zip->memberNames();
  my %hash = map { $_ => 1 } @files;
  
  return ( $zip, \%hash );
}

sub load_xml_files {
  my ( $self, $zip, $files ) = @_;
  
  my $workbook = $zip->contents( { memberOrZipName => "xl/workbook.xml" } );
  my $strings  = $zip->contents( { memberOrZipName => "xl/sharedStrings.xml" } );
  my $styles   = $zip->contents( { memberOrZipName => "xl/styles.xml" } );
  my @sheetdata;
  for( my $i=0;$i<20;$i++ ) {
    my $internal = "xl/worksheets/sheet$i.xml";
    if( $files->{ $internal } ) {
      my $data = $zip->contents( { memberOrZipName => $internal } );
      push( @sheetdata, $data );
    }
  }
  
  return $self->{'data'} = {
    workbook => $workbook,
    strings => $strings,
    sheetdata => \@sheetdata,
    styles => $styles
  };
}

sub parse_data {
  my ( $self, $data ) = @_;
  
  my ( $ob, $workbook ) = XML::Bare->new( text => $data->{'workbook'} );
  my $sheets = forcearray( $workbook->{'workbook'}{'sheets'}{'sheet'} );
  
  my ( $ob1, $strings ) = XML::Bare->new( text => $data->{'strings'} );
  my $string_data = forcearray( $strings->{'sst'}{'si'} );
  
  my ( $ob3, $styles ) = XML::Bare->new( text => $data->{'styles'} );
  my $cell_styles = forcearray( $styles->{'styleSheet'}{'cellXfs'}{'xf'} );
  
  # Parse the sheet data and store it in @sheetdata
  my $sheetdata_raw = $data->{'sheetdata'};
  my @sheetdata;
  for my $raw ( @$sheetdata_raw ) {
    my ( $ob2, $data_xml ) = XML::Bare->new( text => $raw );
    push( @sheetdata, $data_xml );
  }
  
  my $parsed = $self->parse_sheets( \@sheetdata, $sheets, $string_data, $cell_styles );
  return $parsed;
}

sub parse_sheets {
  my ( $self, $sheet_datasets, $rawsheets, $string_data, $cell_styles ) = @_;
  
  my @sheets;
  my $sheeti = 0;
  for my $rawsheet ( @$rawsheets ) {
    my %sheet = (
      name => xval( $rawsheet->{'name'} )
    );
    
    my $sheet_data = $sheet_datasets->[ $sheeti ];
    my $raw_rows = forcearray( $sheet_data->{'worksheet'}{'sheetData'}{'row'} );
    $sheet{'numrows'} = scalar @$raw_rows;
    
    my @rows;
    for my $raw_row ( @$raw_rows ) {
      my $raw_cells = forcearray( $raw_row->{'c'} );
      my @cells;
      my $empty = 1;
      for my $cell ( @$raw_cells ) {
        my $val = xval $cell->{'v'};
        if( $cell->{'t'} ) {
          my $t = xval $cell->{'t'};
          if( $t eq 's' ) {
            $val = xval $string_data->[ $val ]{'t'};
            $val =~ s/^\s+|\s+$//g;
          }
        }
        if( $cell->{'s'} ) {
          my $s = xval $cell->{'s'};
          my $style = $cell_styles->[ $s ];
          my $format_id = xval $style->{'numFmtId'};
          # Consider using CPAN Spreadsheet::ParseExcel::Utility
          if( $format_id >= 14 && $format_id <= 17 ) {
            $val = decode_date( $val );
          }
        }
        $empty = 0 if( $val ne '' );
        push( @cells, $val );
      }
      push( @rows, { cells => \@cells } ) if( !$empty );
    }
    $sheet{'rows'} = \@rows;
    
    push( @sheets, \%sheet );
    
    $sheeti++;
  }
  
  return \@sheets;
}

sub decode_date {
  my $jd = shift + julian_day(1900, 1, 0) - 1;
  my ($year, $month, $day) = inverse_julian_day($jd);
  return "$month-$day-$year";
}

1;