#!/usr/bin/perl -w
require 5.008;

use strict;
use Spreadsheet::WriteExcel;
use utf8;

my $infile = "";
my $outfile = "";

if( @ARGV == 1 ){
	$outfile = $ARGV[0];
} elsif( @ARGV == 2 ){
	$infile = $ARGV[0];
	$outfile = $ARGV[1];
} else {
	$outfile = pop(@ARGV);
}

if( $outfile !~ m/\.xls$/){$outfile =~ s/$/.xls/ };

my $workbook = Spreadsheet::WriteExcel->new("$outfile");

foreach my $file (@ARGV) {

print "$file\n";
my $worksheet = $workbook->add_worksheet("$file");
   $worksheet->set_column('A:A',10);
my $row=0;

if( "$file" ne "" ){
#open FH, '<:encoding(iso-2022-jp)',$file or die "Couldn't open $file: $!\n";
open FH, '<:encoding(utf8)',$file or die "Couldn't open $file: $!\n";

while (<FH>) {
	chomp;
	my $col=0;
	my @rec = split(/	/);
	foreach my $val(@rec){
		if( $val =~ /^=/ ){ $val = "'"."$val";}
		if( $val =~ /^[０-９]+$/ ){ $val =~ tr/０１２３４５６７８９/0123456789/ ;}
		$worksheet->write($row,$col++,"$val");
	}
	$row++;
}
} else {
binmode STDIN, ':utf8';
binmode STDOUT, ':utf8';
while (<STDIN>) {
	chomp;
	my $col=0;
	my @rec = split(/	/);
	foreach my $val(@rec){
		if( $val =~ /^=/ ){ $val = "'"."$val";}
		if( $val =~ /^[０-９]+$/ ){ $val =~ tr/０１２３４５６７８９/0123456789/ ;}
		$worksheet->write($row,$col++,$val);
	}
	$row++;
}
}
}

__END__
