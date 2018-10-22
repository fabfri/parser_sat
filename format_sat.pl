#!/usr/bin/perl

use strict;
use File::Basename;
use Getopt::Long;

# Allows to print UTF8 characters on STDOUT
binmode(STDOUT, ":utf8");

my $first;
my $letters;
my $gmail;
my $csv;
my $skebby;
my $strict;
my $help;
my $reverse;

sub check_input(){
  GetOptions('gmail' => \$gmail, 'letters' => \$csv, 'msg' => \$skebby, 'strict' => \$strict, 'help' => \$help, 'reverse' => \$reverse);
  
  if($help){
    &usage() and exit(0);
  } 

  unless ($gmail || $csv || $skebby){
    &usage() and die "Missing option.";
  }
  
  if($reverse && !$strict){
    &usage() and die "You must specify --reverse option only with --strict option.";
  }
  
  die "Missing arguments." if (@ARGV ne 1);
  
  my $check = $gmail+$csv+$skebby;
  &usage() and die "Specify only one flag." if($check gt 1);
}

sub usage(){
  my $command =  fileparse($0);
  my $message = "Usage: $command [OPTION]... FILE \
  \t-h, --help\t\tPrints this message. \
  \t-m, --msg\t\tOutputs Skebby SMS output format. \
  \t-g, --gmail\t\tOutputs Gmail Contacts output format. \
  \t-l, --letters\t\tOutputs CSV with basic information. \
  \t-s, --strict\t\tPrints only records strictly valid for the format specified (e.g. If \"--skebby\"
  \t\t\t\tformat is specified and the record does not contain number, the record is not printed. \
  \t-r, --reverse\t\tIn conjunction with \"--strict\" prints only records without the field required.\n";
  print "$message\n";
}

#
# Utility function, removes heading and trailing spaces
#
sub  trim { my $s = shift; $s =~ s/^\s+|\s+$//g; return $s };

#
# Reads SAT format and outputs Skebby import csv format
# $lastname, $firstname, $nickname, $email, $gender, $birthday, $address, $zipcode, $city, $state, $note, $custom1, $custom2, $custom3, $phonenumber, $groups
# #3,#4,#5,#49,,#6,#18-#23,#20,#17,,,,,,#45,Soci
#
sub sat2skebby(){

  if($first){
    print "lastname;firstname;nickname;email;gender;birthday;address;zipcode;city;state;note;custom1;custom2;custom3;phonenumber;groups\n" and $first=0;
    return;
  }

  my @tokens = split/;/ ;
  my ( $lastname, $firstname, $nickname, $email, $gender, $birthday, $address, $zipcode, $city, $state, $note, $custom1, $custom2, $custom3, $phonenumber, $groups) = @tokens[3,4,5,52,0,9,0,23,20,0,0,0,0,0,48];
  $gender = "";
  $address = $tokens[24] . " " . $tokens[25] . " " . $tokens[26] . " - " . $tokens[21] . " (" . $tokens[22] .")";
  $state = "";
  $note = "";
  $custom1 = "";
  $custom2 = "";
  $custom3 = "";
  $groups = "Soci";
  
  $lastname = ucfirst lc $lastname;
  $firstname = ucfirst lc $firstname;

  my @dest = ( $lastname, $firstname, $nickname, $email, $gender, $birthday, $address, $zipcode, $city, $state, $note, $custom1, $custom2, $custom3, $phonenumber, $groups);
  my $string = join(';', @dest);

  if($strict){
    if ($reverse){
      print $string, "\n" if ($phonenumber eq "");
    }
    else{
      print $string, "\n" if ($phonenumber ne "");
    }
  }
  else{
    print $string, "\n"; 
  }
}

# 
# Reads SAT format and output simple csv format with basic information
# $lastname, $firstname, $birthday, $email, $address, $city
#
sub sat2csv(){

  if($first){
    print "cognome;nome;data nascita;email;indirizzo;citta'\n" and $first=0;
    return;
  }
  
  my @tokens = split/;/ ;
  my ( $lastname, $firstname, $nickname, $email, $gender, $birthday, $address, $zipcode, $city, $state, $note, $custom1, $custom2, $custom3, $phonenumber, $groups) = @tokens[3,4,5,52,0,9,0,23,20,0,0,0,0,0,48];
  $address = $tokens[24] . " " . $tokens[25] . " " . $tokens[26] . " - " . $tokens[21] . " (" . $tokens[22] .")";

  $lastname = ucfirst lc $lastname;
  $firstname = ucfirst lc $firstname;

  my @dest = ( $lastname, $firstname, $birthday, $email, $address, $city);
  my $string = join(';', @dest);
  if($strict){
    if ($reverse){
      print $string, "\n" if ($tokens[52] ne "");
    }
    else{
      # salto i soci con indirizzo email
      print $string, "\n" if ($tokens[52] eq "");
    }
  }
  else{
    print $string, "\n"; 
  }
}

#
# Formato per processare la lista dei gestori.
#
sub gestori_skebby(){
  
  #0,#1,,,,,,,,,,,,,#2,Gestori
  my @tokens = split/,/ ;
  foreach my $tok (@tokens){
    chomp $tok;
    $tok = trim($tok);
  }
  my ( $lastname, $firstname, $nickname, $email, $gender, $birthday, $address, $zipcode, $city, $state, $note, $custom1, $custom2, $custom3, $phonenumber, $groups) = @tokens[0,1,0,0,0,0,0,0,0,0,0,0,0,0,2];
  $nickname = "";
  $email = "";
  $gender = "";
  $birthday = "";
  $address = "";
  $zipcode = "";
  $city = "";
  $state = "";
  $note = "";
  $custom1 = "";
  $custom2 = "";
  $custom3 = "";
  $groups = "Gestori";

  my @dest = ( $lastname, $firstname, $nickname, $email, $gender, $birthday, $address, $zipcode, $city, $state, $note, $custom1, $custom2, $custom3, $phonenumber, $groups);
  my $string = join(';', @dest);
  print $string, "\n";
}


# 
# Reads SAT format and outputs simple csv format with Gmail contacts format 
# $lastname, $firstname, $birthday, $email, $address, $city
# $Name, $GivenName, $AdditionalName, $FamilyName, $YomiName, $GivenNameYomi, $AdditionalNameYomi, $FamilyNameYomi, $NamePrefix, $NameSuffix, $Initials, $Nickname, $ShortName, $MaidenName, $Birthday, $Gender, $Location, $BillingInformation, $DirectoryServer, $Mileage, $Occupation, $Hobby, $Sensitivity, $Priority, $Subject, $Notes, $GroupMembership, $Email1Type, $Email1Value, $Email2Type, $Email2Value, $Phone1Type, $Phone1Value, $Phone2Type, $Phone2Value
# #4#3,#4,,#3,,,,,,,,#5,,,,#20,,,,,,#49,,,,#45,,,
# 
sub sat2gmail(){
  if($first){
    print "Name,Given Name,Additional Name,Family Name,Yomi Name,Given Name Yomi,Additional Name Yomi,Family Name Yomi,Name Prefix,Name Suffix,Initials,Nickname,Short Name,Maiden Name,Birthday,Gender,Location,Billing Information,Directory Server,Mileage,Occupation,Hobby,Sensitivity,Priority,Subject,Notes,Group Membership,E-mail 1 - Type,E-mail 1 - Value,E-mail 2 - Type,E-mail 2 - Value,Phone 1 - Type,Phone 1 - Value,Phone 2 - Type,Phone 2 - Value\n" and $first=0;
    return;
  }

  my @tokens = split/;/ ;
  #4#3,#4,,#3,,,,,,,,#5,,,,#20,,,,,,#49,,,,#45,,,
  $tokens[0] = "";
  my ($Name, $GivenName, $AdditionalName, $FamilyName, $YomiName, $GivenNameYomi, $AdditionalNameYomi, $FamilyNameYomi, $NamePrefix, $NameSuffix, $Initials, $Nickname, $ShortName, $MaidenName, $Birthday, $Gender, $Location, $BillingInformation, $DirectoryServer, $Mileage, $Occupation, $Hobby, $Sensitivity, $Priority, $Subject, $Notes, $GroupMembership, $Email1Type, $Email1Value, $Email2Type, $Email2Value, $Phone1Type, $Phone1Value, $Phone2Type, $Phone2Value
  ) = @tokens[0,4,0,3,0,0,0,0,0,0,0,5,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,52,0,0,0,48,0,0];
  $tokens[4] = ucfirst lc $tokens[4];
  $tokens[3] = ucfirst lc $tokens[3];

  $Name = $tokens[4]. " " . $tokens[3];
  $GivenName = $tokens[4];
  $AdditionalName = "";
  $FamilyName = $tokens[3];
  $Location = $tokens[24] . " " . $tokens[25] . " " . $tokens[26] . " - " . $tokens[21] . " (" . $tokens[22] .")";
  $GroupMembership = "* My Contacts ::: Soci";
  $Email1Type = "* Other";
  $Phone1Type = "Home";
  my @dest = ($Name, $GivenName, $AdditionalName, $FamilyName, $YomiName, $GivenNameYomi, $AdditionalNameYomi, $FamilyNameYomi, $NamePrefix, $NameSuffix, $Initials, $Nickname, $ShortName, $MaidenName, $Birthday, $Gender, $Location, $BillingInformation, $DirectoryServer, $Mileage, $Occupation, $Hobby, $Sensitivity, $Priority, $Subject, $Notes, $GroupMembership, $Email1Type, $Email1Value, $Email2Type, $Email2Value, $Phone1Type, $Phone1Value, $Phone2Type, $Phone2Value);
  my $string = join(',', @dest);
 
  if($strict){
    if ($reverse){
      print $string, "\n" if ($Email1Value eq "");
    }
    else{
      print $string, "\n" if ($Email1Value ne "");
    }
  }
  else{
    print $string, "\n"; 
  }
}

&check_input();

my $file = shift;
$first = 1;
open(IN, "<$file") or die "Can't open file $file: $!";

while(<IN>){
  chomp;
 sat2skebby() if($skebby);
 sat2gmail() if($gmail);
 sat2csv() if($csv);
 # gestori_skebby();
 #&usage();
}

