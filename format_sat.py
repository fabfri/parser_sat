#!/usr/bin/env python3
import sys
import argparse

from signal import signal, SIGPIPE, SIG_DFL

# Ignore SIGPIPE and don't throw exception
# BrokenPipeError: [Errno 32] Broken pipe on it
signal(SIGPIPE, SIG_DFL)

db = []
outfile = None

def usage():
  print('In usage')

def oprint(line):
  if outfile != None:
    of.write(line + '\n')
  else:
    print(line)

def parseInput(inputFile):
  firstLine = True
  
  for line in open(inputFile, encoding='iso-8859-1'):
    if firstLine:
      firstLine = False
      continue
    
    fields = {}
    line = line.strip()
    tokens = line.split(';')
    fields['cognome'] = tokens[3]
    fields['nome'] = tokens[4]
    fields['codiceFiscale'] = tokens[5]
    fields['dataNascita'] = tokens[6]
    fields['email'] = tokens[49]
    fields['categoria'] = tokens[10]
    fields['anni iscrizione'] = tokens[11]
    fields['provincia'] = tokens[19]
    fields['CAP'] = tokens[20]
    fields['città'] = tokens[17]
    fields['via'] = ' '.join(tokens[21:23])
    fields['numero'] = tokens[45]
    # If numero civico is not empty
    if tokens[23] != '':
      fields['via'] = fields['via'] + ', ' + tokens[23]
    # If frazione is not empty
    if tokens[18] != '':
      fields['frazione'] = tokens[18]
      fields['via'] = fields['via'] + ' FRAZ. ' + fields['frazione']
      
    fields['indirizzo'] =  fields['via'] + ' - ' + fields['CAP'] + ' ' + \
    fields['città'] + ' (' + fields['provincia'] + ')'
  
    # Format fields
    fields['nome'] = fields['nome'].title()
    fields['cognome'] = fields['cognome'].title()
    fields['email'] = fields['email'].lower()
    fields['indirizzo'] = fields['indirizzo'].upper()
    
    db.append(fields)
    
    
#
# Reads SAT format and outputs Skebby import csv format
# $lastname, $firstname, $nickname, $email, $gender, $birthday, $address, 
# $zipcode, $city, $state, $note, $custom1, $custom2, $custom3, $phonenumber, 
# $groups
# #3,#4,#5,#49,,#6,#18-#23,#20,#17,,,,,,#45,Soci
#  
def sat2skebby():
  print('Print skebby format', end='')
  if outfile is not None:
    print(' on file %s' % outfile)
  else:
    print()
  
  header = 'lastname;firstname;nickname;email;gender;birthday;address;\
  zipcode;city;state;note;custom1;custom2;custom3;phonenumber;groups'
  
  oprint(header)
  #initialize record
  skebbyRecord = [''] * 16

  for record in db:
    skebbyRecord[0] = record['cognome']
    skebbyRecord[1] = record['nome']
    skebbyRecord[2] = record['codiceFiscale']
    skebbyRecord[3] = record['email']
    skebbyRecord[4] = ''
    skebbyRecord[5] = record['dataNascita']
    skebbyRecord[6] = record['via']
    skebbyRecord[7] = record['CAP']
    skebbyRecord[8] = record['città']
    skebbyRecord[9] = record['provincia']
    skebbyRecord[10] = ''
    skebbyRecord[11] = ''
    skebbyRecord[12] = ''
    skebbyRecord[13] = ''
    skebbyRecord[14] = record['numero']
    skebbyRecord[15] = 'Soci'
    line = ';'.join(skebbyRecord)

    if strict:
      if reverse:
        if record['numero'] == '': # salto i soci con indirizzo
          oprint(line)
      elif record['numero'] != '':
        oprint(line)
    else:
      oprint(line)
    
  #outfile.write(nline)
  #print(nline)
  
# 
# Reads SAT format and output simple csv format with basic information
# lastname, firstname, birthday, email, phone, address, city
#
def sat2easy():
  print('Print readable format', end='')
  if outfile is not None:
    print(' on file %s' % outfile)
  else:
    print()
  
  header = 'cognome,nome,data nascita,email,numero,indirizzo'

  oprint(header)
  #initialize record
  easyRecord = [''] * 6

  for record in db:
    easyRecord[0] = record['cognome']
    easyRecord[1] = record['nome']
    easyRecord[2] = record['dataNascita']
    easyRecord[3] = record['email']
    easyRecord[4] = record['numero']
    easyRecord[5] = record['indirizzo']

    line = ','.join(easyRecord)
  
    if strict:
      if reverse:
        if record['numero'] != '' or record['email'] != '': # salto i soci con indirizzo
          oprint(line)
      elif record['numero'] == '' and record['email'] == '':
        oprint(line)
    else:
      oprint(line)
# 
# Reads SAT format and outputs simple csv format with Gmail contacts format 
# $lastname, $firstname, $birthday, $email, $address, $city
# $Name, $GivenName, $AdditionalName, $FamilyName, $YomiName, $GivenNameYomi, $AdditionalNameYomi, $FamilyNameYomi, $NamePrefix, $NameSuffix, $Initials, $Nickname, $ShortName, $MaidenName, $Birthday, $Gender, $Location, $BillingInformation, $DirectoryServer, $Mileage, $Occupation, $Hobby, $Sensitivity, $Priority, $Subject, $Notes, $GroupMembership, $Email1Type, $Email1Value, $Email2Type, $Email2Value, $Phone1Type, $Phone1Value, $Phone2Type, $Phone2Value
# #4#3,#4,,#3,,,,,,,,#5,,,,,#20,,,,,,#49,,,,#45,,,
# 
def sat2gmail():
  print('Print gmail format', end='')
  if outfile is not None:
    print(' on file %s' % outfile)
  else:
    print()

  header = 'Name,Given Name,Additional Name,Family Name,Yomi Name,\
  Given Name Yomi,Additional Name Yomi,Family Name Yomi,Name Prefix,\
  Name Suffix,Initials,Nickname,Short Name,Maiden Name,Birthday,Gender,\
  Location,Billing Information,Directory Server,Mileage,Occupation,Hobby,\
  Sensitivity,Priority,Subject,Notes,Group Membership,E-mail 1 - Type,\
  E-mail 1 - Value,E-mail 2 - Type,E-mail 2 - Value,Phone 1 - Type,\
  Phone 1 - Value,Phone 2 - Type,Phone 2 - Value'
  oprint(header)
  #initialize record
  gmailRecord = [''] * 35
  
  for record in db:
    gmailRecord[0] = record['nome'] + ' ' + record['cognome']
    gmailRecord[1] = record['nome']
    gmailRecord[2] = ''
    gmailRecord[3] = record['cognome']
    gmailRecord[4] = ''
    gmailRecord[5] = ''
    gmailRecord[6] = ''
    gmailRecord[7] = ''
    gmailRecord[8] = ''
    gmailRecord[9] = ''
    gmailRecord[10] = ''
    gmailRecord[11] = record['codiceFiscale']
    gmailRecord[12] =  ''
    gmailRecord[13] = ''
    gmailRecord[14] = record['dataNascita']
    gmailRecord[15] = ''
    gmailRecord[16] = record['indirizzo']
    gmailRecord[17] = ''
    gmailRecord[18] = ''
    gmailRecord[19] = ''
    gmailRecord[20] = ''
    gmailRecord[21] = ''
    gmailRecord[22] = ''
    gmailRecord[23] = ''
    gmailRecord[24] = ''
    gmailRecord[25] = ''
    gmailRecord[26] = '* My Contacts ::: Soci'
    gmailRecord[27] = '* Other'
    gmailRecord[28] = record['email']
    gmailRecord[29] = ''
    gmailRecord[30] = ''
    gmailRecord[31] = 'Home'
    gmailRecord[32] = record['numero']
    gmailRecord[33] = ''
    gmailRecord[34] = ''
    line = ','.join(gmailRecord)
    if strict:
      if reverse:
        if record['email'] == '': # salto i soci con indirizzo email
          oprint(line)
      elif record['email'] != '':
        oprint(line) 
    else:
      oprint(line)
    

def main():
  parser = argparse.ArgumentParser(description='''Parses SAT soci\'s file and
    exports different CSV formats''')
  parser.add_argument('inputFile', metavar='FILE', nargs=1, \
                    help='SAT format input file')
  parser.add_argument('-m','--skebby', action='store_true',
                    help='Outputs Skebby SMS output format')
  parser.add_argument('-g','--gmail', action='store_true', \
                    help='Outputs Gmail Contacts output format')
  parser.add_argument('-l','--letters', action='store_true', \
          help='Outputs CSV with basic information')
  parser.add_argument('-s','--strict', action='store_true', \
                    help='''Prints only records strictly valid for the format 
                    specified (e.g. If "--msg" format is specified and the 
                    record does not contain number, the record is not printed.
                    ''')
  parser.add_argument('-r','--reverse', action='store_true', \
                    help='''In conjunction with "--strict" prints only records 
                    without the field required.
                    ''')
  parser.add_argument('-o','--outfile', metavar='FILE', nargs=1, \
                    help='''Instead of printing to stdout writes output to FILE.
                    ''')

  args = parser.parse_args()
  
  global strict
  global reverse
  strict = args.strict
  reverse = args.reverse

  if not (args.gmail or args.letters or args.skebby):
    print('Specify either one of -m, -g, or -l options.')
    parser.print_usage()
    sys.exit(1)
  
  if args.reverse and not args.strict:
    print('You must specify --reverse option only with --strict option.')
    parser.print_usage()
    sys.exit(1)

  if args.inputFile == '':
    print('Missing argument.')
    parser.print_usage()
    sys.exit(1)

  max_one_is_true = (not (args.gmail and args.letters)) and (not (args.gmail \
    and args.skebby)) and (not (args.letters and args.skebby))
  
  if not max_one_is_true:
    print('Specify only one flag.')
    parser.print_usage()
    sys.exit(1)

  if args.outfile != None:
    global outfile
    outfile = args.outfile.pop()
    try:
      global of
      of = open(outfile,'w', encoding='utf-8')
    except IOError as e:
      print('Could not open file %s for writing: %s' % outfile, e.strerror)

  parseInput(args.inputFile[0])

  if args.gmail:
    sat2gmail()
  if args.letters:
    sat2easy()
  if args.skebby:
    sat2skebby()

  if outfile is not None:
    of.close()

if __name__ == "__main__":
  main()
