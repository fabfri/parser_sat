#!/usr/bin/env python3
import sys
import argparse
import xlrd
from datetime import datetime, date

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
    
    # 0  Codice
    # 1  Categoria
    # 2  Cognome
    # 3  Nome
    # 4  Codice fiscale
    # 5  Data nascita
    # 6  Anno rinnovo
    # 7  Data rinnovo
    # 8  N. anni tess.
    # 9  Email
    # 10 Email 2
    # 11 Tel
    # 12 Tel 2

    this_year = int(date.today().strftime('%Y'))
    firstLine = True
 
    book = xlrd.open_workbook(filename=inputFile)
    sheet = book.sheet_by_index(0)
    n_rows = sheet.nrows
    for i in range(0, n_rows):
        if firstLine:
            firstLine = False
            continue
        line = sheet.row_slice(rowx=i, start_colx=0, end_colx=sheet.ncols)
        if int(line[6].value) < this_year:
            continue
    
        fields = {}
        fields['categoria'] = line[1].value
        fields['cognome'] = line[2].value
        fields['nome'] = line[3].value
        fields['codice_fiscale'] = line[4].value
        fields['data_nascita'] = datetime(*xlrd.xldate_as_tuple(line[5].value, 
            book.datemode)).strftime('%d/%m/%Y')
        fields['anno_rinnovo'] = line[6].value
        fields['data_rinnovo'] = line[7].value
        fields['n_anni_tess'] = line[8].value
        fields['email'] = line[9].value
        fields['tel'] = line[11].value
  
        # Format fields
        fields['nome'] = fields['nome'].title()
        fields['cognome'] = fields['cognome'].title()
        fields['email'] = fields['email'].lower()
        
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
    skebbyRecord[2] = record['codice_fiscale']
    skebbyRecord[3] = record['email']
    skebbyRecord[4] = ''
    skebbyRecord[5] = record['data_nascita']
    #skebbyRecord[6] = record['via']
    skebbyRecord[6] = ''
    #skebbyRecord[7] = record['CAP']
    skebbyRecord[7] = ''
    #skebbyRecord[8] = record['città']
    skebbyRecord[8] = ''
    #skebbyRecord[9] = record['provincia']
    skebbyRecord[9] = ''
    skebbyRecord[10] = ''
    skebbyRecord[11] = ''
    skebbyRecord[12] = ''
    skebbyRecord[13] = ''
    skebbyRecord[14] = record['tel']
    skebbyRecord[15] = 'Soci'
    line = ';'.join(skebbyRecord)

    if strict:
      if reverse:
        if record['tel'] == '': # salto i soci con indirizzo
          oprint(line)
      elif record['tel'] != '':
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
    easyRecord[2] = record['data_nascita']
    easyRecord[3] = record['email']
    easyRecord[4] = record['tel']
    #easyRecord[5] = record['indirizzo']
    easyRecord[5] = ''

    line = ','.join(easyRecord)
  
    if strict:
      if reverse:
        if record['tel'] != '' or record['email'] != '': # salto i soci con indirizzo
          oprint(line)
      elif record['tel'] == '' and record['email'] == '':
        oprint(line)
    else:
      oprint(line)
# 
# Reads SAT format and outputs simple csv format with Gmail contacts format 
# $lastname, $firstname, $birthday, $email, $address, $city
# Name,Given Name,Additional Name,Family Name,Yomi Name,Given Name Yomi,Additional Name Yomi,Family Name Yomi,Name Prefix,Name Suffix,Initials,Nickname,Short Name,Maiden Name,Birthday,Gender,Location,Billing Information,Directory Server,Mileage,Occupation,Hobby,Sensitivity,Priority,Subject,Notes,Language,Photo,Group Membership,E-mail 1 - Type,E-mail 1 - Value,E-mail 2 - Type,E-mail 2 - Value,Phone 1 - Type,Phone 1 - Value
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
Sensitivity,Priority,Subject,Notes,Language,Photo,Group Membership,E-mail 1 - Type,\
E-mail 1 - Value,E-mail 2 - Type,E-mail 2 - Value,Phone 1 - Type,\
Phone 1 - Value'
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
    gmailRecord[11] = record['codice_fiscale']
    gmailRecord[12] =  ''
    gmailRecord[13] = ''
    gmailRecord[14] = record['data_nascita']
    gmailRecord[15] = ''
    #gmailRecord[16] = record['indirizzo']
    gmailRecord[16] = ''
    gmailRecord[17] = ''
    gmailRecord[18] = ''
    gmailRecord[19] = ''
    gmailRecord[20] = ''
    gmailRecord[21] = ''
    gmailRecord[22] = ''
    gmailRecord[23] = ''
    gmailRecord[24] = ''
    gmailRecord[25] = ''
    gmailRecord[26] = ''
    gmailRecord[27] = ''
    gmailRecord[28] = '* My Contacts ::: Soci'
    gmailRecord[29] = '* Other'
    gmailRecord[30] = record['email']
    gmailRecord[31] = ''
    gmailRecord[32] = ''
    gmailRecord[33] = 'Home'
    gmailRecord[34] = record['tel']
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
