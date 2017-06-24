'''
Denne filen inneholder klassen og koden for å importere informasjon fra søknader
på pdf. Det første som er her nå er bare kopiert fra tidligere program.
'''


# -*- coding: utf-8 -*-
from __future__ import print_function
from builtins import input


import datetime
import sys, os
import PyPDF2 as p
reload(sys)
sys.setdefaultencoding('utf8')

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice

from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams

from cStringIO import StringIO

import xlwt

def pdf_to_text(path):
    # PDFMiner boilerplate
    rsrcmgr = PDFResourceManager()
    sio = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, sio, codec=codec, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    # Extract text
    fp = file(path, 'rb')
    for page in PDFPage.get_pages(fp):
        interpreter.process_page(page)
    fp.close()

    # Get text from StringIO
    text = sio.getvalue()

    # Cleanup
    device.close()
    sio.close()
    text.encode('utf-8')
    return text

def find_orgdata_name(txt):
	# Finding the Organiastion name
	a = txt.find('OrganisasjonsdataNavn: ')
	b = txt.find('Nettsted: ')
	l = len('OrganisasjonsdataNavn: ')
	return txt[a+l:b]

def find_desired_amount(txt):
	# Finding the desired amount
	a = txt.find('p: ')
	b = txt.find('OrganisasjonsdataNavn:')
	l = len('p: ')
	return txt[a+l:b]

def find_leder_name(txt):
	# Finding the name of the leader of the organization
	a = txt.find('lederNavn: ')
	b = txt.find('E-post: ')
	l = len('lederNavn: ')
	return txt[a+l:b]

def find_leder_email(txt):
	# Finding the email of the leader of the organization
	a = txt.find('E-post: ')
	b = txt.find('Telefon:')
	l = len('E-post: ')
	return txt[a+l:b]

def find_leader_phone(txt):
	# Finding the phone number of the leader of the organization
	a = txt.find('Telefon: ')
	b = txt.rfind('Personlig dat')
	l = len('Telefon: ')
	return txt[a+l:b]

def find_kontakt_name(txt):
	# Finding the name of the contact person of the organization
	a = txt.find('knadNavn: ')
	b = txt.rfind('E-post: ')
	l = len('knadNavn: ')
	return txt[a+l:b]

def find_kontakt_email(txt):
	# Finding the email of the contact person of the organization
	a = txt.rfind('E-post: ')
	b = txt.rfind('Telefon:')
	l = len('E-post: ')
	return txt[a+l:b]

def find_kontakt_phone(txt):
	# Finding the phone number of the contact person of the organization
	a = txt.rfind('Telefon: ')
	b = txt.rfind('Bank/kontodataKontoha')
	l = len('Telefon: ')
	return txt[a+l:b]

def find_orgdata_webpage(txt):
	# Finding webpage of the organization
	a = txt.rfind('Nettsted: ')
	b = txt.rfind('Organisasjonsnummer')
	l = len('Nettsted: ')
	return txt[a+l:b]

def find_orgdata_orgNum(txt):
	# Finding organisation number of the organization
	a = txt.find('Organisasjonsnummer: ')
	b = txt.find('Postadresse:')
	l = len('Organisasjonsnummer: ')
	return txt[a+l:b]

def find_orgdata_postal(txt):
	# Finding postal adress of the organization
	a = txt.find('Postadresse: ')
	b = txt.find('Personlig dat')
	l = len('Postadresse: ')
	return txt[a+l:b]

def find_bank_name(txt):
	# Finding name of the accoundtholder
	a = txt.find('kontodataKontohaver: ')
	b = txt.find('Kontonummer: ')
	l = len('kontodataKontohaver: ')
	return txt[a+l:b]

def find_bank_accountNum(txt):
	# Finding accountnuber of the organization
	a = txt.find('Kontonummer: ')
	b = txt.find('Medlemsavgif')
	l = len('Kontonummer: ')
	return txt[a+l:b]

def find_memFee_amount(txt):
	# Finding membership fee
	a = txt.find('p i kr,-: ')
	b = txt.rfind('Bel')
	l = len('p i kr,-: ')
	return txt[a+l:b]

def find_memFee_period(txt):
	# Finding membership fee periodization
	a = txt.find('betales: ')
	b = txt.rfind('Medlemsdata per dags datoTotalt antall med')
	l = len('betales: ')
	return txt[a+l:b]

def find_memberData_total(txt):
	# Finding total members of the organisation
	a = txt.find('Totalt antall medlemmer: ')
	b = txt.find('Betalende studentmedlemmer: ')
	l = len('Totalt antall medlemmer: ')
	return txt[a+l:b]

def find_memberData_payStudent(txt):
	# Finding total number of paying student members of the organisation
	a = txt.find('Betalende studentmedlemmer: ')
	b = txt.find('Andre betalende medlemmer: ')
	l = len('Betalende studentmedlemmer: ')
	return txt[a+l:b]

def find_memberData_nonPayStudents(txt)
	# Finding total number of paying student members of the organisation
	a = txt.find('Betalende studentmedlemmer: ')
	b = txt.find('Andre betalende medlemmer: ')
	l = len('Betalende studentmedlemmer: ')
	return txt[a+l:b]

def find_memberData_payOther(txt):
	# Finding total number of non student paying members of the organisation
	a = txt.find('Andre betalende medlemmer: ')
	b = txt.find('Antall andre medlemmer: ')
	l = len('Andre betalende medlemmer: ')
	return txt[a+l:b]

def find_memberData_other(txt):
	# Finding total number of other members of the organisation
	a = txt.find('Antall andre medlemmer: ')
	b = txt.find('Tidligere st')
	l = len('Antall andre medlemmer: ')
	return txt[a+l:b]

def find_tidligereStotte(txt):
	# Finding out if the organization hast recived support before
	a = txt.find('te tidligere?: ')
	b = txt.rfind('S')
	l = len('te tidligere?: ')
	return txt[a+l:b]

def find_dateGenerated(txt):
	# Finding out when the application was generated
	a = txt.find('le generert ')
	b = txt.find('')
	l = len('le generert ')
	return txt[a+l:b]

def exstract_info(txt):
	# Exstacts the information from the pdf and returns a dict with all information

	application = {}
	application['orgData_name'] = find_orgdata_name(txt)					#0
	application['orgData_desiredAmount'] = find_desired_amount(txt)		#1
	application['orgData_webpage'] = find_orgdata_webpage(txt)				#2
	application['orgData_orgNum'] = find_orgdata_orgNum(txt)				#3
	application['orgData_postal'] = find_orgdata_postal(txt)				#4

	application['leader_name'] = find_leder_name(txt)						#5
	application['leader_email'] = find_leder_email(txt)					#6
	application['leader_phone'] = find_leader_phone(txt)					#7

	application['contact_name'] = find_kontakt_name(txt)					#8
	application['contact_email'] = find_kontakt_email(txt)					#9
	application['contact_phone'] = find_kontakt_phone(txt)					#10

	application['bank_name'] = find_bank_name(txt)							#11
	application['bank_accountNum'] = find_bank_accountNum(txt)				#12

	application['memFee_amount'] = find_memFee_amount(txt)					#13
	application['memFee_period'] = find_memFee_period(txt)					#14

	application['memberData_total'] = find_memberData_total(txt)			#15
	application['memberData_payStudent'] = find_memberData_payStudent(txt)	#16
	application['memberData_payOther'] = find_memberData_payOther(txt)		#17
	application['memberData_other'] = find_memberData_other(txt)			#18

	application['tigligereStotte'] = find_tidligereStotte(txt)				#19

	application['dateGenerated'] = find_dateGenerated(txt)					#20

	return application

def self_path():
	return os.path.dirname(os.path.realpath(__file__))

def initialize_sreadsheet():
	sheet_name = 'applications'
	workbook = xlwt.Workbook(encoding="utf-8")
	sheet = workbook.add_sheet(sheet_name)
	row_number = 0

	col1_name = 'Organisasjonsnavn'
	col2_name = 'Onsket sponsorbelp'
	col3_name = 'Nettside'
	col4_name = 'Organisasjonsnummer'
	col5_name = 'Postadresse'
	col6_name = 'Leders navn'
	col7_name = 'Leders epost'
	col8_name = 'Leders telefonnummer'
	col9_name = 'Kontakt navn'
	col10_name = 'Kontakt email'
	col11_name = 'Kontakt telefonnummer'
	col12_name = 'Kontohaver'
	col13_name = 'Konto nummer'
	col14_name = 'Medlemsavgift'
	col15_name = 'Medlems periodisering'
	col16_name = 'Totalt antall medlemmer'
	col17_name = 'Betalende studentmedlemmer'
	col18_name = 'Andre betalende medlemmer'
	col19_name = 'Andre medlemmer'
	col20_name = 'Fatt stotte for?'
	col21_name = 'Soknad gennerert'


	sheet.write(row_number, 0, col1_name)
	sheet.write(row_number, 1, col2_name)
	sheet.write(row_number, 2, col3_name)
	sheet.write(row_number, 3, col4_name)
	sheet.write(row_number, 4, col5_name)
	sheet.write(row_number, 5, col6_name)
	sheet.write(row_number, 6, col7_name)
	sheet.write(row_number, 7, col8_name)
	sheet.write(row_number, 8, col9_name)
	sheet.write(row_number, 9, col10_name)
	sheet.write(row_number, 10, col11_name)
	sheet.write(row_number, 11, col12_name)
	sheet.write(row_number, 12, col13_name)
	sheet.write(row_number, 13, col14_name)
	sheet.write(row_number, 14, col15_name)
	sheet.write(row_number, 15, col16_name)
	sheet.write(row_number, 16, col17_name)
	sheet.write(row_number, 17, col18_name)
	sheet.write(row_number, 18, col19_name)
	sheet.write(row_number, 19, col20_name)
	sheet.write(row_number, 20, col21_name)

	return workbook, sheet, row_number

def print_line_to_spreadsheet(sheet, row_number, application, dirname):
	row_number = row_number + 1
	error_message = '\t There has been an Encoding error'
	error_print = 'encodingError'

	if len(application['orgData_name']) > 100:
		sheet.write(row_number, 0, dirname)
		sheet.write(row_number, 1, 'Engelsk soknad ---- fyll ut selv')
	else:
		try:
			sheet.write(row_number, 0, application['orgData_name'])
		except:
			print(error_message)
			sheet.write(row_number, 0, error_print)

		try:
			sheet.write(row_number, 1, application['orgData_desiredAmount'])
		except:
			print(error_message)


		try:
			sheet.write(row_number, 2, application['orgData_webpage'])
		except:
			print(error_message)
			sheet.write(row_number, 2, error_print)

		try:
			sheet.write(row_number, 3, application['orgData_orgNum'])
		except:
			print(error_message)
			sheet.write(row_number, 3, error_print)

		try:
			sheet.write(row_number, 4, application['orgData_postal'])
		except:
			print(error_message)
			sheet.write(row_number, 4, error_print)

		try:
			sheet.write(row_number, 5, application['leader_name'])
		except:
			print(error_message)
			sheet.write(row_number, 5, error_print)

		try:
			sheet.write(row_number, 6, application['leader_email'])
		except:
			print(error_message)
			sheet.write(row_number, 6, error_print)

		try:
			sheet.write(row_number, 7, application['leader_phone'])
		except:
			print(error_message)
			sheet.write(row_number, 7, error_print)

		try:
			sheet.write(row_number, 8, application['contact_name'])
		except:
			print(error_message)
			sheet.write(row_number, 8, error_print)

		try:
			sheet.write(row_number, 9, application['contact_email'])
		except:
			print(error_message)
			sheet.write(row_number, 9, error_print)

		try:
			sheet.write(row_number, 10, application['contact_phone'])
		except:
			print(error_message)
			sheet.write(row_number, 10, error_print)

		try:
			sheet.write(row_number, 11, application['bank_name'])
		except:
			print(error_message)
			sheet.write(row_number, 11, error_print)

		try:
			sheet.write(row_number, 12, application['bank_accountNum'])
		except:
			print(error_message)
			sheet.write(row_number, 12, error_print)

		try:
			sheet.write(row_number, 13, application['memFee_amount'])
		except:
			print(error_message)
			sheet.write(row_number, 13, error_print)

		try:
			sheet.write(row_number, 14, application['memFee_period'])
		except:
			print(error_message)
			sheet.write(row_number, 14, error_print)

		try:
			sheet.write(row_number, 15, application['memberData_total'])
		except:
			print(error_message)
			sheet.write(row_number, 15, error_print)

		try:
			sheet.write(row_number, 16, application['memberData_payStudent'])
		except:
			print(error_message)
			sheet.write(row_number, 16, error_print)

		try:
			sheet.write(row_number, 17, application['memberData_payOther'])
		except:
			print(error_message)
			sheet.write(row_number, 17, error_print)

		try:
			sheet.write(row_number, 18, application['memberData_other'])
		except:
			print(error_message)
			sheet.write(row_number, 18, error_print)

		try:
			sheet.write(row_number, 19, application['tigligereStotte'])
		except:
			print(error_message)
			sheet.write(row_number, 19, error_print)

		try:
			sheet.write(row_number, 20, application['dateGenerated'])
		except:
			print(error_message)
			sheet.write(row_number, 20, error_print)

	return row_number

def finilize_spreadsheet(workbook, filename):
	workbook.save(filename)

filename = self_path() + '//test.xls'
workbook, sheet, row_number = initialize_sreadsheet()



for root, dirs, files in os.walk(self_path()):
	if (root != self_path()):
		pdf_files = [f for f in files if f.endswith('pdf')]
		dirname1 = os.path.basename(root)
		print(dirname1)
		for f_name in pdf_files:

			if f_name == 'application.pdf':
				try:
					txt = pdf_to_text(dirname1 +  '//' + 'application.pdf')
					application = exstract_info(txt)
					row_number = print_line_to_spreadsheet(sheet, row_number, application, dirname1)
				except:
					print('Error at ' + dirname1)



finilize_spreadsheet(workbook, filename)
