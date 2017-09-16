'''
Denne modulen tar seg av å sende epost med avslag til alle de som får avslag.
Den kommer til å bassere seg på en IMAP/SMTP eller Gmail konto.

NB: Denne er kopiert fra en betaversjon fra en tidligere kode
NB: Dette er 2.7 kode

'''


# -*- coding: utf-8 -*-
from __future__ import print_function
from builtins import input
import httplib2

from apiclient import discovery
from apiclient import errors
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from email.mime.text import MIMEText
import base64
import email
import xlrd

import datetime
import sys, os

reload(sys)
sys.setdefaultencoding('utf8')

def self_path():
	return os.path.dirname(os.path.realpath(__file__))

## Set up GMAIL API
try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None


# Set global constants
SCOPES = 'https://www.googleapis.com/auth/gmail.send'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Kulturstyret download applications'

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.
	NOTE: Downloaded from google.

    Returns:
        Credentials, the obtained credential.
    """
    credential_dir = os.path.join(self_path(), '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,'gmail-python.json')
    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def create_message(sender, to, subject, message_text):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEText(message_text)
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def send_message(service, user_id, message):
  """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.
  """
  try:
    message = (service.users().messages().send(userId=user_id, body=message)
               .execute())
    print('Message Id: ' + message['id'])
    return message
  except errors.HttpError, error:
    print('An error occurred:')
    print(error)

def create_message_text():
	to = 'johanlor@stud.ntnu.no'
	subject = 'Sit Sponsorprogram - Vedtak'

	start = 'Hei, \nKulturstyret har fattet følgende beslutning vedrørende deres søknad om støtte i Sit sponsorprogram denne perioden:\n'


	slutt = '\nNB: Avgjørelsen kan ikke ankes eller påklages.\n\nMvh. \nJohannes Lorentzen\nLeder av Kulturstyret'

	message_text = start + slutt
	return to, subject, message_text


def import_data_excell(file):
	workbook = xlrd.open_workbook(file)
	workbook = xlrd.open_workbook(file, on_demand = True)
	worksheet = workbook.sheet_by_index(0)
	first_row = [] # The row where we stock the name of the column
	for col in range(worksheet.ncols):
		first_row.append( worksheet.cell_value(0,col) )
	print(first_row)
	#my_dic = pd.read_excel('names.xlsx', index_col=0).to_dict()

## The main part of the program
def main():

	# Set credentials and api
	print('Setting up credentials - This might take some time \n')
	#credentials = get_credentials()
	#http = credentials.authorize(httplib2.Http())
	#service = discovery.build('gmail', 'v1', http=http)
	user_id = 'me'

	to = 'pippids+testing@gmail.com'
	sender = 'pippidis@gmail.com'
	subject = 'test av api'


	file = 'test.xlsx'
	data = import_data_excell(file)

	to, subject, message_text = create_message_text()
	#message = create_message(sender, to, subject, message_text)

	#send_message(service, user_id, message)

if __name__ == '__main__':
    main()
