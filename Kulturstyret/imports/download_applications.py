'''
Denne filen inneholder klassen og koden for å laste ned søknader fra gmail.
det som er her nå er bare kopiert fra tidligere program
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
import base64
import email

import datetime
import sys, os

reload(sys)
sys.setdefaultencoding('utf8')



## Set up GMAIL API
try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# Set global constants
SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
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
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'gmail-python-quickstart.json')

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




## Searching for messages
def ListMessagesMatchingQuery(service, user_id, query=''):
  """List all Messages of the user's mailbox matching the query.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    query: String used to filter messages returned.
    Eg.- 'from:user@some_domain.com' for Messages from a particular sender.

  Returns:
    List of Messages that match the criteria of the query. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate ID to get the details of a Message.
  """
  try:
    response = service.users().messages().list(userId=user_id,
                                               q=query).execute()
    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=user_id, q=query,
                                         pageToken=page_token).execute()
      messages.extend(response['messages'])

    return messages
  except errors.HttpError, error:
    print('An error occurred: ' + error)


def ListMessagesWithLabels(service, user_id, label_ids=[]):
  """List all Messages of the user's mailbox with label_ids applied.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    label_ids: Only return Messages with these labelIds applied.

  Returns:
    List of Messages that have all required Labels applied. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate id to get the details of a Message.
  """
  try:
    response = service.users().messages().list(userId=user_id,
                                               labelIds=label_ids).execute()
    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=user_id,
                                                 labelIds=label_ids,
                                                 pageToken=page_token).execute()
      messages.extend(response['messages'])

    return messages
  except errors.HttpError, error:
    print('An error occurred: ' + error)


def GetMessage(service, user_id, msg_id):
  """Get a Message with given ID.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    msg_id: The ID of the Message required.

  Returns:
    A Message - only metadata
  """
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id, format='metadata').execute()
    return message
  except errors.HttpError, error:
	print('An error occurred: ' + error)


def GetAttachments(service, user_id, msg_id, store_dir):
	try:
		message = service.users().messages().get(userId=user_id, id=msg_id).execute()

		for part in message['payload']['parts']:
			if part['filename']:
				if 'data' in part['body']:
					data=part['body']['data']
				else:
					att_id=part['body']['attachmentId']
					att=service.users().messages().attachments().get(userId=user_id, messageId=msg_id,id=att_id).execute()
					data=att['data']
				#print(data)
				file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))


				path = store_dir + '\\' + part['filename']
				try:
					print('\t' + part['filename'])
				except:
					print('\t -Something wrong with the file')
#				file_data.replace(' ','+')
#				file_data.replace('-','+')
#				file_data.replace('_','/')
				with open(path, 'wb') as f:
				# Note the 'wb', as this ensures that it write binary and not text
					f.write(file_data)
	except errors.HttpError, error:
		print('An error occurred: ' + error)


def self_path():
    path = os.path.dirname(sys.argv[0])
    if not path:
        path = '.'
    return path

def getPeriod():
	p_start = input("Start date (YYYY/MM/DD): ")
	if len(p_start) < 3:
		if int(datetime.date.today().month) < 6:
			p_start = str(datetime.date.today().year-1) + '/11/15'
		else:
			p_start = datetime.date.today().year + '/06/01'
		print('\tStart date is set to: ' + p_start)

	p_slutt = input("\n\nEnd date (YYYY/MM/DD): ")
	if len(p_slutt) < 3:
		if int(datetime.date.today().month) < 6:
			p_slutt = str(datetime.date.today().year) + '/3/15'
		else:
			p_slutt = str(datetime.date.today().year) + '/10/15'
		print('\tEnd date is set to: ' + p_slutt +'\n\n')

	return p_start, p_slutt

def setOutputFolder(org):
	soknad_folder = ''
	if int(datetime.date.today().month) < 6:
		soknad_folder = 'Søknader våren + ' + str(datetime.date.today().year)
	else:
		soknad_folder = 'Søknader høsten ' + str(datetime.date.today().year)

	store_dir = self_path() + '\\' + soknad_folder + '\\' + org
	if not os.path.exists(store_dir):
		os.makedirs(store_dir)
	return store_dir


## The main part of the program
def main():

	# Set credentials and api
	print('Setting up credentials - This might take some time \n')
	credentials = get_credentials()
	http = credentials.authorize(httplib2.Http())
	service = discovery.build('gmail', 'v1', http=http)
	user_id = 'me'

	# Get the period to download emails
	p_start, p_slutt = getPeriod()

	# Finding the emails
	query = 'From:no-reply@sit.no AND has:attachment AND after:' + p_start + ' AND before:' + p_slutt + '  AND subject:Sponsorsøknad AND In:anywhere'
	messages = ListMessagesMatchingQuery(service, user_id, query)

	# Saving attachments in folder

	for message in messages:
		# Get the name of the organisation
		m = GetMessage(service, user_id, message['id']) #Get metadata
		for header in m['payload']['headers']:
			if header['name'] == 'Subject':
				subject = (header['value'])
		org = subject.replace('Sponsorsøknad: ','')
		print(org)

		# Create the storage direcory
		store_dir = setOutputFolder(org)

		# Get the attachments
		GetAttachments(service, user_id, message['id'], store_dir)

if __name__ == '__main__':
    main()
