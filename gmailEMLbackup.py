'''
Script to backup emails via google API to EML format messages.

Message are stored in folders per year.
Mostly re-useing code from Google API examples.
https://developers.google.com/gmail/api/

Author: Riaan Doorduin
Date: 2018
'''

from __future__ import print_function
import httplib2
import os
import sys
import time
import base64
import email
import argparse
from collections import defaultdict
import json

try:
	from apiclient import discovery
	from apiclient import errors
	from oauth2client import client
	from oauth2client import tools
	from oauth2client.file import Storage
	#import google.oauth2.credentials
	#import google_auth_oauthlib.flow
except:
	print ("pip install --upgrade google-api-python-client oauth2client")
	exit()
import humanfriendly

#http://www.five-ten-sg.com/libpst/
#https://stackoverflow.com/questions/1795202/python-code-to-convert-mail-from-pst-to-eml-format

args = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
#SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
CLIENT_SECRET_FILE = 'client_id.json'
#APPLICATION_NAME = 'Gmail API Python Quickstart'
APPLICATION_NAME = 'gmailEMLBackup'
SCOPES = [
	'https://www.googleapis.com/auth/gmail.modify',
	'https://www.googleapis.com/auth/userinfo.email',
	'https://www.googleapis.com/auth/userinfo.profile',
	'https://www.googleapis.com/auth/gmail.readonly'
]

# Based on Google API's
def get_credentials():
	"""Gets valid user credentials from storage.

	If nothing has been stored, or if the stored credentials are invalid,
	the OAuth2 flow is completed to obtain the new credentials.

	Returns:
		Credentials, the obtained credential.
	"""
	home_dir = os.path.abspath('.')
	#home_dir = os.path.expanduser('~')
	credential_dir = os.path.join(home_dir, '.credentials')
	if not os.path.exists(credential_dir):
		os.makedirs(credential_dir)
	credential_path = os.path.join(credential_dir,
								'gmail-python-quickstart.json')
	print(credential_path)

	store = Storage(credential_path)
	credentials = store.get()
	if not credentials or credentials.invalid:	
		flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
		flow.user_agent = APPLICATION_NAME
		if args:
			credentials = tools.run_flow(flow, store, args)
		else: # Needed only for compatibility with Python 2.6
			credentials = tools.run(flow, store)
		print('Storing credentials to ' + credential_path)
	return credentials

"""Get a list of Messages from the user's mailbox.
"""
def ListMessagesMatchingQuery(service, user_id, query=''):
	"""List all Messages of the user's mailbox matching the query.

	Args:
		service: Authorized Gmail API service instance.
		user_id: User's email address. The special value "me"
		can be used to indicate the authenticated user.
		query: String used to filter messages returned.
		Eg.- 'from:user@some_domain.com' for Messages from a particular sender.
		https://support.google.com/mail/answer/7190?hl=en

	Returns:
		List of Messages that match the criteria of the query. Note that the
		returned list contains Message IDs, you must use get with the
		appropriate ID to get the details of a Message.
	"""
	try:
		response = service.users().messages().list(userId=user_id,q=query).execute()
		messages = []
		if 'messages' in response:
			messages.extend(response['messages'])

		while 'nextPageToken' in response:
			page_token = response['nextPageToken']
			response = service.users().messages().list(userId=user_id, q=query, pageToken=page_token).execute()
			messages.extend(response['messages'])

		return messages
	except (errors.HttpError, error):
		print ('An error occurred: %s' % error)

"""Get a list of Messages from the user's mailbox.
		Copied from ListMessagesMatchingQuery
"""
def ListThreadsMatchingQuery(service, user_id, query=''):
	"""List all Threads of the user's mailbox matching the query.

	Args:
		service: Authorized Gmail API service instance.
		user_id: User's email address. The special value "me"
		can be used to indicate the authenticated user.
		query: String used to filter messages returned.
		Eg.- 'from:user@some_domain.com' for Messages from a particular sender.
		https://support.google.com/mail/answer/7190?hl=en

	Returns:
		List of Threads that match the criteria of the query. Note that the
		returned list contains Thread IDs, you must use get with the
		appropriate ID to get the details of a Thread.
	"""
	try:
		response = service.users().threads().list(userId=user_id,
												   q=query).execute()
		threads = []
		if 'threads' in response:
			threads.extend(response['threads'])

		while 'nextPageToken' in response:
			page_token = response['nextPageToken']
			response = service.users().threads().list(userId=user_id, q=query,
											 pageToken=page_token).execute()
			threads.extend(response['threads'])

		return threads
	except (errors.HttpError, error):
		print ('An error occurred: %s' % error)

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
			essages.extend(response['messages'])

			return messages
	except (errors.HttpError, error):
		print ('An error occurred: %s' % error)

def GetMessage(service, user_id, msg_id,msg_format='minimal'):
	"""Get a Message with given ID.

	Args:
		service: Authorized Gmail API service instance.
		user_id: User's email address. The special value "me"
		can be used to indicate the authenticated user.
		msg_id: The ID of the Message required.

	Returns:
		A Message.
	"""
	try:
		message = service.users().messages().get(userId=user_id, id=msg_id,format=msg_format).execute()

		#print ('\tMessage snippet: %s' % message['snippet'])
		return message

	except (errors.HttpError, error):
		print ('An error occurred: %s' % error)

def GetMimeMessage(service, user_id, msg_id):
	"""Get a Message and use it to create a MIME Message.

	Args:
		service: Authorized Gmail API service instance.
		user_id: User's email address. The special value "me"
		can be used to indicate the authenticated user.
		msg_id: The ID of the Message required.

  Returns:
		A MIME Message, consisting of data from Message.
	"""
	try:
		message = service.users().messages().get(userId=user_id, id=msg_id,
												format='raw').execute()

		#print ('\tMessage snippet: %s' % message['snippet'])

		msg_str = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
		#print (msg_str)
		return msg_str
		#mime_msg = email.message_from_string(msg_str)
		#return mime_msg
	except (errors.HttpError, error):
		print ('An error occurred: %s' % error)

def DeleteMessage(service, user_id, msg_id):
	"""Delete a Message.

	Args:
		service: Authorized Gmail API service instance.
		user_id: User's email address. The special value "me"
		can be used to indicate the authenticated user.
		msg_id: ID of Message to delete.
	"""
	try:
		#print ('...')
		res = service.users().messages().delete(userId=user_id, id=msg_id).execute()
		print ('\tMessage with id: %s deleted successfully.' % msg_id)
		return res
	except (errors.HttpError, error):
		print ('An error occurred: %s' % error)

def TrashMessage(service, user_id, msg_id):
	"""Delete a Message.

	Args:
		service: Authorized Gmail API service instance.
		user_id: User's email address. The special value "me"
		can be used to indicate the authenticated user.
		msg_id: ID of Message to delete.
	"""
	try:
		#print ('...')
		res = service.users().messages().trash(userId=user_id, id=msg_id).execute()
		#print ('...')
		print ('\tMessage with id: %s trashed successfully.' % msg_id)
		return res
	except (errors.HttpError, error):
		print ('An error occurred: %s' % error)

#Own Work:
def ListMessagesFromThreads(service, user_id, threads):
	"""List all Messages of the user's mailbox expanded from the threads applied.

	Args:
		service: Authorized Gmail API service instance.
		user_id: User's email address. The special value "me"
		can be used to indicate the authenticated user.
		label_ids: List of Threads

	Returns:
		List of Messages that have all required Threads applied. Note that the
		returned list contains Message IDs, you must use get with the
		appropriate id to get the details of a Message.
	"""
	try:
		messages = []
		i = 0;
		print ("#Threads:",len(threads))
		for thread in threads:
			print("i=",i)
			i = i+1
			printDictionary(thread,"Thread:")
			response = service.users().threads().get(userId=user_id,
														id=thread['id'],
														format='minimal').execute()
			printDictionary(response,'GetThread:')
			threadMessages = response['messages']
			print ("\tSize:",len(threadMessages))
			if len(threadMessages) > 1:
				for message in threadMessages:
					printDictionary(message,"\t\tMessage:")
					print("\tCTime:",time.ctime(int(message['internalDate'])*0.001))
					return
			if 'messages' in response:
				messages.extend(threadMessages)
			while 'nextPageToken' in response:
				page_token = response['nextPageToken']
				response = service.users().threads().get(userId=user_id,
														id=thread['id'],
														pageToken=page_token).execute()
				messages.extend(response['messages'])
			if i > 10:
				return

		return messages
	except (errors.HttpError, error):
		print ('An error occurred: %s' % error)
	

def uprint(*objects, sep=' ', end='\n', file=sys.stdout):
	enc = file.encoding
	if enc == 'UTF-8':
		print(*objects, sep=sep, end=end, file=file)
	else:
		f = lambda obj: str(obj).encode(enc, errors='backslashreplace').decode(enc)
		print(*map(f, objects), sep=sep, end=end, file=file)

def printDictionary(dictionary,heading=None):
	if heading is not None:
		print (heading)
	for key in dictionary.keys():
		try:
			print ("\t",key, ":", dictionary[key])
		except:
			uprint ("\t\t",key, ":", dictionary[key])
	
def doMain(q,older_than=False,byYear=False,sizePerYear=False,purging=False,theYear=None):
	"""Shows basic usage of the Gmail API.
	Creates a Gmail API service object and outputs a list of label names
	of the user's Gmail account.
	"""

	#SetupBackupPath per Year
	backup_dir = os.path.abspath('.')
	#Setup indexing options
	if sizePerYear:
		sizeBinPerYear = defaultdict(list)

	#Authenticate
	credentials = get_credentials()
	http = credentials.authorize(httplib2.Http())
	service = discovery.build('gmail', 'v1', http=http)
	profile = service.users().getProfile(userId='me').execute()
	printDictionary(profile,"Profile:")

	#threads = ListThreadsMatchingQuery(service,'me', q)  
	#messages = ListMessagesFromThreads(service,'me',threads)
	print("Query to run: ",q)
	messages = ListMessagesMatchingQuery(service,'me', q)

	lenMessages = len(messages)
	print ('Len:',lenMessages)
	i = 0
	yearToday = time.gmtime().tm_year
	print ("yearToday=",yearToday)
	for message in messages:
		i = i + 1
		if args.verbose:
			printDictionary(message,"Message: %d / %d"%(i,lenMessages))
		theMessageID = message['id']
		if args.query:			
			message = GetMessage(service,'me',theMessageID,'metadata')
		else:
			if args.subject:
				message = GetMessage(service,'me',theMessageID,'metadata')
			else:
				message = GetMessage(service,'me',theMessageID)
		if args.verbose:
			printDictionary(message)
		if not args.query:			
			if 'labelIds' in message:
				label_ids = message['labelIds']
				if args.verbose:
					print ('\t',label_ids)
				if 'TRASH' in label_ids:
					print ("\tAlready Trashed...")
					continue
			else:
				if args.verbose:
					print ("\tNo LabelOIDs")
		if ((args.query is not None) or (args.subject)):
					subject = ""
					for pair in message['payload']['headers']:
						if pair['name'] in "Subject":
							subject = pair['value']
							print (pair['name'],pair['value'])
					tmpStr = ""
					for c in subject:
						if c in '-_.()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789':
							tmpStr+= c
						else:
							tmpStr+= '_'
					subject	= tmpStr
					print (subject)
		if args.query:
			if 'internalDate' in message:
				if args.verbose:
					print("\tCTime:",time.ctime(int(message['internalDate'])*0.001))
				theTimeStamp = time.gmtime(int(message['internalDate'])*0.001)
				theTimeStampStr = "%04d%02d%02d"%(theTimeStamp.tm_year,theTimeStamp.tm_mon,theTimeStamp.tm_mday)
				if args.verbose:
					print("\tgmTime:",theTimeStamp)
				#Copy and Purge
				if True:
					#SetupDir and File Path
					if args.subject:
						fileName = theTimeStampStr+"_"+ theMessageID+"_"+subject+".eml"
					else:
						fileName = theTimeStampStr+"_"+ theMessageID+".eml"
					filebackup_dir = os.path.join(backup_dir, str(theTimeStamp.tm_year))
					if not os.path.exists(filebackup_dir):
						os.makedirs(filebackup_dir)
					fileName_path = os.path.join(filebackup_dir, fileName)
					print("\t",fileName_path)
					#Get Full Message
					rawMsg = GetMimeMessage(service,'me',theMessageID)
					print ("\t",fileName)
					try:
						f = open(fileName_path,"wb")
						f.write(rawMsg)
						f.close()
						if purging:
							print ("\tDeleting Message", fileName)
							TrashMessage(service,'me',theMessageID)
					except:
						print ("Failure saving EML: %s ! "%fileName)
		else:
			if 'internalDate' in message:
				if args.verbose:
					print("\tCTime:",time.ctime(int(message['internalDate'])*0.001))
					
				theTimeStamp = time.gmtime(int(message['internalDate'])*0.001)
				theTimeStampStr = "%04d%02d%02d"%(theTimeStamp.tm_year,theTimeStamp.tm_mon,theTimeStamp.tm_mday)
				if args.verbose:
					print("\tgmTime:",theTimeStamp)
				if sizePerYear:
					msgSize = int(message['sizeEstimate'])
					if args.verbose:
						print("\t",theTimeStamp.tm_year,"\t:\t",humanfriendly.format_size(msgSize))
					else:
						print("Message: %d / %d\t"%(i,lenMessages),theTimeStamp.tm_year,"\t:\t",humanfriendly.format_size(msgSize))
					try:
						tmpSize = sizeBinPerYear[theTimeStamp.tm_year]
						tmpSize = tmpSize + msgSize
						sizeBinPerYear[theTimeStamp.tm_year] = tmpSize
					except:
						sizeBinPerYear[theTimeStamp.tm_year] = msgSize
					if args.verbose:
						#print (sizeBinPerYear)
						print ("---")
				#Double Check
				print ("Message Year:",theTimeStamp.tm_year)
				pleaseContinue = False
				if byYear and theYear == theTimeStamp.tm_year:
					pleaseContinue = True
				if older_than and (theTimeStamp.tm_year <= (yearToday-theYear)):
					pleaseContinue = True
				if not pleaseContinue:
					continue
				#Copy and Purge
				if pleaseContinue:
					#SetupDir and File Path
					fileName = theTimeStampStr+"_"+theMessageID+"_"+subject+".eml"
					filebackup_dir = os.path.join(backup_dir, str(theTimeStamp.tm_year))
					if not os.path.exists(filebackup_dir):
						os.makedirs(filebackup_dir)
					fileName_path = os.path.join(filebackup_dir, fileName)
					fileName_json = os.path.join(filebackup_dir, "index.json")
					print("\t",fileName_path)
					#Get Full Message
					rawMsg = GetMimeMessage(service,'me',theMessageID)
					print ("\t",fileName)
					if args.subject:
						try:
							f = open(fileName_json,'a')
							json.dump(message,f,indent=1)
							f.write("\n")
							f.close()
						except:
							f.close()
							print ("Failure saving JSON: %s ! "%fileName_json)
					try:
						f = open(fileName_path,"wb")
						f.write(rawMsg)
						f.close()
						if purging:
							print ("\tDeleting Message", fileName)
							TrashMessage(service,'me',theMessageID)
					except:
						print ("Failure saving EML: %s ! "%fileName)
#		if i > 10:
#			return
	if sizePerYear:
		printDictionary(sizeBinPerYear)
		for theSize in sizeBinPerYear:
			print(theSize,"\t:\t",humanfriendly.format_size(sizeBinPerYear[theSize]))
	return

	print (" ")
	results = service.users().labels().list(userId='me').execute()
	labels = results.get('labels', [])

	if not labels:
		print('No labels found.')
	else:
		print('Labels:')
		for label in labels:
			print(label['name'])


def main():
	global args
	parser = argparse.ArgumentParser(description='Backs up emails to .EML format message per year. Default is to backup only. However purging can be activated as well.',parents=[tools.argparser])
	parser.add_argument('--older_than',  dest='older_than', default= None, type=int, help='Sets the search query to the search for message older than specified year')
	#parser.add_argument('--query', dest='query', action='store_true', default=False, help='Specify query to backup')
	parser.add_argument('--query', dest='query', default=None, help='Specify query to backup')
	parser.add_argument('--year', dest='year', type=int, default=None, help='Select a specific year to backup')
	parser.add_argument('--sizePerYear', dest='sizePerYear', action='store_true', default=False, help='Calculates the size per year (from 2000)')
	parser.add_argument('--subject', dest='subject', action='store_true', default=False, help='Include Subject in email filename (only for when query is active)')
	parser.add_argument('--purging', dest='purging', action='store_true', default=False, help='Removes messages from google servers into trash.')
	parser.add_argument('--verbose', dest='verbose', action='store_true', default=False, help='Increases verbosity')	
	args = parser.parse_args()
	print (args)
	#Query for messages:
	# https://support.google.com/mail/answer/7190?hl=en
	if args.older_than is not None:
		q = 'older_than:%dy'%args.older_than
		doMain(q,older_than=True,sizePerYear=args.sizePerYear,purging=args.purging,theYear=args.year)
	if args.year is not None:
		q = 'after:%d/01/01 before:%d/01/01'%(args.year,args.year+1)
		doMain(q,byYear=True,sizePerYear=args.sizePerYear,purging=args.purging,theYear=args.year)
	if args.query is not None:
		#q = 'has:attachment after:2018/11/4 before:2018/11/7'
		q = args.query
		print (args.query)
		doMain(q,byYear=False,sizePerYear=None,purging=args.purging,theYear=None)

if __name__ == '__main__':
	main()