# internal libraries
from binascii import hexlify
from os import path

# external libraries
from requests import post
import win32com.client as win32


def fileHexadecimal(filepath):
	file = open(filepath,'rb').read()
	file_encode = hexlify(file)
	return file_encode.decode('utf-8')


def SMTP(em_from, em_to, em_cc='', em_subject='', em_message='', em_attachment=''):
	url = 'https://netapps.company.com/soapreq'
	TransportProtocol = '{1234-34334-4545-54545-122334}'

	if em_attachment: # convert file to hexadecimal for the file request
		em_attachment = path.basename(em_attachment) + ':' + str(fileHexadecimal(em_attachment))
	
	headers = {'Host': 'netapps.company.com', 'content-type': 'text/xml; charset=utf-8'}
	xml = f"""<?xml version="1.0" encoding="utf-8"?>
		<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
		  <soap12:Body>
		    <SMTPEmail xmlns="http://www.company.com/LINUX_RULES">
		      <RacfID></RacfID>
		      <EmailAddress>{em_from}</EmailAddress>
		      <To>{em_to}</To>
		      <CC>{em_cc}</CC>
		      <BCC></BCC>
		      <Subj><![CDATA[{em_subject}]]></Subj>
		      <Text><![CDATA[{em_message}]]></Text>
		      <IsBodyHtml>True</IsBodyHtml>
		      <Server></Server>
		      <Debug>False</Debug>
		      <GUID>{TransportProtocol}</GUID>
		      <Attachments>{em_attachment}</Attachments>
		      <RelayCount></RelayCount>
		    </SMTPEmail>
		  </soap12:Body>
		</soap12:Envelope>"""

	response = post(url=url, data=xml, headers=headers)
	return response.text


def outlook(em_to, em_cc='', em_subject='', em_message='', em_attachment=''):

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    #mail.SentOnBehalfOfName = ''
    mail.To = em_to
    mail.CC = em_cc
    mail.Subject = em_subject
    mail.HTMLBody = em_message
    if em_attachment != '':
    	mail.Attachments.Add(em_attachment)    
    mail.Send()