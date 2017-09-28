# -*- coding:utf-8 -*-

import os
import json
import requests
from xml.sax.saxutils import escape
try:
	import xml.etree.cElementTree as ET
except ImportError:
	import xml.etree.ElementTree as ET

class Authorizer:
	auth_url = "https://api.cognitive.microsoft.com/sts/v1.0/issueToken"

	def __init__(self):
		self.access_token = ""

	def fetch_token(self):
		client_info_file = "azure_client.json"
		if not os.path.exists(client_info_file):
			print "Please save Azure client information into " + client_info_file
			print 'format: {"client_id":"Azure Application ID","client_secret":"Application "}'
			exit()
		fp = open(client_info_file, 'r')
		secret = json.load(fp)["client_secret"]
		fp.close()
		print "fetching access token..." + self.auth_url
		key = "Ocp-Apim-Subscription-Key"
		response = requests.post(self.auth_url, headers={key: secret})
		response.raise_for_status()
		self.access_token = response.content
		print "access_token=" + self.access_token


class Translator:
	translate_url = "http://api.microsofttranslator.com/v2/Http.svc/TranslateArray"

	def __init__(self, authorizer):
		self.authorizer = authorizer

	def request_body(self, texts, from_language, to_language):
		body_header = u"<TranslateArrayRequest><AppId /><From>" + from_language + u"</From><Texts>"
		body_footer = u"</Texts><To>" + to_language + u"</To></TranslateArrayRequest>"
		body = []
		for text in texts:
			body.append(u'<string xmlns="http://schemas.microsoft.com/2003/10/Serialization/Arrays">' + escape(text) + u'</string>')
		return body_header + u"".join(body) + body_footer

	def translate(self, texts, from_language, to_language):
		if len(self.authorizer.access_token) == 0:
			self.authorizer.fetch_token()
		post_data = self.request_body(texts, from_language, to_language)
		print "translating {0} texts {1} -> {2}...".format(len(texts), from_language, to_language)
		# print post_data
		headers = {
			"Authorization": "Bearer " + self.authorizer.access_token,
			"Content-Type": "text/xml",
		}
		response = requests.post(self.translate_url, headers=headers, data=post_data.encode("utf-8"))
		response.raise_for_status()
		data = response.content
		# print data
		root = ET.fromstring(data)
		result = []
		for item in root:
			for attr in item:
				if attr.tag.endswith("TranslatedText"):
					result.append(attr.text)
					break
		return result
