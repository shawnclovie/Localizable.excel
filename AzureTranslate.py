# -*- coding:utf-8 -*-

import os
import pycurl
import StringIO
import httplib, urllib
import json
from xml.sax.saxutils import escape
try:
	import xml.etree.cElementTree as ET
except ImportError:
	import xml.etree.ElementTree as ET


class CURLRequest:
	def __init__(self, url, method = "GET"):
		self.buffer = StringIO.StringIO()
		c = pycurl.Curl()
		c.setopt(pycurl.URL, url)
		c.setopt(pycurl.CUSTOMREQUEST, method)
		c.setopt(pycurl.WRITEFUNCTION, self.buffer.write)
		self.curl = c

	def set_headers(self, headers):
		self.curl.setopt(pycurl.HTTPHEADER, headers)
		return self

	def set_post_data(self, data):
		if type(data) is dict:
			data = urllib.urlencode(data)
		self.curl.setopt(pycurl.POSTFIELDS, data)
		return self

	def perform(self, format = None):
		self.curl.perform()
		data = self.buffer.getvalue()
		self.buffer.close()
		if format == "json":
			return json.loads(data)
		return data


class Authorizer:
	auth_url = "https://datamarket.accesscontrol.windows.net/v2/OAuth2-13/"

	def __init__(self):
		self.access_token = ""

	def fetch_token(self):
		client_info_file = "azure_client.json"
		if not os.path.exists(client_info_file):
			print "Please save Azure client information into " + client_info_file
			print 'format: {"client_id":"Azure Application ID","client_secret":"Application "}'
			exit()
		params = json.load(client_info_file)
		params["grant_type"] = "client_credentials"
		params["scope"] = "http://api.microsofttranslator.com"
		print "fetching access token..." + self.auth_url
		req = CURLRequest(self.auth_url, "POST")
		req.curl.setopt(pycurl.SSL_VERIFYPEER, False)
		# req.curl.setopt(pycurl.VERBOSE, True)
		data = req.set_post_data(params).perform(format="json")
		self.access_token = data["access_token"]


class Translator:
	translate_url = "http://api.microsofttranslator.com/v2/Http.svc/TranslateArray"

	def __init__(self, authorizer):
		self.authorizer = authorizer

	def request_body(self, texts, from_language, to_language):
		body_header = "<TranslateArrayRequest><AppId /><From>" + from_language + "</From><Texts>"
		body_footer = "</Texts><To>" + to_language + "</To></TranslateArrayRequest>"
		body = []
		for text in texts:
			body.append('<string xmlns="http://schemas.microsoft.com/2003/10/Serialization/Arrays">' + escape(text) + '</string>')
		return body_header + "".join(body) + body_footer

	def translate(self, texts, from_language, to_language):
		if len(self.authorizer.access_token) == 0:
			self.authorizer.fetch_token()
		post_data = self.request_body(texts, from_language, to_language)
		print "translating {0} texts {1} -> {2}...".format(len(texts), from_language, to_language)
		# print post_data
		headers = [
			"Authorization: Bearer " + self.authorizer.access_token,
			"Content-Type: text/xml",
		]
		req = CURLRequest(self.translate_url, "POST")
		req.curl.setopt(pycurl.SSL_VERIFYPEER, False)
		req.set_headers(headers)
		req.set_post_data(post_data)
		data = req.perform()
		root = ET.fromstring(data)
		result = []
		for item in root:
			for attr in item:
				if attr.tag.endswith("TranslatedText"):
					result.append(attr.text)
					break
		return result
