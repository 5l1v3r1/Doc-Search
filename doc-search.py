#!/usr/bin/env python
import argparse
import glob
import re
import os
import textract
import pyexcel as pe
import email.utils
import olefile as OleFile

from email.parser import Parser as EmailParser


EMAIL_REGEX = re.compile(r"(?i)([a-z0-9._-]{1,}@[a-z0-9-]{1,}\.[a-z]{2,})")
FLAT_FORMATS = ['txt', 'out', 'log', 'csv', 'ini']
BAD_FILES = ['exe', 'py', 'pyc', 'pyd', 'dll', 'js' 'css', 'ico']


def main():
	parser = argparse.ArgumentParser(description='Search a directory containing documents for emails addresses')
	parser.add_argument('directory', help='Directory containing documents')
	parser.add_argument('-o', '--outfile', help='File to write found emails address to', default='emails_out.txt')
	args = parser.parse_args()

	found_emails = []
	unprocessed_files = []

	directories = get_files(args.directory)
	for doc in directories:
		try:
			extension = (doc.split('.')[-1]).lower()

			# Skip bad files
			if extension in BAD_FILES:
				continue

			# Process xlsm documents
			elif extension == 'xlsm':
				emails = search_xlsm(doc)

			# Process msg files
			elif extension == 'msg':
				emails = search_msg(doc)

			# Process text documents
			elif extension in FLAT_FORMATS:
				emails = search_text(doc)

			# Process all other documents
			else:
				emails = search_docs(doc)

			# Unique emails
			if len(emails) > 0:
				print("{0} -> {1}".format(emails, doc))
				for email in emails:
					email = email.lower()
					if email in found_emails:
						continue
					else:
						found_emails.append(email)

		except Exception as error:
			print("[-] Unable to process: {0}".format(doc))
			unprocessed_files.append(doc)
			continue

	# Write emails to file
	if len(found_emails) > 0:
		display_emails(found_emails, args.outfile, unprocessed_files)
	else:
		print("[-] No emails found in '{0}'".format(args.directory))


def get_files(directory):
	directories = []

	for root, dirs, filenames in os.walk(directory):
		for filename in filenames:
			directories.append(os.path.join(root, filename))

	return directories


def search_msg(doc):
	emails = []

	outfile = "/tmp/{0}.txt".format(doc.split('/')[-1].split('.')[0])
	msg = Message(doc)
	msg.save(outfile)

	emails = search_text(outfile)

	return emails


def search_text(doc):
	emails = []

	text = open(doc, 'rb')
	for line in text:
		email = EMAIL_REGEX.search(line)
		if email:
			emails.append(email.group(0))

	return emails


def search_xlsm(doc):
	emails = []

	doc_name = doc.split('/')[-1]
	new_doc = "{0}.xls".format(doc_name.split('.')[0])
	sheet = pe.get_book(file_name=doc)
	sheet.save_as("/tmp/{0}".format(new_doc))

	emails = search_docs("/tmp/{0}".format(new_doc))

	return emails


def search_docs(doc):
	emails = []

	text = textract.process(doc)
	emails = EMAIL_REGEX.findall(text)

	return emails


def display_emails(emails, outfile, unprocessed):
	f = open(outfile, 'a')

	for email in emails:
		f.write("{0}\n".format(email))
		print(email)

	f = open(outfile + '.unprocessed', 'a')
	for u in unprocessed:
		f.write("{0}\n".format(u))


class Message(OleFile.OleFileIO):

	def __init__(self, filename):
		OleFile.OleFileIO.__init__(self, filename)

	def _getStream(self, filename):
		if self.exists(filename):
			stream = self.openstream(filename)
			return stream.read()
		else:
			return None

	def _getStringStream(self, filename, prefer='unicode'):
		if isinstance(filename, list):
			filename = "/".join(filename)

		asciiVersion = self._getStream(filename + '001E')
		unicodeVersion = windowsUnicode(self._getStream(filename + '001F'))
		if asciiVersion is None:
			return unicodeVersion
		elif unicodeVersion is None:
			return asciiVersion
		else:
			if prefer == 'unicode':
				return unicodeVersion
			else:
				return asciiVersion

	@property
	def subject(self):
		return self._getStringStream('__substg1.0_0037')

	@property
	def header(self):
		try:
			return self._header
		except Exception:
			headerText = self._getStringStream('__substg1.0_007D')
			if headerText is not None:
				self._header = EmailParser().parsestr(headerText)
			else:
				self._header = None
			return self._header

	@property
	def sender(self):
		try:
			return self._sender
		except Exception:

			if self.header is not None:
				headerResult = self.header["from"]
				if headerResult is not None:
					self._sender = headerResult
					return headerResult

			text = self._getStringStream('__substg1.0_0C1A')
			email = self._getStringStream('__substg1.0_0C1F')
			result = None
			if text is None:
				result = email
			else:
				result = text
				if email is not None:
					result = result + " <" + email + ">"

			self._sender = result
			return result

	@property
	def to(self):
		try:
			return self._to
		except Exception:

			if self.header is not None:
				headerResult = self.header["to"]
				if headerResult is not None:
					self._to = headerResult
					return headerResult

			display = self._getStringStream('__substg1.0_0E04')
			self._to = display
			return display

	@property
	def cc(self):
		try:
			return self._cc
		except Exception:

			if self.header is not None:
				headerResult = self.header["cc"]
				if headerResult is not None:
					self._cc = headerResult
					return headerResult

			display = self._getStringStream('__substg1.0_0E03')
			self._cc = display
			return display

	@property
	def body(self):
		# Get the message body
		return self._getStringStream('__substg1.0_1000')

	def save(self, outfile):

		def xstr(s):
			return '' if s is None else str(s)

		# Save the message body
		f = open("{0}".format(outfile), "w")

		f.write("From: " + xstr(self.sender) + "\n")
		f.write("To: " + xstr(self.to) + "\n")
		f.write("CC: " + xstr(self.cc) + "\n")
		f.write("Subject: " + xstr(self.subject) + "\n")
		f.write("-----------------\n\n")
		f.write((self.body).encode('utf-8'))

		f.close()


def windowsUnicode(string):
	if string is None:
		return None

	return unicode(string, 'utf_16_le')


if __name__ == '__main__':
	main()
