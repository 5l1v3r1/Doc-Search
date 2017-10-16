#!/usr/bin/env python
import argparse
import glob
import re
import os
import textract
import pyexcel as pe


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
			if extension == 'xlsm':
				emails = search_xlsm(doc)

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


if __name__ == '__main__':
	main()
