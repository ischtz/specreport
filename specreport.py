#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import sys
import glob
import argparse

def main():

	# Parse command line arguments
	app_desc = 'Create a summary report as HTML table from multiple Speccy XML files'
	parser = argparse.ArgumentParser(description=app_desc)

	parser.add_argument(dest='xmlfiles', action='store', help='Speccy XML file or folder of files')
	parser.add_argument('-o', '--output', dest='outfile', action='store', 
						metavar='<output file>', help='Output file for HTML report')
	opt = parser.parse_args()

	# Create list of input files
	if os.path.isdir(opt.xmlfiles):
		d = os.path.join(os.path.abspath(opt.xmlfiles),'*.xml')
		infiles = glob.glob(d)

	elif os.path.isfile(opt.xmlfiles):
		infiles = [opt.xmlfiles]

	else:
		raise ValueError('Input argument is not a valid file or directory!')

	# Open output file
	if opt.outfile is None:
		opt.outfile = 'specreport.html'
	html = open(opt.outfile, 'w')
	html.write(html_header())

	# Process XML files
	for cur_f in infiles:
		pass

	# Close output file
	html.write(html_footer())
	html.close()


def html_header():
	head = """
	<html>
	<head><title>PC Specifications Report</title></head>
	<body>
	"""
	return head


def html_footer():
	footer = """</body>
	</html>
	"""
	return footer


def table_row():
	pass


if __name__ == '__main__':
	main()
