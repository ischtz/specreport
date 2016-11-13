#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import sys
import time
import glob
import argparse

import xml.etree.ElementTree as ET

# XML item names are internationalized, this holds possible translations
tr = { 'summary': 	['Summary', 'Zusammenfassung'],
	   'hostname':	['NetBIOS Name'],
	 }
	
# Output table headers
out_head = ['Host Name', 'User', 'OS', 'CPU', 'CPU Speed', 'RAM', 'HDD', 'Date of report'] 

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

	# Open output file, write table header
	if opt.outfile is None:
		opt.outfile = 'specreport.html'
	html = open(opt.outfile, 'w')
	html.write(html_header())
	html.write(table_row(out_head, header=True))

	# Process XML files
	for cur_f in infiles:
		fields = ['', '', '', '', '', '', '', '']
		spec = ET.parse(cur_f)
		specr = spec.getroot() 
		
		for mainsection in specr.findall('mainsection'):

			if mainsection.attrib['title'] in tr['summary']:
				# Summary: Get OS, CPU, RAM, HDD
				summary = mainsection.findall('section')
				for item in summary:
					item_id = item.attrib['id']
					first_entry = item.find('entry')
					
					if item_id == '1':
						# OS
						fields[2] = first_entry.attrib['title'].strip()

					elif item_id == '2':
						# CPU + Speed
						cpu = first_entry.attrib['title'].split('@')
						fields[3] = cpu[0].strip()
						fields[4] = cpu[1].strip()

					elif item_id == '3':
						# RAM
						ram = first_entry.attrib['title'].split('@')
						fields[5] = ram[0].split()[0].strip()

					elif item_id == '6':
						# Storage
						drives = ''
						for drive in item.findall('entry'):
							drives += '{:s}</br>'.format(drive.attrib['title'].replace(' ATA Device ', ' '))
						fields[6] = drives

		html.write(table_row(fields))

	# Close output file
	html.write(html_footer())
	html.close()


def html_header():
	""" Header for HTML output including table """
	head = """
	<html>
	<head><title>PC Specifications Summary</title></head>
	<body>
	PC Specifications Summary, generated on {:s}</br></br>
	<table border>
	"""
	t = time.strftime('%d.%m.%y, %H:%M:%S', time.localtime())
	return head.format(t)


def html_footer():
	""" Footer for HTML table / file """
	footer = """</table></body>
	</html>
	"""
	return footer


def table_row(fields, header=False):
	""" Writes a table row """
	row = '<tr>'
	if not header:
		row += '<td>{:s}</td>' * len(fields)
	else:
		row += '<th>{:s}</th>' * len(fields)
	row += '</tr>'
	return row.format(*fields)


if __name__ == '__main__':
	main()
