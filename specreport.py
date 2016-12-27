#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import sys
import time
import glob
import argparse

from pandas import DataFrame

import xml.etree.ElementTree as ET

# XML item names are internationalized, this holds possible translations
tr = { 'summary': 	['Summary', 'Zusammenfassung'],
	   'hostname':	['NetBIOS Name'],
	 }
	
# Output table headers
COLUMNS = ['HostName', 'UserName', 'OS', 'CPU', 'CPUSpeed', 'RAM', 'HDD', 'ReportDate', 'ReportTime'] 

def scan_xml_files(xml_files=[]):
	""" Scan a folder of Speccy XML files into a DataFrame. 

	Args:
	xml_files: 		List of XML file names to include in DataFrame
	"""

	df = DataFrame(columns=COLUMNS)

	for xidx, xfile in enumerate(xml_files):
		spec = ET.parse(xfile)
		specr = spec.getroot()
		fields = {}
		
		# Date and time of report
		rt = time.strptime(specr.attrib['localtime'], '%Y%m%dT%H%M%S%z')
		fields['ReportDate'] = time.strftime('%d.%m.%Y', rt)
		fields['ReportTime'] = time.strftime('%H:%M:%S', rt)

		for mainsection in specr.findall('mainsection'):

			if mainsection.attrib['title'] in tr['summary']:
				# Summary: Get OS, CPU, RAM, HDD
				summary = mainsection.findall('section')
				for item in summary:
					item_id = item.attrib['id']
					first_entry = item.find('entry')
					
					if item_id == '1':
						# OS
						fields['OS'] = first_entry.attrib['title'].strip()

					elif item_id == '2':
						# CPU + Speed
						cpu = first_entry.attrib['title'].split('@')
						fields['CPU'] = cpu[0].strip()
						fields['CPUSpeed'] = cpu[1].strip()

					elif item_id == '3':
						# RAM
						ram = first_entry.attrib['title'].split('@')
						fields['RAM'] = ram[0].split()[0].strip()

					elif item_id == '6':
						# Storage
						drives = ''
						for drive in item.findall('entry'):
							drives += '{:s}, '.format(drive.attrib['title'].replace(' ATA Device ', ' '))
						fields['HDD'] = drives

		df = df.append(fields, ignore_index=True)

	return df


def _main():

	# Parse command line arguments
	app_desc = 'Create a summary report from multiple Speccy XML files'
	parser = argparse.ArgumentParser(description=app_desc)

	parser.add_argument(dest='xmlfiles', action='store', help='Speccy XML file or folder of files')
	parser.add_argument(dest='outfile', action='store', help='Output file name (default: report.csv)')
	parser.add_argument('-x', '--excel', dest='xlsx', action='store_true', 
						help='Output summary as XLSX file.')
	parser.add_argument('-c', '--csv', dest='csv', action='store_true', 
						help='Output summary as CSV file.')
	parser.add_argument('-t', '--html', dest='html', action='store_true', 
						help='Output summary as HTML file.')
	parser.add_argument('-j', '--json', dest='json', action='store_true', 
						help='Output summary as JSON file.')
	opt = parser.parse_args()

	# Create list of input files
	if os.path.isdir(opt.xmlfiles):
		d = os.path.join(os.path.abspath(opt.xmlfiles),'*.xml')
		infiles = glob.glob(d)

	elif os.path.isfile(opt.xmlfiles):
		infiles = [opt.xmlfiles]

	else:
		raise ValueError('Input argument is not a valid file or directory!')

	# Output CSV report by default
	if not opt.xlsx and not opt.csv:
		opt.csv = True

	# Process files
	report = scan_xml_files(infiles)

	# Save reports
	if opt.csv:
		report.to_csv(opt.outfile + '.csv', index=False)
	if opt.xlsx:
		report.to_excel(opt.outfile + '.xlsx', index=False)
	if opt.html:
		report.to_html(opt.outfile + '.html', index=False)
	if opt.json:
		report.to_json(opt.outfile + '.json')


if __name__ == '__main__':
	_main()
