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
	   'hostsect': 	['Computer Name'],
	   'env': 		['Umgebungsvariablen', 'Environment Variables'],
	   'ip': 		['IP Adresse', 'IP Address'],
	   'ramsect': 	['Memory', 'Speicher'],
	   'ramtype': 	['Type', 'Typ'],
	   'ramsize': 	['Größe', 'Size'],
	   'rambsect':	['Speicherbänke', 'Memory slots'],
	   'rambtot':	['Gesamte Speicherbänke', 'Total memory slots'],
	   'rambused': 	['Genutzte Speicherbänke', 'Used memory slots']
	 }
	
# Output table headers
COLUMNS = ['HostName', 'UserName', 'OS', 'CPU', 'CPUSpeed', 'RAMType', 'RAMSlots', 'RAMSlotsUsed', 'RAMSize', 'HDD', 'OpticalDrive', 'IPAddress', 'ReportDate', 'ReportTime'] 

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

					elif item_id == '6':
						# Storage
						drives = ''
						for drive in item.findall('entry'):
							drives += '{:s}, '.format(drive.attrib['title'].replace(' ATA Device ', ' '))
						fields['HDD'] = drives

					elif item_id == '7':
						# Optical Drives
						odrives = ''
						for drive in item.findall('entry'):
							odrives += '{:s}, '.format(drive.attrib['title'].replace(' ATA Device ', ' '))
						fields['OpticalDrive'] = odrives


			if 'id' in mainsection.attrib:
				# Other sections
				
				if mainsection.attrib['id'] == '1':	
					# OS Section
					os_sect = mainsection.findall('section')
					for item in os_sect:

						if item.attrib['title'] in tr['env']:
							# Environment Variables
							for envvar in item.findall('entry'):
								if envvar.attrib['title'] == 'USERPROFILE':
									# User name from profile folder
									userprofile = envvar.attrib['value']
									fields['UserName'] = os.path.split(userprofile)[1]
				
				if mainsection.attrib['id'] == '3':	
					# RAM Section
					ram_sect = mainsection.findall('section')
					for item in ram_sect:
						if item.attrib['title'] in tr['ramsect']:
							# Memory - general
							for ram in item.findall('entry'):
								if ram.attrib['title'] in tr['ramtype']:
									fields['RAMType'] = ram.attrib['value']
								
								elif ram.attrib['title'] in tr['ramsize']:
									fields['RAMSize'] = ram.attrib['value'].split(' ')[0]

						elif item.attrib['title'] in tr['rambsect']:
							# Memory - slots
							for ram in item.findall('entry'):
								if ram.attrib['title'] in tr['rambtot']:
									fields['RAMSlots'] = int(ram.attrib['value'])
								
								elif ram.attrib['title'] in tr['rambused']:
									fields['RAMSlotsUsed'] = int(ram.attrib['value'])
								

				if mainsection.attrib['id'] == '10':
					# Networking section
					
					net_ent = mainsection.findall('entry')
					for entry in net_ent:
						# Primary IP address
						if entry.attrib['title'] in tr['ip']:
							fields['IPAddress'] = entry.attrib['value']

					net_sect = mainsection.findall('section')
					for item in net_sect:

						if item.attrib['title'] in tr['hostsect']:
							# Computer Name 
							for entry in item.findall('entry'):

								if entry.attrib['title'] in tr['hostname']:
									# Host name (NetBIOS)
									fields['HostName'] = entry.attrib['value']

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
