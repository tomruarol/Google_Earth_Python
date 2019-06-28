'''
This script drinks information provided from an Excel output and generates a .kmz file
'''

import geocoding_for_kml
import csv
import xml.dom.minidom
import sys


def extractAddress(row):

	# This extracts an address from a row and returns it as a string. This requires knowing
	# ahead of time what the columns are that hold the address information.

	return '%s,%s,%s,%s,%s' % (row['Address1'], row['Address2'], row['City'], row['State'], row['Zip'])


def createPlacemark(kmlDoc, row, order):

	# This creates a  element for a row of data.
	# A row is a dict.
	placemarkElement = kmlDoc.createElement('Placemark')
	extElement = kmlDoc.createElement('ExtendedData')
	placemarkElement.appendChild(extElement)

	# Loop through the columns and create a  element for every field that has a value.
	for key in order:
		if row[key]:
			dataElement = kmlDoc.createElement('Data')
			dataElement.setAttribute('name', key)
			valueElement = kmlDoc.createElement('value')
			dataElement.appendChild(valueElement)
			valueText = kmlDoc.createTextNode(row[key])
			valueElement.appendChild(valueText)
			extElement.appendChild(dataElement)

	pointElement = kmlDoc.createElement('Point')
	placemarkElement.appendChild(pointElement)

	'''
	This code from here is not needed as we already have the longitude and latitude from our desired points.
	'''
	# coordinates = geocoding_for_kml.geocode(extractAddress(row))
	# coorElement = kmlDoc.createElement('coordinates')
	# coorElement.appendChild(kmlDoc.createTextNode(coordinates))
	# pointElement.appendChild(coorElement)

	'''
	Though, we need to add our long & latitude manually
	'''
	# TODO
	coordinates = str() + ',' + str() # 14.25,15.23 --> lat,long
	coorElement = kmlDoc.createElement('coordinates')
	coorElement.appendChild(kmlDoc.createTextNode(coordinates))
	pointElement.appendChild(coorElement)

	return placemarkElement


def createKML(csvReader, fileName, order):

	# This constructs the KML document from the CSV file.
	kmlDoc = xml.dom.minidom.Document()

	kmlElement = kmlDoc.createElementNS('http://earth.google.com/kml/2.2', 'kml')
	kmlElement.setAttribute('xmlns', 'http://earth.google.com/kml/2.2')
	kmlElement = kmlDoc.appendChild(kmlElement)
	documentElement = kmlDoc.createElement('Document')
	documentElement = kmlElement.appendChild(documentElement)

	# Skip the header line.
	csvReader.next()

	for row in csvReader:
		placemarkElement = createPlacemark(kmlDoc, row, order)
		documentElement.appendChild(placemarkElement)
	kmlFile = open(fileName, 'w')
	kmlFile.write(kmlDoc.toprettyxml('  ', newl='\n', encoding='utf-8'))


def main():
	# This reader opens up 'google-addresses.csv', which should be replaced with your own.
	# It creates a KML file called 'google.kml'.

	# If an argument was passed to the script, it splits the argument on a comma
	# and uses the resulting list to specify an order for when columns get added.
	# Otherwise, it defaults to the order used in the sample.

	if len(sys.argv) > 1:
		order = sys.argv[1].split(',')
	else:
		order = ['Office', 'Address1', 'Address2', 'Address3', 'City', 'State', 'Zip', 'Phone', 'Fax']
	csvreader = csv.DictReader(open('google-addresses.csv'), order) # TODO
	kml = createKML(csvreader, 'google-addresses.kml', order) # TODO


if __name__ == '__main__':
	main()