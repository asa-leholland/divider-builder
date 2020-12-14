# builder.py


# import the os module to obtain the folder path of files
import os

# use regex to locate strings in the soup file
import re

# import to reference template divider file
import resources



def get_list_of_divider_names(path):
	"""
	Returns a list of the names of dividers needed for a provided filepath
	"""

	# Open the file
	

	# Locate the set of paragraphs of the appendices representing the names of the dividers we need to add
	

	# Set up a list variable to hold the strings of the names of the dividers
	divider_names = []

	# Iterate over the paragraphs in the provided file, adding each divider name to our list
	
	
	# return the list
	return divider_names
	


def make_dividers(filepath_to_word_document):
	"""
	Creates Dividers for a provided word document
	"""

	# Obtain the folder containing the provided filename (this is where we will build the divider files)
	destination_folder = os.path.dirname(path=filepath_to_word_document)

	# Obtain the list of divider names we need to create
	divider_appendices = get_list_of_divider_names(path=filepath_to_word_document)

	# Create variable for the path of the divider template doc
	divider_template = "resources\\Document Divider Template.doc"

	# Start word
	word_app = comtypes.client.CreateObject('Word.Application')

	# iterate over the names
	for divider_appendix_string in divider_appendices:

		# Open word using the template
		divider_doc = word_app.Documents.Open(divider_template)

		# Create a divider for each name provided
		make_divider(path=closeout_folder, divider_string=divider_appendix_string) 
