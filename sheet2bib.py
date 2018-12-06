#!/usr/bin/env python3
# By Brian Ballsun-Stanton
# GPL v3 

import google
from collections import defaultdict
from pprint import pprint
from docx import Document
from docx.shared import Cm

sheet = google.getBibliographicItems()
document = Document()


output = defaultdict(list)

for item in sheet:
	pprint(item)
	
	output[item['CategoryName']].append(item)

pprint(output)

for category in output:
	document.add_heading(category, 1)
	for item in output[category]:
		p = document.add_paragraph()
		p.add_run("{}{}".format(item['Book_Title'], item['Volume_Count'])).bold = True
		p.add_run(" {}".format(item['Accession_Number']))

		p2 = document.add_paragraph("〔{}〕{}".format(item['Author_1_Period'], item['Authorship_String_1']))
		paragraph_format = p2.paragraph_format
		paragraph_format.left_indent = Cm(0.5)

		if item['Book_title_relative_to_author_1']:
			p2.add_run("《{}》".format(item['Book_title_relative_to_author_1']))

		if item['Authorship_String_2']:
			p2.add_run("〔{}〕{}".format(item['Author_2_Period'], item['Authorship_String_2']))

		p3 = document.add_paragraph("{}".format(item['Date_of_Publication']))
		paragraph_format = p3.paragraph_format
		paragraph_format.left_indent = Cm(0.5)

		if 'Western_Year' in item and item['Western_Year']:
			p3.add_run("（{}）".format(item['Western_Year']))

		if 'Name_of_Publisher' in item and item['Name_of_Publisher']:
			p3.add_run(" {}".format(item['Name_of_Publisher']))

		if 'Note' in item:
			p4 = document.add_paragraph("{}".format(item['Note']))
			paragraph_format = p4.paragraph_format
			paragraph_format.left_indent = Cm(0.5)




"""

Repeating:
	[Bold:]Book title Volume Count " " [Not bold:]Accession Number
	(Period) Commentator's Name Etal Ownership verb 《Original Book Title》(Period) Commentator's Name Etal Ownership verb
	Date of Publication (Western Year) Name of Publisher Edition
	Note
	"""

document.save('test.docx')