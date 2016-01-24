# coding: utf-8

import appex
import clipboard
import zipfile
import xml.etree.ElementTree

def main():
	# Step one, open the zipfile
	word_file = appex.get_file_path()
	if not word_file:
		return
	if not zipfile.is_zipfile(word_file):
		return
	
	with zipfile.ZipFile(word_file, 'r') as pkg:
		w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
		# Step two, read and save all character styles and paragraph styles
		styles = {}
		style_tags = xml.etree.ElementTree.parse(pkg.open("word/styles.xml")).getroot()
		search_str = './/{{{ns}}}style'.format(ns=w)
		print search_str
		tally = 0
		for style in style_tags.findall(search_str):
			tally += 1
			# TODO: Get styles and save them in a class maybe?
			# <w:style w:type="paragraph" w:styleId="Heading1">
			if style.get('{{{ns}}}type'.format(ns=w)) == 'paragraph' or style.get('{{{ns}}}type'.format(ns=w)) == 'character':
				style_id = style.get('styleId')

				bold_tag = style.find('.//{{{ns}}}rPr/{{{ns}}}b'.format(ns=w))
				if bold_tag is not None:
					bold = bold_tag.get('{{{ns}}}val'.format(ns=w)) != '0' and bold_tag.get('{{{ns}}}val'.format(ns=w)) != 'false' and bold_tag.get('{{{ns}}}val'.format(ns=w)) != 'f'
				else:
					bold = None

				italic_tag = style.find('.//{{{ns}}}rPr/{{{ns}}}i'.format(ns=w))
				if italic_tag is not None:
					italic = italic_tag.get('{{{ns}}}val'.format(ns=w)) != '0' and italic_tag.get('{{{ns}}}val'.format(ns=w)) != 'false' and italic_tag.get('{{{ns}}}val'.format(ns=w)) != 'f'
				else:
					italic = None

				styles[style_id] = { 'bold' : bold, 'italic' : italic }
		print "Found: {0} styles".format(tally)
		# Step three, grab all paragraphs from document.xml
		markdown = []
		document = xml.etree.ElementTree.parse(pkg.open("word/document.xml")).getroot()
		search_str = './/{{{ns}}}p'.format(ns=w)
		print search_str
		tally = 0
		for paragraph in document.findall(search_str):
			tally += 1
			paragraph_style = { 'bold' : False, 'italic' : False }

			style_id = paragraph.get('{{{ns}}}pStyle'.format(ns=w))
			if style_id:
				referenced_style = styles[style_id]
				if referenced_style:
					paragraph_style['bold'] = referenced_style['bold']
					paragraph_style['italic'] = referenced_style['italic']

			if paragraph_style['italic']:
				markdown.append("_")
			if paragraph_style['bold']:
				markdown.append("**")

			# Step four, grab all runs within each paragraph
			for run in paragraph.findall('.//{{{ns}}}r'.format(ns=w)):
				# Step five, apply styles to each run & render all <w:t> and <w:br> tags
				# Get run_style id for any document-level styles
				run_style = { 'bold' : False, 'italic' : False }

				style_id = run.get('{{{ns}}}rStyle'.format(ns=w))
				if style_id:
					referenced_style = styles[style_id]
					if referenced_style:
						run_style['bold'] = referenced_style['bold']
						run_style['italic'] = referenced_style['italic']
				
				# Get rPr for any local styles
				bold_tag = run.find('.//{{{ns}}}rPr/{{{ns}}}b'.format(ns=w))
				if bold_tag is not None:
					inline_bold = bold_tag.get('{{{ns}}}val'.format(ns=w)) != '0' and bold_tag.get('{{{ns}}}val'.format(ns=w)) != 'false' and bold_tag.get('{{{ns}}}val'.format(ns=w)) != 'f'
				else:
					inline_bold = None
				
				italic_tag = run.find('.//{{{ns}}}rPr/{{{ns}}}i'.format(ns=w))
				if italic_tag is not None:
					inline_italic = italic_tag.get('{{{ns}}}val'.format(ns=w)) != '0' and italic_tag.get('{{{ns}}}val'.format(ns=w)) != 'false' and italic_tag.get('{{{ns}}}val'.format(ns=w)) != 'f'
				else:
					inline_italic = None

				# Combine styles to come up with whether or not we're actually supposed to be bold or italic here
				if inline_bold != None:
					run_style['bold'] = inline_bold
				if inline_italic != None:
					run_style['italic'] = inline_italic				

				# ' _**', ' **', ' _', '**_ ', '** ', '_ ', '** _', '_ **'
				prepend = ''
				append = ''

				if not paragraph_style['bold'] and run_style['bold'] == True and not paragraph_style['italic'] and run_style['italic'] == True:
					prepend = ' _**'
					append = '**_ '
				elif not paragraph_style['bold'] and run_style['bold'] == True and not paragraph_style['italic'] and not run_style['italic']:
					prepend = ' **'
					append = '** '
				elif not paragraph_style['bold'] and not run_style['bold'] and not paragraph_style['italic'] and run_style['italic'] == True:
					prepend = ' _'
					append = '_ '
				elif paragraph_style['bold'] == True and run_style['bold'] == False and paragraph_style['italic'] == True and run_style['italic'] == False:
					prepend = '**_ '
					append = ' _**'
				elif paragraph_style['bold'] == True and run_style['bold'] == False and not paragraph_style['italic'] and not run_style['italic']:
					prepend = '** '
					append = ' **'
				elif not paragraph_style['bold'] and not run_style['bold'] and paragraph_style['italic'] == True and run_style['italic'] == False:
					prepend = '_ '
					append =' _'
				elif paragraph_style['bold'] == True and run_style['bold'] == False and not paragraph_style['italic'] and run_style['italic'] == True:
					prepend = '** _'
					append = '_ **'
				elif not paragraph_style['bold'] and run_style['bold'] == True and paragraph_style['italic'] == True and run_style['italic'] == False:
					prepend = '_ **'
					append = '** _'

				markdown.append(prepend)

				for child in run:
					if child.tag == '{{{ns}}}t'.format(ns=w):
						markdown.append(child.text)
					elif child.tag == '{{{ns}}}br'.format(ns=w):
						markdown.append("  \n")

				markdown.append(append)

			if paragraph_style['bold']:
				markdown.append("**")
			if paragraph_style['italic']:
				markdown.append("_")		 
			markdown.append("\n")

		print "Found {0} paragraphs".format(tally)
		# Concatenate all text from all runs from all paragraphs and output it!
		out = ''.join(markdown)
		clipboard.set(out)		
	
if __name__ == '__main__':
	main()