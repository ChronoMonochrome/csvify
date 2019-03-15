#/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import errno
import glob

from docx import Document

def _module_path():
	if hasattr(sys, "frozen"):
		return os.path.dirname(sys.executable)

	if "__file__" in globals():
		return os.path.dirname(__file__)
	
	return "."
	
def setup_env():	
	this_path = _module_path()

	if not this_path in sys.path:
		sys.path.append(this_path)
	
def mkdir_p(path):
	try:
		os.makedirs(path)
	except OSError as exc:  # Python >2.5
		if exc.errno == errno.EEXIST and os.path.isdir(path):
			pass
		else:
			raise
			
def parse_tbl(tbl):
	"""
	Parse a python-docx table object *tbl* and return
	the data in csv-compatible format.
	"""
	
	return [[cell.text for cell in iter(row.cells)]
	                   for row  in tbl.rows]
					   
def write_csv(lines, output_file):
	"""
	Write *lines* to *output_file*.
	"""
	with open(output_file, "w") as f:
		writer = csv.writer(f)
		writer.writerows(lines)
		
def tbl2csv(tbl, output_file):
	"""
	Parse a python-docx table object *tbl* and write
	the result as *output_file* in CSV format.
	"""
	write_csv(parse_tbl(tbl), output_file)

def main(input_dir, output_dir = ""):
	setup_env()

	if not output_dir:
		output_dir = os.path.join(input_dir, "out")

	old_pwd = os.getcwd()
	os.chdir(input_dir)

	found_docs = glob.glob("**/*.docx")

	dirs = set([os.path.dirname(doc) for doc in found_docs])
	for d in iter(dirs):
		mkdir_p(os.path.join(output_dir, d))

	# main processing routine
	for docx_filename in found_docs:
		print("processing {file}".format(file = docx_filename))
		out_dir = os.path.join(output_dir, os.path.dirname(docx_filename))
		tables = Document(docx_filename).tables
		for tbl in tables:
			if hasattr(tbl, "caption") and tbl.caption:
				out_file = os.path.join(out_dir, 
								"{tbl_name}.csv".format(tbl_name = tbl.caption)
						   )
				print("found table {tbl_name}, saving as {out_file}"
						.format(tbl_name = tbl.caption, out_file = out_file))
				tbl2csv(tbl, out_file)

	os.chdir(old_pwd)
	
def usage():
	print("Usage: python word2csv.py input_dir [output_dir]")
	sys.exit(0)
	
if __name__ == "__main__":
	if len(sys.argv) < 2:
		usage()
		
	input_dir = sys.argv[1]
	output_dir = ""
	if len(sys.argv) > 2:
		output_dir = sys.argv[2]

	main(input_dir, output_dir)