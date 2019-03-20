#/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import errno
import glob
import csv
	
import argparse

SUBMODULES = ["python-docx"]
TBL_HEADER_MAX_SIZE = 2

def _module_path():
	if "__file__" in globals():
		return os.path.dirname(os.path.realpath(__file__))
	
	return ""
		
def setup_env():
	global SUBMODULES
	
	this_path = _module_path()
	SUBMODULES = [os.path.join(this_path, sub) for sub in SUBMODULES]
	sys.path = SUBMODULES + sys.path

setup_env()

import docx

def mkdir_p(path):
	try:
		os.makedirs(path)
	except OSError as exc:  # Python >2.5
		if exc.errno == errno.EEXIST and os.path.isdir(path):
			pass
		else:
			raise
			
def parse_tbl(tbl, keep_header = False, header_size = -1, keep_newlines = False):
	"""
	Parse a python-docx table object *tbl* and return
	the data in csv-compatible format.
	"""
	res = []
	detect_header_size = (header_size == -1)
	header_size = 0

	if detect_header_size and not keep_header:
		header_size = 1

		if len(tbl.rows) < 2:
			return [[]]

		if (tbl.rows[0].cells[0]._tc == tbl.rows[1].cells[0]._tc):
			header_size = TBL_HEADER_MAX_SIZE

	for nrow, row in enumerate(tbl.rows[header_size:]):
		last_tc = None
		buf = []
		for cell in row.cells:
			# ignore merged and empty cells
			if (cell._tc != last_tc):
				text = cell.text
				if (not keep_newlines):
					text = cell.text.replace("\r\n", " ") \
							.replace("\n", " ")
				buf.append(text)
			last_tc = cell._tc
		res.append(buf)
	return res

def write_csv(lines, output_file):
	"""
	Write *lines* to *output_file*.
	"""
	with open(output_file, "w") as f:
		writer = csv.writer(f)
		writer.writerows(lines)
		
def tbl2csv(tbl, output_file, keep_header = False, header_size = 1, keep_newlines = False):
	"""
	Parse a python-docx table object *tbl* and write
	the result as *output_file* in CSV format.
	"""
	write_csv(parse_tbl(tbl, keep_header, header_size, keep_newlines), output_file)

def main(input_dir, output_dir = "", use_captions = True,
			keep_header = False, header_size = 1, keep_newlines = False):
	if not output_dir:
		output_dir = os.path.join(input_dir, "out")

	old_pwd = os.getcwd()
	os.chdir(input_dir)
	print(keep_header, header_size)

	found_docs = glob.glob("**/*.docx")

	# main processing routine
	for docx_filename in found_docs:
		print("processing {file}".format(file = docx_filename))
		out_dir = os.path.join(output_dir, docx_filename)
		mkdir_p(out_dir)
		tables = docx.Document(docx_filename).tables
		for n, tbl in enumerate(tables):
			if use_captions and hasattr(tbl, "caption") and tbl.caption:
				out_file = os.path.join(out_dir, 
					"{tbl_name}.csv".format(tbl_name = tbl.caption))
				print("found table {tbl_name}, saving as {out_file}"
					.format(tbl_name = tbl.caption, out_file = out_file))
			else:
				out_file = os.path.join(out_dir, 
					"{tbl_num}.csv".format(tbl_num = n))
				print("found table #{tbl_num}, saving as {out_file}"
					.format(tbl_num = n, out_file = out_file))

			tbl2csv(tbl, out_file, keep_header, header_size, keep_newlines)

	os.chdir(old_pwd)

if __name__ == "__main__":
	parser = argparse.ArgumentParser(description = "Convert docx tables to CSV files.")
	parser.add_argument("input_dir", metavar = "input_dir", type = str, 
                    help = "an input directory to process docx files")
	parser.add_argument("-o", metavar = "output_dir", type = str, 
                    help = "an output directory to save CSV files")
	parser.add_argument("-c", action = "store_false",
                    help = "convert all tables (not only those containing a caption)")
	parser.add_argument("-k", action = "store_true",
                    help = "keep table header in output files (default: false)")
	parser.add_argument("-n", action = "store_true",
                    help = "keep newlines in the table cells (default: false)")
	parser.add_argument("-s", metavar = "header_size", type = int, default = -1,
                    help = "a size of the table header (default: -1 (try detecting header size))")

	args = parser.parse_args()
	main(input_dir = args.input_dir,
		output_dir = args.o,
		use_captions = args.c,
		keep_header = args.k,
		header_size = args.s,
		keep_newlines = args.n)