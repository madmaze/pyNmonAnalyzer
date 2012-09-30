#!/usr/bin/env python
import nmonParser
import argparse
if __name__ == "__main__":
	parser = argparse.ArgumentParser(description='''
nmonParser converts NMON monitor files into time-sorted CSV/Spreadsheets for easier analysis, without the use of the MS Excel Macro
''')
	
	parser.add_argument("-x","--overwrite", action="store_true", dest="overwrite", help="overwrite existing results (Default: False)")
	parser.add_argument("input_file",  default="test.nmon", help="Input NMON file")
	parser.add_argument("-o","--output", dest="outdir", default="./data/", help="Output dir for CSV (Default: ./data/")
	args = parser.parse_args()
	
	P=nmonParser.nmonParser(args.input_file, args.outdir, args.overwrite)
	P.parse()
