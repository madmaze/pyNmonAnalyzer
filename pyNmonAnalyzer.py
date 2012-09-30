#!/usr/bin/env python
'''
Copyright (c) 2012 Matthias Lee

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
'''

import nmonParser
import argparse

if __name__ == "__main__":
	parser = argparse.ArgumentParser(description="nmonParser converts NMON monitor files into time-sorted CSV/Spreadsheets for easier analysis, without the use of the MS Excel Macro")
	parser.add_argument("-x","--overwrite", action="store_true", dest="overwrite", help="overwrite existing results (Default: False)")
	parser.add_argument("input_file", default="test.nmon", help="Input NMON file")
	parser.add_argument("-o","--output", dest="outdir", default="./data/", help="Output dir for CSV (Default: ./data/")
	args = parser.parse_args()
	
	# This is where the magic begins
	nmonP=nmonParser.nmonParser(args.input_file, args.outdir, args.overwrite)
	nmonP.parse()
	
	# TODO implement plotting options
