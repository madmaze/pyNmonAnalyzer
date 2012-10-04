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
import os
from shutil import rmtree 
import pyNmonParser
import pyNmonPlotter
import argparse

class pyNmonAnalyzer:
	# Holds final 2D arrays of each stat
	processedData = {}
	
	nmonParser=""
	
	# Holds System Info gathered by nmon
	sysInfo=[]
	bbbInfo=[]
	args=[]
	
	def __init__(self, args):
		self.args=args
		# check ouput dir, if not create
		if os.path.exists(self.args.outdir) and args.overwrite:
			try:
				rmtree(self.args.outdir)
			except:
				print "[ERROR] removing old dir:",self.args.outdir
				exit()
				
		elif os.path.exists(self.args.outdir):
			print "[ERROR] results directory already exists, please remove or use '-x' to overwrite"
			exit()
			
		# Create results path if not existing
		try:
			os.makedirs(self.args.outdir)
		except:
			print "[ERROR] creating results dir:",self.args.outdir
			exit()
		
		# This is where the magic begins
		self.nmonParser = pyNmonParser.pyNmonParser(args.input_file, args.outdir, args.overwrite)
		self.processedData = self.nmonParser.parse()
		
		if self.args.outputCSV:
			print "Preparing CSV files.."
			self.outputData("csv")
		if self.args.buildReport:
			print "Preparing graphs.."
			self.buildReport()
		
		print "\nAll done, exiting."
	
	def buildReport(self):
		nmonPlotter = pyNmonPlotter.pyNmonPlotter(self.processedData, args.outdir, debug=self.args.debug)
		stdReport = ["CPU","DISKBUSY","MEM","NET"]
		
		# TODO implement plotting options
		nmonPlotter.plotStats(stdReport)
		
	def outputData(self, outputFormat):
		self.nmonParser.output(outputFormat)
		
if __name__ == "__main__":
	parser = argparse.ArgumentParser(description="nmonParser converts NMON monitor files into time-sorted CSV/Spreadsheets for easier analysis, without the use of the MS Excel Macro")
	parser.add_argument("-x","--overwrite", action="store_true", dest="overwrite", help="overwrite existing results (Default: False)")
	parser.add_argument("-d","--debug", action="store_true", dest="debug", help="debug? (Default: False)")
	parser.add_argument("input_file", default="test.nmon", help="Input NMON file")
	parser.add_argument("-o","--output", dest="outdir", default="./data/", help="Output dir for CSV (Default: ./data/)")
	parser.add_argument("-c","--csv", action="store_true", dest="outputCSV", help="CSV output? (Default: False)")
	parser.add_argument("-b","--buildReport", action="store_true", dest="buildReport", help="report output? (Default: False)")
	args = parser.parse_args()
	
	nmonAnalyzer=pyNmonAnalyzer(args)
	