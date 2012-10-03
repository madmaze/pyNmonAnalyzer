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
	
	def buildReport(self):
		# This is where the magic begins
		nmonParser=pyNmonParser.pyNmonParser(args.input_file, args.outdir, args.overwrite)
		self.processedData=nmonParser.parse()
		
		nmonPlotter=pyNmonPlotter.pyNmonPlotter(self.processedData, debug=self.args.debug)
		stdReport=["CPU","DISKBUSY","MEM","NET"]
		# TODO implement plotting options
		nmonPlotter.plotStats(stdReport)
		#nmonParser.output("csv")
		#nmonPlotter.plotStat("CPU")
		#nmonP.plotStat("CPU02")
		#nmonP.plotStat("CPU_ALL")
		#nmonP.plotStat("NET")
		
if __name__ == "__main__":
	parser = argparse.ArgumentParser(description="nmonParser converts NMON monitor files into time-sorted CSV/Spreadsheets for easier analysis, without the use of the MS Excel Macro")
	parser.add_argument("-x","--overwrite", action="store_true", dest="overwrite", help="overwrite existing results (Default: False)")
	parser.add_argument("-d","--debug", action="store_true", dest="debug", help="debug? (Default: False)")
	parser.add_argument("input_file", default="test.nmon", help="Input NMON file")
	parser.add_argument("-o","--output", dest="outdir", default="./data/", help="Output dir for CSV (Default: ./data/)")
	args = parser.parse_args()
	
	nmonAnalyzer=pyNmonAnalyzer(args)
	nmonAnalyzer.buildReport()
	