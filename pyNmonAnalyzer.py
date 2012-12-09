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
import argparse

import pyNmonParser
import pyNmonPlotter
import pyNmonReport

class pyNmonAnalyzer:
	# Holds final 2D arrays of each stat
	processedData = {}
	nmonParser = None
	
	# Holds System Info gathered by nmon
	sysInfo = []
	bbbInfo = []
	args = []
	
	def __init__(self, args):
		self.args = args
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
	
	def saveReportConfig(self, reportConf, configFname="report.config"):
		# TODO: add some error checking
		f = open(configFname,"w")
		header = '''
# Plotting configuration file.
# =====
# please edit this file carefully, generally the CPU and MEM options are left blank
# 	since there is under the hood calculations going on to plot used vs total mem and 
#	CPU plots usr/sys/wait for all CPUs on the system
# Do adjust DISKBUSY and NET to plot the desired data
#
# Defaults:
# CPU=
# DISKBUSY=sda1,sdb1
# MEM=
# NET=eth0

'''
		f.write(header)
		for stat, fields in reportConf:
			line = stat + "="
			if len(fields) > 0:
				line += ",".join(fields)
			line += "\n"
			f.write(line)
		f.close()
	
	def loadReportConfig(self, configFname="report.config"):
		# TODO: add some error checking
		f = open(configFname, "r")
		reportConfig = []
		
		# loop over all lines
		for l in f:
			l = l.strip()
			stat=""
			fields = []
			# ignore lines beginning with #
			if l[0:1] != "#":
				bits = l.split("=")
				# check whether we have the right number of elements
				if len(bits) == 2:
					stat = bits[0]
					if bits[1] != "":
						fields = bits[1].split(",")
						
					if self.args.debug:
						print stat, fields
						
					# add to config
					reportConfig.append((stat,fields))
					
		f.close()
		return reportConfig
	
	def buildReport(self):
		nmonPlotter = pyNmonPlotter.pyNmonPlotter(self.processedData, args.outdir, debug=self.args.debug)
		
		stdReport = [("CPU",[]),("DISKBUSY",["sda1","sdb1"]),("MEM",[]),("NET",["eth0"])]
		# Note: CPU and MEM both have different logic currently, so they are just handed empty arrays []
		#       For DISKBUSY and NET please do adjust the collumns you'd like to plot
		
		if os.path.exists(self.args.confFname):
			reportConfig = self.loadReportConfig(configFname=self.args.confFname)
		else:
			# TODO: this could be broken out into a wizard or something
			print "WARNING: looks like the specified config file(\""+self.args.confFname+"\") does not exist."
			
			if os.path.exists("report.config") == False:
				ans = raw_input("\t Would you like us to write the default file out for you? [y/n]:")
				
				if ans.strip().lower() == "y":
					self.saveReportConfig(stdReport)
					print "\nwrote default config to report.config"
			else:
				print "\nNOTE: you could try using the default config file with: -r report.config"
				
			exit()
			
		
		# TODO implement plotting options
		outFiles = nmonPlotter.plotStats(reportConfig)
		
		# Build HTML report
		pyNmonReport.createReport(outFiles, self.args.outdir)
			
		
	def outputData(self, outputFormat):
		self.nmonParser.output(outputFormat)
		
if __name__ == "__main__":
	parser = argparse.ArgumentParser(description="nmonParser converts NMON monitor files into time-sorted CSV/Spreadsheets for easier analysis, without the use of the MS Excel Macro. Also included is an option to build an HTML report with graphs, which is configured through report.config.")
	parser.add_argument("-x","--overwrite", action="store_true", dest="overwrite", help="overwrite existing results (Default: False)")
	parser.add_argument("-d","--debug", action="store_true", dest="debug", help="debug? (Default: False)")
	parser.add_argument("input_file", default="test.nmon", help="Input NMON file")
	parser.add_argument("-o","--output", dest="outdir", default="./data/", help="Output dir for CSV (Default: ./data/)")
	parser.add_argument("-c","--csv", action="store_true", dest="outputCSV", help="CSV output? (Default: False)")
	parser.add_argument("-b","--buildReport", action="store_true", dest="buildReport", help="report output? (Default: False)")
	parser.add_argument("-r","--reportConfig", dest="confFname", default="./report.config", help="Report config file, if none exists: we will write the default config file out (Default: ./report.config)")
	args = parser.parse_args()
	
	nmonAnalyzer=pyNmonAnalyzer(args)
	