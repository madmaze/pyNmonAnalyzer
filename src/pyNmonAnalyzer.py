#!/usr/bin/env python
'''
Copyright (c) 2012-2017 Matthias Lee

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
from __future__ import print_function
import os
import sys
from shutil import rmtree 
import argparse
import logging as log

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
	
	stdReport = [('CPU_ALL', ['user', 'sys', 'wait'], 'stackedGraph: true, fillGraph: true'), ('DISKBUSY', ['sda1', 'sdb1'], ''), ('MEM', ['memtotal', 'active'], ''), ('NET', ['eth0'], '')]
	
	def __init__(self, args=None, raw_args=None):
		if args is None and raw_args is None:
			log.error("args and rawargs cannot be None.")
			sys.exit()
		if args is None:
			self.args = self.parseargs(raw_args)
		else:
			self.args = args
			
		if self.args.defaultConf:
			# write out default report and exit
			log.warn("Note: writing default report config file to " + self.args.confFname)
			self.saveReportConfig(self.stdReport, configFname=self.args.confFname)
			sys.exit()
		
		if self.args.buildReport:
			# check whether specified report config exists
			if os.path.exists(self.args.confFname) == False:
				log.warn("looks like the specified config file(\""+self.args.confFname+"\") does not exist.")
				ans = raw_input("\t Would you like us to write the default file out for you? [y/n]:")
				
				if ans.strip().lower() == "y":
					self.saveReportConfig(self.stdReport, configFname=self.args.confFname)
					log.warn("Wrote default config to report.config.")
					log.warn("Please adjust report.config to ensure the correct devices will be graphed.")
				else:
					log.warn("\nNOTE: you could try using the default config file with: -r report.config")
				sys.exit()
		
		# check ouput dir, if not create
		if os.path.exists(self.args.outdir) and self.args.overwrite:
			try:
				rmtree(self.args.outdir)
			except:
				log.error("Removing old dir:",self.args.outdir)
				sys.exit()
				
		elif os.path.exists(self.args.outdir):
			log.error("Results directory already exists, please remove or use '-x' to overwrite")
			sys.exit()
			
		# Create results path if not existing
		try:
			os.makedirs(self.args.outdir)
		except:
			log.error("Creating results dir:", self.args.outdir)
			sys.exit()
		
		# This is where the magic begins
		self.nmonParser = pyNmonParser.pyNmonParser(self.args.input_file, self.args.outdir, self.args.overwrite)
		self.processedData = self.nmonParser.parse()
		
		if self.args.outputCSV or "inter" in self.args.reportType.lower():
			log.info("Preparing CSV files..")
			self.outputData("csv")
		if self.args.buildReport:
			if "stat" in self.args.reportType.lower():
				log.info("Preparing static Report..")
				self.buildReport()
			elif "inter" in self.args.reportType.lower():
				log.info("Preparing interactive Report..")
				self.buildInteractiveReport(self.processedData, self.args.dygraphLoc)
			else:
				log.error("Report type: \"%s\" is not recognized" % self.args.reportType)
				sys.exit()
		
		log.info("All done, exiting.")
	
	def parseargs(self, raw_args):
		parser = argparse.ArgumentParser(description="nmonParser converts NMON monitor files into time-sorted CSV/Spreadsheets for easier analysis, without the use of the MS Excel Macro. Also included is an option to build an HTML report with graphs, which is configured through report.config.")
		parser.add_argument("-x","--overwrite", action="store_true", dest="overwrite", help="overwrite existing results (Default: False)")
		parser.add_argument("-d","--debug", action="store_true", dest="debug", help="debug? (Default: False)")
		parser.add_argument("--force", action="store_true", dest="force", help="force using of config (Default: False)")
		parser.add_argument("-i","--inputfile",dest="input_file", default="test.nmon", help="Input NMON file")
		parser.add_argument("-o","--output", dest="outdir", default="./report/", help="Output dir for CSV (Default: ./report/)")
		parser.add_argument("-c","--csv", action="store_true", dest="outputCSV", help="CSV output? (Default: False)")
		parser.add_argument("-b","--buildReport", action="store_true", dest="buildReport", help="report output? (Default: False)")
		parser.add_argument("-t","--reportType", dest="reportType", default="interactive", help="Should we be generating a \"static\" or \"interactive\" report (Default: interactive)")
		parser.add_argument("-r","--reportConfig", dest="confFname", default="./report.config", help="Report config file, if none exists: we will write the default config file out (Default: ./report.config)")
		parser.add_argument("--dygraphLocation", dest="dygraphLoc", default="http://dygraphs.com/1.1.0/dygraph-combined.js", help="Specify local or remote location of dygraphs library. This only applies to the interactive report. (Default: http://dygraphs.com/1.1.0/dygraph-combined.js)")
		parser.add_argument("--defaultConfig", action="store_true", dest="defaultConf", help="Write out a default config file")
		parser.add_argument("-l","--log",dest="logLevel", default="INFO", help="Logging verbosity, use DEBUG for more output and showing graphs (Default: INFO)")
		args = parser.parse_args(raw_args)
		
		if len(sys.argv) == 1:
			# no arguments specified
			parser.print_help()
			sys.exit()
		
		logLevel = getattr(log, args.logLevel.upper())
		if logLevel is None:
			print("ERROR: Invalid logLevel:", args.loglevel)
			sys.exit()
		if args.debug:
			log.basicConfig(level=logLevel, format='%(asctime)s - %(filename)s:%(lineno)d - %(levelname)s - %(message)s')
		else:
			log.basicConfig(level=logLevel, format='%(levelname)s - %(message)s')
		
		return args
	
	def saveReportConfig(self, reportConf, configFname="report.config"):
		# TODO: add some error checking
		f = open(configFname,"w")
		header = '''
# Plotting configuration file.
# =====
# Please edit this file carefully, generally the CPU and MEM options are left with 
# 	their defaults. For the static report, these have special under the hood calculations
#   to give you the used memory vs total memory instead of free vs total.
# For the Interactive report, the field names are used to pic out the right field from the CSV
# files for plotting.
# 
# Do adjust DISKBUSY and NET to plot the desired data
#
# Defaults for Linux Systems:
# CPU_ALL=user,sys,wait{stackedGraph: true, fillGraph: true}
# DISKBUSY=sda1,sdb1{}
# MEM=memtotal,active{}
# NET=eth0{}
#
# Defaults for AIX Systems
# CPU_ALL=user,sys,wait{stackedGraph: true, fillGraph: true}
# DISKBUSY=hdisk1,hdisk10{}
# MEM=Real total(MB),Real free(MB){}
# NET=en2{}

'''
		f.write(header)
		for stat, fields, plotOpts in reportConf:
			line = stat + "="
			if len(fields) > 0:
				line += ",".join(fields)
			line += "{%s}\n" % plotOpts
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
					# interactive/dygraph report options
					optStart=-1
					optEnd=-1
					if ("{" in bits[1]) != ("}" in bits[1]):
						log.error("Failed to parse, {..} mismatch")
					elif "{" in bits[1] and "}" in bits[1]:
						optStart=bits[1].find("{")+1
						optEnd=bits[1].rfind("}")
						plotOpts=bits[1][optStart:optEnd].strip()
					else:
						plotOpts = ""
						
					stat = bits[0]
					if bits[1] != "":
						if optStart != -1:
							fields = bits[1][:optStart-1].split(",")
						else:
							fields = bits[1].split(",")
						
					if self.args.debug:
						log.debug("%s %s" % (stat, fields))
						
					# add to config
					reportConfig.append((stat,fields,plotOpts))
					
		f.close()
		return reportConfig
	
	def buildReport(self):
		nmonPlotter = pyNmonPlotter.pyNmonPlotter(self.processedData, self.args.outdir, debug=self.args.debug)
				
		# Note: CPU and MEM both have different logic currently, so they are just handed empty arrays []
		#       For DISKBUSY and NET please do adjust the columns you'd like to plot
		
		if os.path.exists(self.args.confFname):
			reportConfig = self.loadReportConfig(configFname=self.args.confFname)
		else:
			log.error("something went wrong.. looks like %s is missing. run --defaultConfig to generate a template" % (self.args.confFname))
			sys.exit()
		
		if self.isAIX():
			# check whether a Linux reportConfig is being used on an AIX nmon file
			wrongConfig = False
			indicators = {"DISKBUSY":"sd","NET":"eth","MEM":"memtotal"}
			for cat,param,_ in reportConfig:
				if cat in indicators and indicators[cat] in param:
					wrongConfig=True
			
			if wrongConfig:
				if not self.args.force:
					log.error("It looks like you might have the wrong settings in your report.config.")
					log.error("From what we can see you have settings for a Linux system but an nmon file of an AIX system")
					log.error("if you want to ignore this error, please use --force")
					sys.exit()
		
		# TODO implement plotting options
		outFiles = nmonPlotter.plotStats(reportConfig, self.isAIX())
		
		# Build HTML report
		pyNmonReport.createReport(outFiles, self.args.outdir)
	
	def isAIX(self):
		#TODO: find better test to see if it is AIX
		if "PROCAIO" in self.processedData:
			return True
		return False
	
	def buildInteractiveReport(self, data, dygraphLoc):
		# Note: CPU and MEM both have different logic currently, so they are just handed empty arrays []
		#       For DISKBUSY and NET please do adjust the collumns you'd like to plot
		
		if os.path.exists(self.args.confFname):
			reportConfig = self.loadReportConfig(configFname=self.args.confFname)
		else:
			log.error("something went wrong.. looks like %s is missing. run --defaultConfig to generate a template" % (self.args.confFname))
			sys.exit()

		# Build interactive HTML report using dygraphs
		pyNmonReport.createInteractiveReport(reportConfig, self.args.outdir, data=data, dygraphLoc=dygraphLoc)
			
		
	def outputData(self, outputFormat):
		self.nmonParser.output(outputFormat)
		
if __name__ == "__main__":
	_ = pyNmonAnalyzer(raw_args=sys.argv[1:])
	
