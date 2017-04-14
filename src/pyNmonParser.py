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

import os
import logging as log
import datetime

class pyNmonParser:
	
	def __init__(self, fname="./test.nmon",outdir="./data/",overwrite=False,debug=False):
		# TODO: check input vars or "die"
		self.fname = fname
		self.outdir = outdir
		self.debug = debug
		
		# Holds final 2D arrays of each stat
		self.processedData = {}
		# Holds System Info gathered by nmon
		self.sysInfo = []
		self.bbbInfo = []
		# Holds timestamps for later lookup
		self.tStamp = {}
		
		
	def outputCSV(self, stat):
		outFile = open(os.path.join(self.outdir,stat+".csv"),"w")
		line=""
		if len(self.processedData[stat]) > 0:
			# Iterate over each row
			for n in range(len(self.processedData[stat][0])):
				line=""
				# Iterate over each column
				for col in self.processedData[stat]:
					if line == "":
						# expecting first column to be date times
						if n == 0:
							# skip headings
							line+=col[n]
						else:
							tstamp = datetime.datetime.strptime(col[n], "%d-%b-%Y %H:%M:%S")
							line += tstamp.strftime("%Y-%m-%d %H:%M:%S")
					else:
						line+=","+col[n]
				outFile.write(line+"\n")
	
	def processLine(self,header,line):
		if "AAA" in header:
			# we are looking at the basic System Specs
			self.sysInfo.append(line[1:])
		elif "BBB" in header:
			# more detailed System Spec
			# do more granular processing
			# refer to pg 11 of analyzer handbook
			self.bbbInfo.append(line)
		elif "ZZZZ" in header:
			self.tStamp[line[1]]=line[3]+" "+line[2]
		else:
			if "TOP" in line[0] and len(line) > 3:
					# top lines are the only ones that do not have the timestamp
					# as the second column, therefore we rearrange for parsing.
					# kind of a hack, but so is the rest of this parsing
					pid = line[1]
					line[1] = line[2]
					line[2] = pid
					
			if line[0] in self.processedData.keys():
				table = self.processedData[line[0]]
				for n,col in enumerate(table):
					# line[1] give you the T####
					if n == 0 and line[n+1] in self.tStamp.keys():
						# lookup the time stamp in tStamp
						col.append(self.tStamp[line[n+1]])

					elif n == 0 and line[n+1] not in self.tStamp.keys():
						log.warn("Discarding line with missing Timestamp %s" % line)
						break
						
					else:
						# TODO: do parsing(str2float) here
						if len(line) > n+1:
							col.append(line[n+1])
						else:
							# somehow we are missing an entry here
							# As in we have a heading, but no data
							log.debug("We found more column titles than data for the category:"+line[0]+". This has been observed with some versions of NMON on AIX")
							log.debug("This tends to happen with the LPAR readings, double check whether your data makes sense, if so you can ignore this.")
							col.append("0")
						# this should always be a float
						#try:
						#	col.append(float(line[n+1]))
						#except:
						#	print line[n+1]
						#	col.append(line[n+1])
					
			else:
				# new column, hoping these are headers
				# We are expecting a header row like:
				# CPU01,CPU 1 the-gibson,User%,Sys%,Wait%,Idle%
				header=[]
				if "TOP" in line[0] and len(line) < 3:
					# For some reason Top has two header rows, the first with only
					# two columns and then the real one therefore we skip the first row
					pass
				else:
					for h in line[1:]:
						# make it an array
						tmp=[]
						tmp.append(h)
						header.append(tmp)
					self.processedData[line[0]]=header
			
	def parse(self):
		# TODO: check fname
		f = open(self.fname,"r")
		rawdata = f.readlines()
		for l in rawdata:
			l=l.strip()
			bits=l.split(',')
			self.processLine(bits[0],bits)

		return self.processedData
	
	def output(self,outType="csv"):
		
		if len(self.processedData) <= 0:
			# nothing has been parsed yet
			log.error("Output called before parsing")
			exit()
		
		# make output dir
		self.outdir = os.path.join(self.outdir,outType)
		if not (os.path.exists(self.outdir)):
			try:
				os.makedirs(self.outdir)
			except:
				log.error("Creating results dir:",self.outdir)
				exit()
				
		# switch for different output types	
		if outType.lower()=="csv":
			# Write out to multiple CSV files
			for l in self.processedData.keys():
				self.outputCSV(l)
		else:
			log.error("Output type: %s has not been implemented." % (outType))
			exit()

		

