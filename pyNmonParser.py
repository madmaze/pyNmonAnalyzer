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

class pyNmonParser:
	fname = ""
	outdir = ""

	# Holds final 2D arrays of each stat
	processedData = {}
	# Holds System Info gathered by nmon
	sysInfo=[]
	bbbInfo=[]
	# Holds timestamps for later lookup
	tStamp={}
	
	def __init__(self, fname="./test.nmon",outdir="./data/",overwrite=False,debug=False):
		# TODO: check input vars or "die"
		self.fname = fname
		self.outdir = outdir
		self.debug = debug
		
	def outputCSV(self, stat):
		outFile = open(os.path.join(self.outdir,stat+".csv"),"w")
		line=""
		# Iterate over each row
		for n in range(len(self.processedData[stat][0])):
			line=""
			# Iterate over each column
			for col in self.processedData[stat]:
				if line == "":
					line+=col[n]
				else:
					line+=","+col[n]
			outFile.write(line+"\n")
	
	def processLine(self,header,line):
		if "AAA" in header:
			# we are looking at the basic System Specs
			self.sysInfo.append(line[1:])
		elif "BBB" in header:
			# more detailed System Spec
			# do more grandular processing
			# refer to pg 11 of analyzer handbook
			self.bbbInfo.append(line)
		elif "ZZZZ" in header:
			self.tStamp[line[1]]=line[3]+" "+line[2]
		else:
			if line[0] in self.processedData.keys():
				#print "already here"
				table=self.processedData[line[0]]
				for n,col in enumerate(table):
					# line[1] give you the T####
					if line[n+1] in self.tStamp.keys():
						# lookup the time stamp in tStamp
						col.append(self.tStamp[line[n+1]])
					else:
						# TODO: do parsing(str2float) here
						col.append(line[n+1])
						# this should always be a float
						#try:
						#	col.append(float(line[n+1]))
						#except:
						#	print line[n+1]
						#	col.append(line[n+1])
					
			else:
				# new collumn, hoping these are headers
				header=[]
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
			print "Error: output called before parsing"
			exit()
			
		if outType.lower()=="csv":
			# Write out to multiple CSV files
			for l in self.processedData.keys():
				self.outputCSV(l)
		else:
			print "Error: output type: %s has not been implemented." % (outType)
			exit()

		

