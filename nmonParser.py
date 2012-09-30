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
import matplotlib.pyplot as plt

class nmonParser:
	fname = ""
	outdir = ""
	rawdata = []
	
	# Holds final 2D arrays of each stat
	outData = {}
	# Holds System Info gathered by nmon
	sysInfo=[]
	bbbInfo=[]
	# Holds timestamps for later lookup
	tStamp={}
	
	def __init__(self, fname="./test.nmon",outdir="./data/",overwrite=False):
		self.fname = fname
		self.outdir = outdir
		# check ouput dir, if not create
		if os.path.exists(self.outdir) and overwrite:
			try:
				rmtree(self.outdir)
			except:
				print "[ERROR] removing old dir:",self.outdir
				exit()
				
		elif os.path.exists(self.outdir):
			print "[ERROR] results directory already exists, please remove or use '-x' to overwrite"
			exit()
			
		# Create results path if not existing
		try:
			os.makedirs(self.outdir)
		except:
			print "[ERROR] creating results dir:",self.outdir
			exit()
		
	def outputCSV(self, stat):
		outFile = open(os.path.join(self.outdir,stat+".csv"),"w")
		line=""
		# Iterate over each row
		for n in range(len(self.outData[stat][0])):
			line=""
			# Iterate over each column
			for col in self.outData[stat]:
				if line == "":
					line+=col[n]
				else:
					line+=","+col[n]
			outFile.write(line+"\n")
	
	def plotStat(self, stat):
		fig = plt.figure(figsize=(10,6))
		plot = fig.add_subplot(1,1,1)
		cnt = len(self.outData["CPU01"][0][1:])
		plot.plot(self.outData["CPU01"][1][1:])
		
		lcnt=10
		ticks=[]
		labels=[]
		for n,l in enumerate(self.outData["CPU01"][0][1:]):
			if n==cnt or n%lcnt==0:
				labels.append(l)
				ticks.append(n)
				
		plt.xticks(ticks,labels, size='small', rotation=90)
		
		plot.set_ylabel("CPU usage")
		plot.set_xlabel("Time")
		plt.show()
	
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
			if line[0] in self.outData.keys():
				#print "already here"
				table=self.outData[line[0]]
				for n,col in enumerate(table):
					# line[1] give you the T####
					if line[n+1] in self.tStamp.keys():
						# lookup the time stamp in tStamp
						col.append(self.tStamp[line[n+1]])
					else:
						col.append(line[n+1])
					
			else:
				# new collumn, hoping these are headers
				header=[]
				for h in line[1:]:
					# make it an array
					tmp=[]
					tmp.append(h)
					header.append(tmp)
				self.outData[line[0]]=header
			
				
	def parse(self):
		# TODO: check fname
		f = open(self.fname,"r")
		self.rawdata = f.readlines()
		for l in self.rawdata:
			l=l.strip()
			bits=l.split(',')
			self.processLine(bits[0],bits)

		# Write out to multiple CSV files
		for l in self.outData.keys():
			self.outputCSV(l)

		

