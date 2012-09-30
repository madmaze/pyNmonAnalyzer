#!/usr/bin/env python

class nmonParser:
	fname = ""
	outdir = ""
	rawdata = []
	outData = {}
	
	sysInfo=[]
	bbbInfo=[]
	tStamp={}
	
	def __init__(self, fname="./test.nmon",outdir=""):
		self.fname = fname
		self.outdir = outdir
	
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
			self.tStamp[line[1]]=line[2]
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
			#if bits[1] == "T0002":
			#	for l in self.outData.keys():
			#		print self.outData[l]
			#	exit()
		#self.parseSysInfo()
		#self.parseCols()
		#self.parseBBBP()
		for l in self.outData.keys():
			print l, self.outData[l]
		#self.parseSnapshots()
		

