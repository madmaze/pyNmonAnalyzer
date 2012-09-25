#!/usr/bin/env python

class nmonParser:
	fname = ""
	outdir = ""
	data = []
	dataPtr=0
	sysInfo=[]
	def __init__(self, fname="./nmon",outdir=""):
		self.fname = fname
		self.outdir = outdir
	
	def parseSysInfo(self):
		for l in self.data:
			if "AAA," in l:
				# TODO: Strip row heading
				self.sysInfo.append(l.strip())
				self.dataPtr+=1
			else:
				#if we have finished reading AAA break out 
				break
	def parseCols(self):
		for l in self.data[self.dataPtr:]:
			if "BBBP," in l:
				break
			else:
				self.dataPtr+=1
				print l.strip()

	def parse(self):
		# TODO: check fname
		f = open(self.fname,"r")
		self.data = f.readlines()
		self.parseSysInfo()
		self.parseCols()


