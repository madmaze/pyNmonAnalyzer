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
import datetime
import logging as log

htmlheader='''<html>
<head><title>pyNmonReport %s </title></head>
<body>
<table>	
''' % (datetime.datetime.now())
	
def createReport(outFiles, outPath, fname="report.html"):
	reportPath = os.path.join(outPath,fname)
	try:
		report = open(reportPath, "w")
	except:
		log.error("Could not open report file!")
		exit()
	
	# write out the html header
	report.write(htmlheader)
	
	for f in outFiles:
		report.write('''	<tr>
		<td><br /><br />
		<b><center>%s</center></b><br />
		<img src="%s" />
		</td>
	</tr>
		''' % ("".join(os.path.basename(f).split(".")[:-1]), os.path.relpath(f,outPath)))
	
	report.write('''</table>
</body>
</html>
''')
	report.close()

def createInteractiveReport(reportConfig, outPath, data=None, dygraphLoc="http://dygraphs.com/dygraph-dev.js",fname="interactiveReport.html", templateFile="interactiveReport.tpl.html"):
	pkgdir = os.path.abspath(os.path.dirname(__file__))
	templateFile = os.path.join(pkgdir,templateFile)
	if not os.path.exists(templateFile):
		log.error("Template file for interactive report went missing.. "+templateFile)
		exit()
		
	if data is None:
		log.error("createInteractiveReport was not passed any data to process.")
		exit()
	
	tplFile = open(templateFile,"r").readlines()
	
	reportPath = os.path.join(outPath,fname)
	try:
		report = open(reportPath, "w")
	except:
		log.error("Could not open report file!")
		exit()
	
	# collects paths to plottable corresponding CSV
	dataSources=[]
	displayCols=[]
	specialOpts=[]
	# This list collects every stat/config we verified plot as plottable
	verifiedConfigs=[]
	basepath=os.path.join(outPath,"csv")
	relpath="csv"
	for k in reportConfig:
		log.debug(k)
		# check path relative to us running
		candidatePath=os.path.join(basepath,k[0]+".csv")
		if os.path.exists(candidatePath):
			# add path relative to where the output is
			dataSources.append(os.path.join(relpath,k[0]+".csv"))
			verifiedConfigs.append(k)
			
			# add to display cols
			localMin=None
			localMax=None
			headings=[]
			
			# go through data, skip the first column which contains the dates
			for c in data[k[0]][1:]:
				for i in k[1]:
					# match anything that contains a key from reportConfig
					if i.lower() in c[0].lower():
						headings.append(c[0])
						numericArray = [ float(x) for x in c[1:] ]
						if max(numericArray) > localMax or localMax == None:
							localMax = max(numericArray)
						if  min(numericArray) < localMin or localMin == None:
							localMin = min(numericArray)
			displayCols.append(headings)
			localMin = (0.0 if localMin==None else localMin)
			localMax = (0.0 if localMax==None else localMax)
			
			if k[0] in ["CPU_ALL","DISKBUSY"]:
				# its a prct so FORCE range 0-105
				localMin=0.0
				localMax=105.0
			
			# bring in opts from config
			if k[2] != "":
				localOpts = ",\n".join([k[2],'valueRange: [%f, %f]' % (0.0, localMax*1.05)])
			else:
				localOpts = 'valueRange: [%f, %f]' % (0.0, localMax*1.05)
			specialOpts.append((localOpts))
		
		else:
			log.warn("Key: %s, was not found in the processed nmon data, please manually check the available stats." % k[0])
	
	# fill in place holders in template file
	for l in tplFile:
		if "[__datetime__]" in l:
			line = l.replace("[__datetime__]", str(datetime.datetime.now()))
		elif "[__plots__]" in l:
			line = ""
			for i in range(len(dataSources)):
				line += '<h2>'+verifiedConfigs[i][0]+'</h2><br />\n <div id="plot' + str(i) + '"  style="width:1000px; height:300px;">loading...</div> <br /><br /> \n'
		elif "[__dataSources__]" in l:
			line = ""
			for s in dataSources:
				if line == "":
					line += '"'+s+'"'
				else:
					line += ',\n"'+s+'"'
					
		elif "[__specialOpts__]" in l:
			line = ""
			for s in specialOpts:
				log.debug(s)
				if line == "":
					line += '{' + s + '}'
				else:
					line += ',\n{' + s + '}'
					
		elif "[__displayCols__]" in l:
			line = ""
			for s in displayCols:
				if line == "":
					line += '["' + '","'.join(s) + '"]'
				else:
					line += ',\n["' + '","'.join(s) + '"]'
		
		elif "[__dygraphLoc__]" in l:
			line = l.replace("[__dygraphLoc__]", dygraphLoc)
			
		else:
			line = l
					
		report.write(line)
		
	report.close()
