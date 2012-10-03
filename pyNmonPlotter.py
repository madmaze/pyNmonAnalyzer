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
import matplotlib.pyplot as plt
import matplotlib as mpl
import datetime
import numpy as np

class pyNmonPlotter:
	# Holds final 2D arrays of each stat
	processedData = {}
	
	def __init__(self, processedData, outdir="./data/", overwrite=False, debug=False):
		# TODO: check input vars or "die"
		self.imgPath = outdir
		self.debug = debug
		self.processedData = processedData
		
	def plotStats(self, todoList):
		if len(todoList) <= 0:
			print "Error: nothing to plot"
			exit()
		
		for stat in todoList:
			if "CPU" in stat:
				# parse NMON date/timestamps and produce datetime objects
				times = [datetime.datetime.strptime(d, "%d-%b-%Y %H:%M:%S") for d in self.processedData["CPU_ALL"][0][1:]]
				values=[]
				values.append((self.processedData["CPU_ALL"][1][1:],"usr"))
				values.append((self.processedData["CPU_ALL"][2][1:],"sys"))
				values.append((self.processedData["CPU_ALL"][3][1:],"wait"))
				data=(times,values)
				self.plotStat(data, xlabel="Time", ylabel="CPU load (%)", title="CPU vs Time", isPrct=True)
			elif "DISKBUSY" in stat:
				times = [datetime.datetime.strptime(d, "%d-%b-%Y %H:%M:%S") for d in self.processedData["CPU_ALL"][0][1:]]
				values=[]
				values.append((self.processedData["DISKBUSY"][1][1:],self.processedData["CPU_ALL"][1][:1]))
				data=(times,values)
				self.plotStat(data, xlabel="Time", ylabel="Disk Busy (%)", title="Disk Busy vs Time", isPrct=True)
			elif "MEM" in stat:
				# parse NMON date/timestamps and produce datetime objects
				times = [datetime.datetime.strptime(d, "%d-%b-%Y %H:%M:%S") for d in self.processedData["CPU_ALL"][0][1:]]
				values=[]
				
				mem=np.array(self.processedData["MEM"])
				# total - free - buffers - chache
				print np.array(mem[1][1:]).dtype
				print np.array(mem[5][1:]).dtype
				used = np.array(float(mem[1][1:])) - np.array(float(mem[5][1:]))
				print used, type(mem[1][1:]), mem[5][1:]
				exit()
				#used = np.subtract(used, mem[10][1:])
				#used = np.subtract(used, mem[13][1:])
				print mem[1][:1], mem[5][:1], mem[10][:1], mem[13][:1]
				values.append((used,"used mem"))
				values.append((mem[1][1:],"total mem"))
				data=(times,values)
				self.plotStat(data, xlabel="Time", ylabel="Memory in MB", title="Memory vs Time", isPrct=True, yrange=[0,max(mem[1][1:])*1.2])
		
	def plotStat(self, data, xlabel="time", ylabel="", title="title", isPrct=True, yrange=[0,100]):
		
		# figure dimensions
		fig = plt.figure(figsize=(10,6))
		ax = fig.add_subplot(1,1,1)
		
		# parse NMON date/timestamps and produce datetime objects
		times, values = data 
		#data =self.processedData[stat][1][1:]
		#data2 =self.processedData["CPU02"][1][1:]
		#dataAll =self.processedData["CPU_ALL"][1][1:]
		
		# plot
		for v,label in values:
			ax.plot_date(times, v, "-")
		#ax.plot_date(dates, data2, "-")
		#ax.plot_date(dates, dataAll, "-")
		
		# format axis
		ax.xaxis.set_major_locator(mpl.ticker.MaxNLocator(10))
		ax.xaxis.set_major_formatter(mpl.dates.DateFormatter("%m-%d %H:%M:%S"))
		ax.xaxis.set_minor_locator(mpl.ticker.MaxNLocator(100))
		ax.autoscale_view()
		if isPrct:
			ax.set_ylim([0,100])
		ax.grid(True)

		fig.autofmt_xdate()
		ax.set_ylabel(ylabel)
		ax.set_xlabel(xlabel)
		if self.debug:
			plt.show()
		else:
			plt.savefig(os.path.join(self.imgPath,title.replace (" ", "_")+".png"))


		

