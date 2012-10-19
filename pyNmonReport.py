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
import datetime

htmlheader='''<html>
<head><title>pyNmonReport %s </title></head>
<body>
<table>	
''' % (datetime.datetime.now())
	
def createReport(outFiles, outPath, fname="report.html"):
	reportPath = os.path.join(outPath,fname)
	report = open(reportPath, "w")
	
	# write out the html header
	report.write(htmlheader)
	
	for f in outFiles:
		#print os.path.relpath(f,outPath)
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
