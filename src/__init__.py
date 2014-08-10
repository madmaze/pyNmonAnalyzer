'''
Copyright (c) 2012-2014 Matthias Lee

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

from pynmonanalyzer import pyNmonAnalyzer as pna
import sys
import argparse
import logging as log

def main(args=None):

	parser = argparse.ArgumentParser(description="nmonParser converts NMON monitor files into time-sorted CSV/Spreadsheets for easier analysis, without the use of the MS Excel Macro. Also included is an option to build an HTML report with graphs, which is configured through report.config.")
	parser.add_argument("-x","--overwrite", action="store_true", dest="overwrite", help="overwrite existing results (Default: False)")
	parser.add_argument("-d","--debug", action="store_true", dest="debug", help="debug? (Default: False)")
	parser.add_argument("-i","--inputfile",dest="input_file", default="test.nmon", help="Input NMON file")
	parser.add_argument("-o","--output", dest="outdir", default="./report/", help="Output dir for CSV (Default: ./report/)")
	parser.add_argument("-c","--csv", action="store_true", dest="outputCSV", help="CSV output? (Default: False)")
	parser.add_argument("-b","--buildReport", action="store_true", dest="buildReport", help="report output? (Default: False)")
	parser.add_argument("--buildInteractiveReport", action="store_true", dest="buildInteractiveReport", help="Compile interactive report? (Default: False)")
	parser.add_argument("-r","--reportConfig", dest="confFname", default="./report.config", help="Report config file, if none exists: we will write the default config file out (Default: ./report.config)")
	parser.add_argument("--dygraphLocation", dest="dygraphLoc", default="http://dygraphs.com/dygraph-dev.js", help="Specify local or remote location of dygraphs library. This only applies to the interactive report. (Default: http://dygraphs.com/dygraph-dev.js)")
	parser.add_argument("--defaultConfig", action="store_true", dest="defaultConf", help="Write out a default config file")
	parser.add_argument("-l","--log",dest="logLevel", default="INFO", help="Logging verbosity, use DEBUG for more output and showing graphs (Default: INFO)")
	args = parser.parse_args()
	
	if len(sys.argv) == 1:
		# no arguments specified
		parser.print_help()
		sys.exit()
	
	logLevel = getattr(log, args.logLevel.upper())
	if logLevel is None:
		print "ERROR: Invalid logLevel:", args.loglevel
		sys.exit()
	if args.debug:
		log.basicConfig(level=logLevel, format='%(asctime)s - %(filename)s:%(lineno)d - %(levelname)s - %(message)s')
	else:
		log.basicConfig(level=logLevel, format='%(levelname)s - %(message)s')

	_ = pna.pyNmonAnalyzer(args)

if __name__ == '__main__':
	sys.exit(main())