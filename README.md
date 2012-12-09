pyNmonAnalyzer
========

A tool for parsing and reshuffeling nmon's output into "normal" csv format.
Nmon puts out a long file with a system-info header at the beginning, followed
by a continuous stream of time stamped readings. This format makes it difficult
for analysis by standard Spreadsheet viewers without a fair amount of preprocessing.
The pyNmonAnalyzer aims to make this simpler, faster and more effective. In one
sweep the pyNmonAnalyzer creates CSV files and an HTML report with graphs. This 
project is currently a work-in-progress and therefore will hopefully improve in 
functionality and usability. If you find a bug or have feature requests, please do
file an issues [here](https://github.com/madmaze/pyNmonAnalyzer/issues)



Goals:
-----
- make nmon log file analysis easier and faster
- create HTML based reports with embedded graphs
- create CSV files for more indepth data analysis
- (way off in the future) interactive graphs for hands on analysis, perhaps using dygraph.js

Getting started:
-----
- [Python Nmon Analyzer: moving away from excel macros](http://matthiaslee.com/?q=node/38)

Usage:
-----
```
usage: pyNmonAnalyzer.py [-h] [-x] [-d] [-o OUTDIR] [-c] [-b] [-r CONFFNAME]
                         input_file

nmonParser converts NMON monitor files into time-sorted CSV/Spreadsheets for
easier analysis, without the use of the MS Excel Macro. Also included is an
option to build an HTML report with graphs, which is configured through
report.config.

positional arguments:
  input_file            Input NMON file

optional arguments:
  -h, --help            show this help message and exit
  -x, --overwrite       overwrite existing results (Default: False)
  -d, --debug           debug? (Default: False)
  -o OUTDIR, --output OUTDIR
                        Output dir for CSV (Default: ./data/)
  -c, --csv             CSV output? (Default: False)
  -b, --buildReport     report output? (Default: False)
  -r CONFFNAME, --reportConfig CONFFNAME
                        Report config file, if none exists: we will write the
                        default config file out (Default: ./report.config)
```

License:
--------
```
Copyright (c) 2012 Matthias Lee, matthias.a.lee[]gmail.com
Last edited: Sept 25th 2012

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
```
