pyNmonAnalyzer
========

A tool for parsing and reshuffeling nmon's output into "normal" csv format.
Nmon puts out a long file with a system-info header at the beginning, followed
by a continuous stream of time stamped readings. This format makes it difficult
for analysis by standard Spreadsheet viewers without a fair amount of preprocessing.
The pyNmonAnalyzer aims to make this simpler, faster and more effective. In one
sweep the pyNmonAnalyzer creates CSV files and two HTML-based reports, one with static 
graphs and one with interactive graphs powered by [dygraphs](http://dygraphs.com). This 
project is currently a work-in-progress and therefore will hopefully improve in 
functionality and usability. If you have questions, find a bug or have feature requests, please do
file an issues [here](https://github.com/madmaze/pyNmonAnalyzer/issues)

- [Example Report](http://matthiaslee.com/scratch/pyNmonAnalyzer/data/report.html)
- [Example Advanced Report](http://matthiaslee.com/scratch/pyNmonAnalyzer/interactiveReport.html)


Goals:
-----
- make nmon log file analysis easier and faster
- create HTML based reports with embedded graphs
- create CSV files for more indepth data analysis
- interactive graphs for hands on analysis, perhaps using dygraph.js

Getting started:
-----
- [Python Nmon Analyzer: moving away from excel macros](http://matthiaslee.com/?q=node/38)

Usage:
-----
```
usage: pyNmonAnalyzer.py [-h] [-x] [-d] [-i INPUT_FILE] [-o OUTDIR] [-c] [-b]
                         [-r CONFFNAME] [--defaultConfig]

nmonParser converts NMON monitor files into time-sorted CSV/Spreadsheets for
easier analysis, without the use of the MS Excel Macro. Also included is an
option to build an HTML report with graphs, which is configured through
report.config.

optional arguments:
  -h, --help            show this help message and exit
  -x, --overwrite       overwrite existing results (Default: False)
  -d, --debug           debug? (Default: False)
  -i INPUT_FILE, --inputfile INPUT_FILE
                        Input NMON file
  -o OUTDIR, --output OUTDIR
                        Output dir for CSV (Default: ./report/)
  -c, --csv             CSV output? (Default: False)
  -b, --buildReport     report output? (Default: False)
  --buildInteractiveReport
                        Compile interactive report? (Default: False)
  -r CONFFNAME, --reportConfig CONFFNAME
                        Report config file, if none exists: we will write the
                        default config file out (Default: ./report.config)
  --defaultConfig       Write out a default config file
  -l LOGLEVEL, --log LOGLEVEL
                        Logging verbosity, use DEBUG for more output and
                        showing graphs (Default: INFO)
```

Example Usage:
-------------
Build HTML report with *interactive* graphs for test.nmon and store results to testReport  
```$> ./pyNmonAnalyzer.py -b -o testReport -i test.nmon```

Build HTML report with graphs for test.nmon and store results to testReport  
```$> ./pyNmonAnalyzer.py -b -o testReport -i test.nmon```

Compile CSV formatted tables for data in test.nmon and store results to testOut  
```$> ./pyNmonAnalyzer.py -c -o testOut -i test.nmon```

License:
-------
```
Copyright (c) 2012-2013 Matthias Lee, matthias.a.lee[]gmail.com
Last edited: June 24th 2013

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
