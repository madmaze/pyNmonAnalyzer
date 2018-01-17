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

**Example Reports:**
---------------
- [Static Report](http://matthiaslee.com/scratch/pyNmonAnalyzer/data/report.html)
- [Interactive Report](http://matthiaslee.com/scratch/pyNmonAnalyzer/interactiveReport.html)


Goals:
-----
- make nmon log file analysis easier and faster
- create HTML based reports with embedded graphs
- create CSV files for more indepth data analysis
- interactive graphs for hands on analysis, perhaps using dygraph.js

Getting started:
-----
- [Python Nmon Analyzer: moving away from excel macros](http://matthiaslee.com/python-nmon-analyzer-moving-away-from-excel-macros/)

Installation:
-----
pyNmonAnalyzer is now available through pip and easy_install.   
If you have pip:   
```$> sudo pip install pyNmonAnalyzer```

If you'd like to mess with the source, please feel free to fork 
this github repo and contribute back changes.

Dependencies:  
-----
This tool depends on the python numpy package and the matplotlib package.
* If you are on a Debian/Ubuntu based system: `sudo apt-get install python-numpy python-matplotlib`  
* If you are on a RHEL/Fedora/Centos system: `sudo yum install python-numpy python-matplotlib`


Usage:
-----
```
usage: pyNmonAnalyzer [-h] [-x] [-d] [--force] [-i INPUT_FILE] [-o OUTDIR]
                      [-c] [-b] [-t REPORTTYPE] [-r CONFFNAME]
                      [--dygraphLocation DYGRAPHLOC] [--defaultConfig]
                      [-l LOGLEVEL]

nmonParser converts NMON monitor files into time-sorted CSV/Spreadsheets for
easier analysis, without the use of the MS Excel Macro. Also included is an
option to build an HTML report with graphs, which is configured through
report.config.

optional arguments:
  -h, --help            show this help message and exit
  -x, --overwrite       overwrite existing results (Default: False)
  -d, --debug           debug? (Default: False)
  --force               force using of config (Default: False)
  -i INPUT_FILE, --inputfile INPUT_FILE
                        Input NMON file
  -o OUTDIR, --output OUTDIR
                        Output dir for CSV (Default: ./report/)
  -c, --csv             CSV output? (Default: False)
  -b, --buildReport     report output? (Default: False)
  -t REPORTTYPE, --reportType REPORTTYPE
                        Should we be generating a "static" or "interactive"
                        report (Default: interactive)
  -r CONFFNAME, --reportConfig CONFFNAME
                        Report config file, if none exists: we will write the
                        default config file out (Default: ./report.config)
  --dygraphLocation DYGRAPHLOC
                        Specify local or remote location of dygraphs library.
                        This only applies to the interactive report. (Default:
                        http://dygraphs.com/dygraph-dev.js)
  --defaultConfig       Write out a default config file
  -l LOGLEVEL, --log LOGLEVEL
                        Logging verbosity, use DEBUG for more output and
                        showing graphs (Default: INFO)
```

Example Usage:
-------------
First generate a report config, most likely the default is all you need. This creates ./report.config
```$> pyNmonAnalyzer --defaultConfig```

Build HTML report with *interactive* graphs for test.nmon and store results to testReport  
```$> pyNmonAnalyzer -b -o testReport -i test.nmon```

Build HTML report with static graphs for test.nmon and store results to testReport  
```$> pyNmonAnalyzer -b -t static -o testReport -i test.nmon```

Compile CSV formatted tables for data in test.nmon and store results to testOut  
```$> pyNmonAnalyzer -c -o testOut -i test.nmon```

Configuration:
-------------
To control which items get graphed(CPU, MEM etc) you need to configure the report.config file. 
This is especially important for AIX NMON systems. To get a sense of what the config file 
should look like, run `pyNmonAnalyzer --defaultConfig` this will generate "report.config" in 
your local directory. It contians two examples, one for Linux and one for AIX systems. 
Adjust them according to your device names, for Linux you'll want to set DISKBUSY to your sda1 or sdb1 or what ever.
You should be able to use **any** nmon performance stats, so DISKBUSY, DISKREAD, CPU1, CPU2 etc.

Troubleshooting:
---------------
- **It crashes or my graphs don't show anything!**   
	Have you looked at your current report.config? Is it customized to your device names?
- **My interactive report will not display! What did I do wrong?**   
	Since the interactive report uses javascript to load CSV files, your browser needs to be allowed to read local files(if you are viewing locally). Firefox has been the most reliable for me, chrome currently does not allow JS to access local files.

**Any other problems, file an issue or send me an email.**


License:
-------
```
Copyright (c) 2012-2018 Matthias Lee, matthias.a.lee[]gmail.com
Last edited: January 16th 2018

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
