nmon2csv
========

A tool for parsing and reshuffeling nmon's output into "normal" csv format.
Currently nmon puts out a long file with a header at the beginning, and continuously
appended by timestamps. This makes it hard to bring the raw data into an easily 
machine readable format.

Goals:
-----
- successfully parse NMON output (modularize it)
- (ability to) produce CSV files for every(or all-in-one) captured resource
- (ability to) produce charts with matplotlib or gnuplot (or both)


License:
--------

Copyright (c) 2012 Matthias Lee, matthias.a.lee[]gmail.com <br />
Last edited: Sept 25th 2012<br />

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
