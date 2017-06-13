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

from setuptools import setup
try:
    import pypandoc
    long_description = pypandoc.convert('README.md', 'rst')
except ImportError:
    long_description = open('README.md').read()
setup(
    name = "pyNmonAnalyzer",
    version = "1.0.6",
    author = "Matthias Lee",
    author_email="pynmonanalyzer@madmaze.net",
    maintainer = "Matthias Lee",
    maintainer_email = "pynmonanalyzer@madmaze.net",
    description = ("Python tool for reformatting and plotting/graphing NMON output"),
    long_description = long_description,
    license = "GPLv3",
    keywords = "python nmon analyzer pynmonanalyzer interactive static report visualization",
    url = "https://github.com/madmaze/pynmonanalyzer",
    packages=['pynmonanalyzer'],
    package_dir={'pynmonanalyzer': 'src'},
    package_data = {'pynmonanalyzer': ['test.nmon','interactiveReport.tpl.html']},
    entry_points={
        "console_scripts": [
            "pyNmonAnalyzer=pynmonanalyzer:main",
        ]}
)
