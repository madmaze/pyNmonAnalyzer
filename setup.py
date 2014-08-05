import os
from setuptools import setup

setup(
    name = "pyNmonAnalyzer",
    version = "1.0.1",
    author = "Matthias Lee",
    author_email="pynmonanalyzer@madmaze.net",
    maintainer = "Matthias Lee",
    maintainer_email = "pynmonanalyzer@madmaze.net",
    description = ("Python tool for reformatting and plotting/graphing NMON output"),
    license = "GPLv3",
    keywords = "python nmon analyzer pynmonanalyzer interactive static report visualization",
    url = "https://github.com/madmaze/pynmonanalyzer",
    packages=['pynmonanalyzer'],
    package_data = {'pynmonanalyzer': ['*.nmon','*.html']}
)
