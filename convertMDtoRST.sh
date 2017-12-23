# in order to please pypi, we use this to convert the README.md into .rst
pandoc --columns=100 --output=README.rst --to rst README.md
