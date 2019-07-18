QUICK START GUIDE

DEPENDENCIES:
	Python 3.x
	pip package manager

pip packages (or installation via any package manager):
	argparse
	pypyodbc

SETUP:

1. Verify python is installed
	in cmd, type "python" for version details
	if not installed, download from 'https://www.python.org/downloads/windows/'

2. Install pip
	save text from 'https://bootstrap.pypa.io/get-pip.py' as get-pip.py
	in cmd, navigate directory with "cd" command to folder containing get-pip.py
	type, "python get-pip.py"
	verify installation with "pip --version"
	optional: upgrade to latest version with "python -m pip install --upgrade pip"

3. Install pip packages
	in cmd, type:
	"pip install argparse"
	"pip install pypyodbc"

USAGE:

(ISSUE - NEEDS FIXING): Line 15 needs to be manually updated with the correct filepath for the Access database. Use double backslashes so Python doesn't register one backslash as an escape character.

1. Create .csv and populate with MANPARTNUM and corresponding INT_BIONUMs. Row A contains headers. In my case, A1:A2 says MANPARTNUM:INT_BIONUM. Column A contains MANPARTNUMs and column B contains INT_BIONUMS.

2. In cmd, navigate to current working directory with "cd" command.

3. Execute python script with arguments:
	python UpdateLib.py --file [FILENAME].csv --commit 
