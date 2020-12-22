# document-link-checker

This script will parse a set of files and extract hyperlinks, and check each link to see if it is reachable. Error URLs and details will be output into a csv file. This guide assumes limited experience with python and git, but some experience using a unix-like terminal emulator. Only `docx` files are currently supported.

## Installation
:warning: **Warnings!**
* writen using **python 3.8.5** no guarentees on compatiiblity with older versions
* setup instructions assume unix-like terminal emulator, [subsystem](https://www.howtogeek.com/249966/how-to-install-and-use-the-linux-bash-shell-on-windows-10/) or [alternate setup](https://docs.microsoft.com/en-us/windows/python/beginners) necessary to run on Windows

1. [Install python](https://www.python.org/downloads/) >= version 3.8.5
1. Verify python version
	```
	python --version
	```
1. [Install git](https://git-scm.com/downloads)
1. Verify git installation
	```
    git --version
    ```
1. Checkout code
	```
	git clone git@github.com:pdumoulin/document-link-checker.git
    cd document-link-checker
	```

3. Create virtual environment to install python packages
	```
	python -m venv env
	```
4. Install python packages
	```
	source env/bin/activate
	pip install -r requirements.txt 
	```
:warning: The message `ERROR: Failed building wheel for python-docx` may appear in frightening red text. You can ignore it.

## Running
1. Start virtual environment (unless you just did install steps)
	```
    cd document-link-checker
    source env/bin/activate
    ```
    
1. Run script
	```
    python run.py
    ```
    
1. Open output file

## Help
See all command line options and their defaults
```
python run.py --help
```
