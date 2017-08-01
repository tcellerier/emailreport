# Email report 

## Description
* Version 1.0 - 08/2017
* Framework to automatize the generation of email reports with:
  * text or html email body
  * local or generated image file included
  * CSS HTML tables predesigned
  * Attached CSV / Excel files with data extracted from databases

## Usage
    usage: emailreport.py [-h] [-e <email>] [-p1 <param1>] [-p2 <param2>] [-p3 <param3>] script_file

    positional arguments:
       script_file      Specify the script file to execute (inside scripts/ directory, without .py)

    optional arguments:
      -h, --help        show this help message and exit
      -e <email>, --email <email>
                        Override email recipient
      -p1 <param1>, --param1 <param1>
                        Parameter #1 to send to script
      -p2 <param2>, --param2 <param2>
                        Parameter #2 to send to script
      -p3 <param3>, --param3 <param3>
                        Parameter #3 to send to script`
    Example:
      ./emailreport.py ZZtemplate_text

## Available template scripts
* scripts/ZZtemplate_text.py
* scripts/ZZtemplate_html.py
* scripts/ZZtemplate_querymail.py

## Dependancies to install
* pip install --user numpy
* pip install --user pandas
* pip install --user openpyxl
* pip install --user xlsxwriter
* sudo apt-get install python-matplotlib
