Project for downloading eBooks from www.gutenberg.org and generating PDFs for them.

To run the script you'll need Python 3 programming language installed on your PC.
Also to install Python's runtime dependecies you'll need python3-pip package.

First, you will need to install all the libraries, that project depends on:
    "pip3 install -r requirements.txt"


python guttenberg2.py -h
usage: python3 guttenberg2.py [options]

Project Guttenberg books scrape script:

usage: python3 guttenberg2.py [options]

Project Guttenberg books scrape script:

options:
  -h, --help            show this help message and exit
  --indexes INDEXES     books indexes to process, comma separated
  -s START, --start START
                        start index of the program
  -e END, --end END     end index of the program
  --word                generate Word documents
  --cover               generate PDF covers

Script will create output folder named as datestamp, and also maintain last processed book index and Excel file with each run spreadsheet

Script usage notes:
It will store last processed book index in the file called "index",
you may manually insert number till which script will download books.

By default, script will produce PDF interior and cover along with Word documents.

To start books scraping script, please run: "python3 guttenberg2.py"

You can also specify start index and end index with: "python3 guttenberg2.py <START_NUM> <END_NUM>"
If not specified, last published index will be fetched from website and will be used as end index, 
and last processing stored index will be used as start index.
