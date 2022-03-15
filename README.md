# ForEx Automation for Personal Finances

## Description

I find it useful for my personal finances to keep updated exchange rates on a spreadsheet
since my expenses are in more than one currency somewhat regularly.

This is an automation to save myself time, I usually would go to [this website](http://www.floatrates.com/)
to get historical exchange rates and fill in my spreadsheet manually every other day.

I finally got around to automating it, using some basic web scraping with the [requests](https://docs.python-requests.org/en/latest/)
library, as well as [openpyxl](https://openpyxl.readthedocs.io/en/stable/).

In case it is useful, an example XML is included [here](output.xml) and [here](output_error.xml).
output.xml is the regular download from the website, output_error.xml is one
where the website happened to not have any data available. The script does not account
for this however, it simply doesn't populate the corresponding cell.

## How to run

* Install requirements shown on [requirements.txt](requirements.txt) with pip in a virtual
environment if desired.
* Change code on line 47 to the path of the file you want to modify.
* Run as such:

```
python3 AutoFillForEx.py <day> <month>
```

* For example if we were to run the script for March 15th it would look like this:

```
python3 AutoFillForEx.py 15 3
```
