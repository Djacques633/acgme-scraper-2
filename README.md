# ACGME Scraper | NRMP Match

The ACGME Scraper and NRMP Match programs are intended to be used by the Ohio University Heritage College of Osteopathic to speed up the process of their residency matching.

## ACGME Scraper

The ACGME Scraper is a scraping program capable of gathering all medical program data from the US and compiling into one spread sheet.

The program uses Python 3.8 and the following libraries:

> tkinter
> BeautifulSoup/bs4
> xlsxwriter
> json
> os

To run the program, install Python 3.8 and the dependencies listed above. Run the program using `python acgme-scrape.py`.

You will then be prompted to input states to scrape from. The states are case insensitive, but see states.json if there is confusion.

After entering all states, hit enter and let the program go to work. It will gather program data from every medical program in the states that were specified and dump it into an excel sheet.

With all states, the program has the potential to scrape 15,000+ entries. The problem with this size of input is it will also take upwards of one to two weeks to complete..

## NRMP Match

The NRMP Match uses special HCOM Data in order to run that holds user's NRMP match data. Using the institution code, program name, and state, the program narrows down the ACGME program number and address with only 1 out of 200+ inaccuracies noticed.

In cases of discrepencies, the program utilizes inquirer to display the multiple options to the user and let them choose the correct match - or none if it does not exist.

The user will be prompted for data three times.

The first input given will be an output file. This **_will_** overwrite any data, so make sure it is an empty file.

The second input should be the excel sheet returned by the ACGME Scraper above. This is the file that will be used to match NRMP data to ACGME.

The third and final input file is the HCOM data file containing NRMP match data.

The following dependencies are needed for this program:

> xlrd
> xlsxwriter
> tkinter
> inquirer
