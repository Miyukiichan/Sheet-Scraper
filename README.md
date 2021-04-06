# SheetScraper

## Overview

Work in progress at the moment but the goal of this is to scrape URLs from a list of the same website and then add them to a google sheet, filling in existing columns as required using data from the webpages

## Dependencies

- lxml
- Pyside6
- gspread
- oauth2client
- validators
- bs4
- requests

## Todo

- Convert the tuples to lists so I can add to them - useful for the loading procedure when the size of things are not known and I don't want empty strings all over the place
- Scraping speedups
  - Check tag validity once (not for every URL) and store it in dictionary
- Handle rate limits
  - Optional field indexing so find() is not needed
  - Cache remaining fields so not to re-read on error
- Options file IO
  - Options file load
- List file IO
  - List file saving + save as
  - Drag and drop file for faster loading
- Clear/edit list
- Shortcut keys
- Table in options for field definitions
- Arrange UI
- CLI options
- Testing mode that outputs results to the screen instead of updating the spreadsheet - sdtout mode essentially
- CSV output
