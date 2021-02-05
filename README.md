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

- Clear/edit list
- Shortcut keys
- Table in options for field definitions
- Handle rate limits
  - Minimise API calls
  - Cache remaining fields so not to re-read
- Options file IO
- List file IO
  - List file saving + save as
  - Drag and drop files
- Arrange UI
- CLI options
- Testing mode that outputs results to the screen instead of updating the spreadsheet
- CSV output
