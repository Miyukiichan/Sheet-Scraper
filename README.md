# SheetScraper

## Overview

Work in progress at the moment but the goal of this is to scrape URLs from a list of the same website and then add them to a google sheet, filling in existing columns as required

## Dependencies

- lxml
- Pyside6
- gspread
- oauth2client
- validators
- bs4
- requests

## Todo

- Instead of scraping exceptions
  - Process the scraped data afterwards
    - Do the saving and print a log of the errors in a window
- Starting column/row
- Have gaps between the fields
- Table in options for field definitions
- Arrange UI
- Limit to being the same website as per the config
- Save data to a spreadsheet when press the go button
- Options file IO
- List file IO
- CLI options
- Testing mode that outputs results to the screen instead of updating the spreadsheet
