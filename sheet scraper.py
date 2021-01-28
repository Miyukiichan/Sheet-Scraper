#!/usr/bin/python

from PySide6.QtWidgets import QMessageBox, QWidget, QDialog, QLabel, QLineEdit, QComboBox, QPushButton, QGridLayout, QApplication, QMainWindow, QMenuBar, QMenu, QVBoxLayout, QListWidget, QListWidgetItem, QFileDialog
from PySide6.QtCore import Slot, Qt
import validators
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
from bs4 import BeautifulSoup
import os.path
import sys


class ControllerException(Exception):
    pass


class NoSpreadSheetName(ControllerException):
    pass


class SpreadSheetNotFound(ControllerException):
    pass


class SheetNameNotFound(ControllerException):
    pass


class ClientSecretNotFound(ControllerException):
    pass


class ClientSecretInvalid(ControllerException):
    pass


def Error(message):
    QMessageBox.critical(None, 'Error', message)


def SetNames(controller, ssName, sName):
    success = False
    try:
        controller.setNames(ssName, sName)
    except NoSpreadSheetName:
        Error('No spreadsheet name provided')
    except SpreadSheetNotFound:
        Error('Spreadsheet provided does not exist')
    except SheetNameNotFound:
        Error('Sheet name does not exist in provided spreadsheet')
    else:
        success = True
    return success


class Controller:
    def __init__(self):
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/spreadsheets',
                 'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/drive']
        jsonFile = 'client_secret.json'
        if not os.path.isfile(jsonFile):
            raise ClientSecretNotFound
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                jsonFile, scope)
        except:
            raise ClientSecretInvalid
        self.client = gspread.authorize(creds)
        self.spreadsheetName = ''
        self.sheetName = ''
        self.sheetNames = []
        self.sheet = 0
        self.spreadsheet = 0
        self.configName = ''
        self.configDomain = ''

    def setNames(self, ssName, sName):
        if ssName == '':
            raise NoSpreadSheetName
        resetSheet = False
        if ssName != self.spreadsheetName:
            try:
                self.spreadsheet = self.client.open(ssName)
            except gspread.SpreadsheetNotFound:
                raise SpreadSheetNotFound
            self.sheetNames = self.spreadsheet.worksheets()
            for i in range(0, len(self.sheetNames)):
                self.sheetNames[i] = self.sheetNames[i].title
            self.spreadsheetName = ssName
            resetSheet = True
            sName = ''
            self.sheetName = ''
        if sName != self.sheetName or resetSheet:
            if sName == '':
                self.sheetName = self.spreadsheet.sheet1.title
            else:
                if self.sheetNames.count(sName) == 0:
                    raise SheetNameNotFound
                self.sheetName = sName
            self.sheet = self.spreadsheet.worksheet(self.sheetName)
        return 0

    def SetDomain(self, url):
        last = len(url) - 1
        if (url[last] == '/'):
            url = url[0:last]
        self.configDomain = url

    def ValidURL(self, url):
        return validators.url(url) and self.configDomain in url

    def ScrapeURLs(self, urls):
        fieldList = ['URL',
                     ('TITLE', 'h1', 'h1 product-name'),
                     ('AUTHOR', 'a', 'author'),
                     ('PRICE', 'span', 'product-price-value')]
        errors = []
        self.processed = []
        for url in urls:
            if not validators.url(url):
                errors.append((url, 'Invalid URL'))
                continue
            try:
                source = requests.get(url).text
                soup = BeautifulSoup(source, 'lxml')
            except:
                errors.append((url, 'Webpage does not exist or is empty'))
                continue
            for field in fieldList:
                if type(field) != tuple and field.upper() == 'URL':
                    self.processed.append((url, field))
                    continue
                if len(field) < 2:
                    errors.append(
                        (url, field, 'Not enough required entries for field'))
                    continue
                fieldName = field[0]
                tag = field[1]
                className = ''
                if len(field) >= 3:
                    className = field[2]
                tags = []
                classes = []
                tags = tag.split('.')
                if className != '':
                    classes = className.split('.')
                    if len(tags) != len(classes):
                        errors.append((url, fieldName,
                                       'Class/tag depth mismatch on field'))
                        continue
                if (len(tags) == 0):
                    tags[0] = tag
                    if className != '':
                        classes[0] = className
                if '' in tags:
                    errors.append(
                        (url, fieldName, 'Cannot declare empty tags for field'))
                    continue
                current = 0
                counter = 0
                for tag in tags:
                    if current == 0:
                        current = soup
                    className = ''
                    if counter < len(classes):
                        className = classes[counter]
                    if className != '':
                        current = current.findChild(tag, {'class': className})
                    else:
                        current = current.findChild(tag)
                    counter += 1
                if current == None:
                    errors.append(
                        (url, fieldName, 'Problem finding tags in webpage'))
                    continue
                self.processed.append((fieldName, current.text.strip()))
        return errors


class OptionsDialog(QDialog):
    def btnOkClick(self):
        if SetNames(self.controller, self.txtSpreadsheetName.text(), self.cbDefaultSheet.currentText()):
            self.accept()

    def __init__(self, controller):
        QDialog.__init__(self)
        self.controller = controller
        # Controls
        self.lblSpreadsheetName = QLabel('Spreadsheet Name')
        self.lblDefaultSheet = QLabel('Default Sheet')
        self.txtSpreadsheetName = QLineEdit()
        self.txtSpreadsheetName.setText(controller.spreadsheetName)
        self.cbDefaultSheet = QComboBox()
        self.cbDefaultSheet.addItems(controller.sheetNames)
        index = self.cbDefaultSheet.findText(controller.sheetName)
        if index == -1:
            index = 0
        self.cbDefaultSheet.setCurrentIndex(index)
        self.btnOK = QPushButton()
        self.btnOK.setText('OK')
        self.btnOK.clicked.connect(self.btnOkClick)
        # Layout
        layout = QGridLayout()
        layout.addWidget(self.lblSpreadsheetName)
        layout.addWidget(self.txtSpreadsheetName)
        layout.addWidget(self.lblDefaultSheet)
        layout.addWidget(self.cbDefaultSheet)
        layout.addWidget(self.btnOK)

        self.setLayout(layout)
        self.exec()
        print("Here")


class ErrorLog(QDialog):
    def __init__(self, errors):
        QDialog.__init__(self)
        self.setWindowTitle('Error, Log')
        list = QListWidget(self)
        for error in errors:
            QListWidgetItem(' - '.join(error), list)
        layout = QVBoxLayout(self)
        layout.addWidget(list)
        self.setLayout(layout)
        self.exec()


class Main(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        # Main root
        self.setWindowTitle('Sheet Scraper')
        self.resize(250, 250)
        # Menu Bar File
        fileMenu = QMenu('File', self)
        fileMenu.addAction('Save').triggered.connect(self.SaveFile)
        fileMenu.addAction('Save As').triggered.connect(self.SaveFileAs)
        fileMenu.addAction('Open').triggered.connect(self.OpenFile)
        fileMenu.addAction('Exit').triggered.connect(self.ExitApp)
        self.menuBar().addMenu(fileMenu)

        # Menu Bar Edit
        editMenu = QMenu('Edit', self)
        editMenu.addAction('Options').triggered.connect(self.OpenOptions)
        self.menuBar().addMenu(editMenu)

        # Controls
        self.txt = QLineEdit()
        self.txt.returnPressed.connect(self.AddUrl)
        self.listbox = QListWidget()

        addUrlButton = QPushButton()
        addUrlButton.setText('Add')
        addUrlButton.clicked.connect(self.AddUrl)

        goButton = QPushButton()
        goButton.setText('Process URLs')
        goButton.clicked.connect(self.ScrapeURLs)

        # Layout

        layout = QVBoxLayout()
        layout.addWidget(self.txt)
        layout.addWidget(addUrlButton)
        layout.addWidget(self.listbox)
        layout.addWidget(goButton)

        widget = QWidget()
        widget.setLayout(layout)

        self.setCentralWidget(widget)

        self.txt.setFocus()
        try:
            self.controller = Controller()
        except ClientSecretNotFound:
            Error('No file client_secret.json found in working directory')
            return
        except ClientSecretInvalid:
            Error('Client secret is invalid')
            return
        if not SetNames(self.controller, 'Books', 'Welsh History'):
            return
        self.controller.configName = 'WHSmiths'
        self.controller.SetDomain('https://www.whsmith.co.uk/')
        self.show()

    def OpenOptions(self):
        OptionsDialog(self.controller)

    def AddUrl(self):
        text = self.txt.text()
        if text == '':
            return
        if (not self.controller.ValidURL(text)):
            if QMessageBox.question(self, 'Use invalid URL', 'The provided URL is not supported under the current configuration. Do you still wish to use it') != QMessageBox.StandardButton.Yes:
                return
        QListWidgetItem(text, self.listbox)

    def ExitApp(self):
        QApplication.quit()

    def ScrapeURLs(self):
        urls = []
        for i in range(0, self.listbox.count()):
            urls.append(self.listbox.item(i).text())
        errors = self.controller.ScrapeURLs(urls)
        for line in self.controller.processed:
            print(line)
        if len(errors) == 0:
            QMessageBox.information(
                self, 'Success', 'Successfully processed all URLs with no errors')
        else:
            ErrorLog(errors)

    def OpenFile(self):
        fileName = QFileDialog.getOpenFileName(self, 'Import URL list')[0]
        if not os.path.isfile(fileName):
            return
        self.listbox.clear()
        reader = open(fileName, 'r')
        lines = []
        for line in reader:
            if not self.controller.ValidURL(line):
                continue
            lines.append(line.strip())
        reader.close()
        if len(lines) == 0:
            Error('No valid URLs found in file')
            return
        for line in lines:
            QListWidgetItem(line, self.listbox)

    def SaveFile(self):
        return

    def SaveFileAs(self):
        return


app = QApplication(sys.argv)
mainWin = Main()

sys.exit(app.exec_())
