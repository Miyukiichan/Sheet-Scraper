#!/usr/bin/python

from PySide6.QtWidgets import QMessageBox, QWidget, QDialog, QLabel, QLineEdit, QComboBox, QPushButton, QGridLayout, QApplication, QMainWindow, QMenuBar, QMenu, QVBoxLayout, QListWidget, QListWidgetItem
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
        Error("No spreadsheet name provided")
    except SpreadSheetNotFound:
        Error("Spreadsheet provided does not exist")
    except SheetNameNotFound:
        Error("Sheet name does not exist in provided spreadsheet")
    else:
        success = True
    return success


class Controller:
    def __init__(self):
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        jsonFile = 'client_secret.json'
        if not os.path.isfile(jsonFile):
            raise ClientSecretNotFound
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                jsonFile, scope)
        except:
            raise ClientSecretInvalid
        self.client = gspread.authorize(creds)
        self.spreadsheetName = ""
        self.sheetName = ""
        self.sheetNames = []
        self.sheet = 0
        self.spreadsheet = 0
        self.configName = ""
        self.configDomain = ""

    def setNames(self, ssName, sName):
        if ssName == "":
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
            sName = ""
            self.sheetName = ""
        if sName != self.sheetName or resetSheet:
            if sName == "":
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


class OptionsDialog(QDialog):
    def btnOkClick(self):
        if SetNames(self.controller, self.txtSpreadsheetName.text(), self.cbDefaultSheet.currentText()):
            self.destroy()

    def __init__(self, parent, controller):
        QDialog.__init__(self, parent)
        self.controller = controller
        # Controls
        self.lblSpreadsheetName = QLabel("Spreadsheet Name")
        self.lblDefaultSheet = QLabel("Default Sheet")
        self.txtSpreadsheetName = QLineEdit()
        self.txtSpreadsheetName.setText(controller.spreadsheetName)
        self.cbDefaultSheet = QComboBox()
        self.cbDefaultSheet.addItems(controller.sheetNames)
        index = self.cbDefaultSheet.findText(controller.sheetName)
        if index == -1:
            index = 0
        self.cbDefaultSheet.setCurrentIndex(index)
        self.btnOK = QPushButton()
        self.btnOK.setText("OK")
        self.btnOK.clicked.connect(self.btnOkClick)
        # Layout
        layout = QGridLayout()
        layout.addWidget(self.lblSpreadsheetName)
        layout.addWidget(self.txtSpreadsheetName)
        layout.addWidget(self.lblDefaultSheet)
        layout.addWidget(self.cbDefaultSheet)
        layout.addWidget(self.btnOK)

        self.setLayout(layout)


class Main(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        # Main root
        self.setWindowTitle("Sheet Scraper")
        self.resize(250, 250)
        # Menu Bar File
        fileMenu = QMenu('File', self)
        fileMenu.addAction("Save").triggered.connect(self.SaveFile)
        fileMenu.addAction("Save As").triggered.connect(self.SaveFileAs)
        fileMenu.addAction("Open").triggered.connect(self.OpenFile)
        fileMenu.addAction("Exit").triggered.connect(self.ExitApp)
        self.menuBar().addMenu(fileMenu)

        # Menu Bar Edit
        editMenu = QMenu('Edit', self)
        editMenu.addAction("Options").triggered.connect(self.OpenOptions)
        self.menuBar().addMenu(editMenu)

        # Controls
        self.txt = QLineEdit()
        self.listbox = QListWidget()

        addUrlButton = QPushButton()
        addUrlButton.setText("Add")
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
            Error("No file client_secret.json found in working directory")
            return
        except ClientSecretInvalid:
            Error("Client secret is invalid")
            return
        if not SetNames(self.controller, "Books", "Welsh History"):
            return
        self.controller.configName = "WHSmiths"
        self.controller.SetDomain("https://www.whsmith.co.uk/")
        self.show()

    def OpenOptions(self):
        OptionsDialog(self, self.controller).show()

    def AddUrl(self):
        text = self.txt.text()
        if text == "":
            return
        if (not self.controller.ValidURL(text)):
            if QMessageBox.question(self, "Use invalid URL", "The provided URL is not supported under the current configuration. Do you still wish to use it") != QMessageBox.StandardButton.Ok:
                return
        QListWidgetItem(text, self.listbox)

    def ExitApp(self):
        QApplication.quit()

    def ScrapeURLs(self):
        for i in range(0, self.listbox.count()):
            url = self.listbox.item(i).text()
            source = requests.get(url).text
            soup = BeautifulSoup(source, 'lxml')
            fieldList = ["URL",
                         ("TITLE", "h1", "h1 product-name"),
                         ("AUTHOR", "a", "author"),
                         ("PRICE", "span.span", "product-price-value.")]
            for field in fieldList:
                if type(field) != tuple and field.upper() == "URL":
                    print(field + " - " + url)
                    continue
                if len(field) < 2:
                    raise Exception(
                        "Not enough required entries for field " + field)
                fieldName = field[0]
                tag = field[1]
                className = ""
                if len(field) >= 3:
                    className = field[2]
                tags = []
                classes = []
                tags = tag.split(".")
                if className != "":
                    classes = className.split(".")
                    if len(tags) != len(classes):
                        raise Exception(
                            "Class/tag depth mismatch on field " + field)
                if (len(tags) == 0):
                    tags[0] = tag
                    if className != "":
                        classes[0] = className
                if "" in tags:
                    raise Exception("Cannot declare empty tags")
                current = 0
                counter = 0
                for tag in tags:
                    if current == 0:
                        current = soup
                    className = ""
                    if counter < len(classes):
                        className = classes[counter]
                    if className != "":
                        current = current.findChild(tag, {"class": className})
                    else:
                        current = current.findChild(tag)
                    counter += 1
                print(fieldName + " - " + current.text.strip())

    def OpenFile(self):
        return

    def SaveFile(self):
        return

    def SaveFileAs(self):
        return


app = QApplication(sys.argv)
mainWin = Main()

sys.exit(app.exec_())
