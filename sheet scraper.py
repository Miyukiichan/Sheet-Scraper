#!/usr/bin/python

import tkinter as tk
from tkinter import messagebox
import tkinter.ttk as ttk
import validators
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
from bs4 import BeautifulSoup
import os.path


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
    tk.messagebox.showerror("Error", message)


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


class OptionsDialog(tk.Toplevel):
    def btnOkClick(self):
        if SetNames(self.controller, self.txtSpreadsheetName.get(), self.cbDefaultSheet.get()):
            self.destroy()

    def __init__(self, parent, controller):
        tk.Toplevel.__init__(self, parent)
        self.controller = controller
        # Controls
        self.lblSpreadsheetName = tk.Label(self, text="Spreadsheet Name")
        self.lblDefaultSheet = tk.Label(self, text="Spreadsheet Name")
        self.txtSpreadsheetName = tk.Entry(self, width=30)
        self.txtSpreadsheetName.insert(0, controller.spreadsheetName)
        self.cbDefaultSheet = ttk.Combobox(
            self, width=27, values=controller.sheetNames)
        self.cbDefaultSheet.current(
            controller.sheetNames.index(controller.sheetName))
        self.btnOK = tk.Button(self, text="OK", command=self.btnOkClick)

        # Layout
        self.lblSpreadsheetName.grid(column=0, row=0)
        self.lblDefaultSheet.grid(column=0, row=1)
        self.txtSpreadsheetName.grid(column=1, row=0)
        self.cbDefaultSheet.grid(column=1, row=1)
        self.btnOK.grid(column=1, row=2)

    def show(self):
        self.focus_set()
        self.grab_set()
        self.wait_window()


class Main(tk.Frame):
    def __init__(self):
        # Main root
        self.root = tk.Tk()
        self.root.title("Sheet Scraper")
        self.root.geometry('250x250')

        # Menu Bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        # Menu Bar File
        fileMenu = tk.Menu(menubar, tearoff=0)
        fileMenu.add_command(label="Save", command=self.SaveFile)
        fileMenu.add_command(label="Save As", command=self.SaveFileAs)
        fileMenu.add_command(label="Open", command=self.OpenFile)
        fileMenu.add_command(label="Exit", command=self.ExitApp)
        menubar.add_cascade(label="File", menu=fileMenu)
        # Menu Bar Edit
        editMenu = tk.Menu(menubar, tearoff=0)
        editMenu.add_command(label="Options", command=self.OpenOptions)
        menubar.add_cascade(label="Edit", menu=editMenu)

        self.entries = tk.Frame(self.root)

        # Controls
        self.txt = tk.Entry(self.entries, width=20)
        addUrlButton = tk.Button(self.entries, text="Add", command=self.AddUrl)
        self.listbox = tk.Listbox(self.root)
        goButton = tk.Button(self.root, text="Process urls",
                             command=self.ScrapeURLs)

        # Layout
        self.txt.grid(column=0, row=0)
        addUrlButton.grid(column=1, row=0)
        self.entries.pack()
        self.listbox.pack()
        goButton.pack()

        self.txt.focus()
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
        self.root.mainloop()

    def OpenOptions(self):
        OptionsDialog(self.root, self.controller).show()

    def AddUrl(self):
        text = self.txt.get()
        if text == "":
            return
        if (not self.controller.ValidURL(text)):
            if not tk.messagebox.askyesno("Use invalid URL", "The provided URL is not supported under the current configuration. Do you still wish to use it"):
                return
        self.listbox.insert(self.listbox.size(), text)

    def ExitApp(self):
        self.root.quit()

    def ScrapeURLs(self):
        for i in range(0, self.listbox.size()):
            if i == 1:
                break
            url = self.listbox.get(i)
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


mainWin = Main()
