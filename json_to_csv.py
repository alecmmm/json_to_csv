# # -*- coding: utf-8 -*-
# """
# Created on Sun Feb  2 23:39:55 2020
#
# @author: Adeline
# """
#
# # -*- coding: utf-8 -*-
# """
# Spyder Editor
#
# Written by Alec McKay, March 2020
#
# Script for combining tables in Excel workbooks.
#
# How to Use:
#
# Requirements:
#
# """

from os import listdir
from tkinter import filedialog
from tkinter import Tk
from tkinter import messagebox
import pandas as pd
# import xlwings as xw
import sys


# interface for opening files
def open_files(filetypes):
    # initiate UI
    root = Tk()

    root.filename = filedialog.askopenfilenames(initialdir="/", title="Select file", filetypes=filetypes)
    root.destroy()
    if root.filename == '':
        sys.exit(0)
    return root.filename


# interface for displaying message
def display_message(message_title, message):
    # initiate UI
    root = Tk()

    # hide root window
    root.withdraw()

    # create message box to display message
    messagebox.showinfo(message_title, message)
    root.destroy()


# appends files together
def append_books(file_names):
    # loop through all files in directory. If file is of .xlsx type, doesn't begin with '~'
    # and isn't 'appendedBook.xlsx', then transform into a dataframe and append onto the
    # empty dataframe that was created.
    append_book = pd.DataFrame()

    directory = file_names[0][0: file_names[0].rfind('/')]

    for file in file_names:

        json_file = pd.read_json(file)

        file_names = []
        titles = []
        descriptions = []
        urls = []
        search_terms =[]

        #        json1 = json_file.iloc[0]
        for a_query in json_file["searchQuery"]:
            for res in json_file["organicResults"][0]:
                file_names.append(file[file.rfind('/') + 1:])
                titles.append(res["title"])
                descriptions.append(res["description"])
                urls.append(res["url"])
                search_terms.append(a_query["term"])

        append_book = append_book.append(pd.DataFrame({"titles": titles, "descriptions": descriptions, "urls": urls,
                                                       "search_terms": search_terms, "file_names": file_names}))

    try:
        # write dataframe into file
        append_book.to_csv(directory + '/appended_json_files.csv', index=False)

    except PermissionError:
        display_message("Error",
                        "Cannot have workbook called appendedBook.xlsx open while running macro. Please close and try "
                        "again")
        sys.exit("appended_book.csv was open")

    # append file names onto string to print
    print_files = ''

    for file in file_names:
        print_files = print_files + file[file.rfind('/') + 1:] + ', '

    print_files = print_files[:len(print_files) - 2]

    # display completion message
    display_message("Macro Completed!",
                    "The following workbooks were aggregated: \n\n" + print_files + "\n\n" + "in appended_book.csv")

    # method to run script


def main():
    append_books(open_files([("JSON files", "*.json")]))


if __name__ == '__main__':
    main()
