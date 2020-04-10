# -*- coding: utf-8 -*-
"""
Created on Sun Feb  2 23:39:55 2020

@author: Adeline
"""

# -*- coding: utf-8 -*-
"""
Spyder Editor

Written by Alec McKay, September, 2019

Script for combining tables in Excel workbooks.

How to Use:

Requirements:

"""

from os import listdir
from tkinter import filedialog
#from tkinter import *
from tkinter import Tk
from tkinter import messagebox
import pandas as pd
import xlwings as xw
import sys


#interface for opening files
def open_files(filetypes):
    
    #initiate UI
    root = Tk()
    
    root.filename = filedialog.askopenfilenames(initialdir="/", title="Select file",filetypes=filetypes)
    root.destroy()
    if root.filename == '':
        sys.exit(0)
    return root.filename

#interface for displaying message
def display_message(messageTitle, message):
    
    #initiate UI
    root = Tk()
    
    #hide root window
    root.withdraw()
    
    #create message box to display message
    messagebox.showinfo(messageTitle, message)
    root.destroy()

#appends files together
def append_books(filenames):
       
    #loop through all files in directory. If file is of .xlsx type, doesn't begin with '~' 
    #and isn't 'appendedBook.xlsx', then transform into a dataframe and append onto the 
    #empty dataframe that was created.
    appendBook = pd.DataFrame()
        
    directory = filenames[0][0: filenames[0].rfind('/')]
    
    
    for file in filenames:
        
        json_file = pd.read_json(file)
        
        
        titles = []
        descriptions = []
        urls = []
        
#        json1 = json_file.iloc[0]

        for res in json_file["organicResults"][0]:
            
            titles.append(res["title"])
            descriptions.append(res["description"])
            urls.append(res["url"])
            
        appendBook = appendBook.append(pd.DataFrame({"titles": titles, "descriptions": descriptions, "urls": urls}))

        
    try:
        #write dataframe into file
        appendBook.to_excel(directory + '\\appendedBook.xlsx')

    except PermissionError:
        display_message("Error", "Cannot have workbook called appendedBook.xlsx open while running macro. Please close and try again")
        sys.exit("appendedBook.xlsx was open")
        
    #open aggregated file
    appendedBook = xw.Book(directory + '\\appended_book.xlsx')
    
    #append file names onto string to print
    printFiles = ''
    
    for file in filenames:
        printFiles = printFiles + file[file.rfind('/') + 1:] + ', ' 
    
    printFiles = printFiles[:len(printFiles)-2]
    
    #display completion message
    display_message("Macro Completed!","The following workbooks were aggregated: \n\n" + printFiles + "\n\n" + "in " + appendedBook.fullname)
   
    #method to run script
def main():
    append_books(open_files([("JSON files","*.json")]))
    
if __name__ == '__main__':
    main()
