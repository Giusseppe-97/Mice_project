"""
Created on Fri Apr  2 15:18:14 2021

@author: josep
"""

# Executing the program as a HD window for windows and exception for running it on mac
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

import tkinter as tk
from tkinter import StringVar, ttk
import tkcalendar as tkc
from tkinter import messagebox

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk)

# from Plotting_mice import MyPlot
import matplotlib.pyplot as plt
# import xlsxwriter as xw
from matplotlib.ticker import MaxNLocator
import pandas as pd
import numpy as np
import math
import seaborn as sns

import datetime as dt


class OpenSave:
    """[Creating a main App class where all the frames are going to be set upon]

    The main App class must inherit from tk.Tk which is the root or main window
    """

    def __init__(self):
        """init is the method for setting default state of the object
        """
        self.open_excel_file_location()
        self.save_results()

    def open_excel_file_location(self):
        """Open the File Explorer to select desired excel file
        """
        global filepath1

        filepath1 = askopenfilename(filetypes=[(
            "xlsx Files", "*.xlsx"), ("csv Files", "*.csv"), ("All Files", "*.*")])

        if not filepath1:
            return
        with open(filepath1, "r"):
            self.textbox1.insert(tk.END, filepath1)

    def save_results(self):
        """Open the file Explorer to select desired location to save results
        """
        global filepath2
        filepath2 = askdirectory()
        if not filepath2:
            return
        self.textbox2.insert(tk.END, filepath2)