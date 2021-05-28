
i = 1
excel_sheet_vector_week = []
for i in range(53):
    excel_sheet_vector_week.append("week "+ str(i))

print(excel_sheet_vector_week)


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
import xlsxwriter as xw
from matplotlib.ticker import MaxNLocator

import pandas as pd
import numpy as np
import math
import seaborn as sns
import datetime as dt

from mice_gui import Application as app

class FigureManager:
    """
    Enables the creation and management of folder structures for multipurpose 
    projects with datetime configurations.
    :param project_name: name of the project for the FolderManager agent.
    """
    def __init__(self):
        self.current_datetime = dt.datetime.now()
        self.generate_figures()

    def generate_figures(self):
        # plot1 = mg.plot_individual_hist().hist_plot
        plot2 = app.plot_4_hist().hist_4_plot
        # fp1 = mg.filepath_hist_plot
        fp2 = app.filepath_4_plot
        return fp2 
        
