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
from matplotlib.ticker import MaxNLocator
import pandas as pd
import numpy as np
import math
import seaborn as sns

import datetime as dt


df = pd.read_excel('C:/Users/josep\Documents/GitHub\Mice_proyect/data/Mice_table.xlsx')
i=0
init_date = self.textbox3.get()
init_date[::-1]
final_date = self.textbox4.get()
final_date[::-1]

print(init_date)
print(final_date)

n_df = df[(init_date <= df.Date_of_birth) & (df.Date_of_birth <= final_date)]
print(df['Date_of_birth'])
print("")
print(n_df['Date_of_birth']) 
print("")
                
l = len(n_df['Date_of_birth'])

df_MWt = pd.DataFrame(n_df.loc[(df['Sex']=='Male') & (n_df['Genotype']=='Wildtype') & (n_df['Status']=='Alive')])
df_FWt = pd.DataFrame(n_df.loc[(df['Sex']=='Female') & (n_df['Genotype']=='Wildtype') & (n_df['Status']=='Alive')])
df_MHet = pd.DataFrame(n_df.loc[(df['Sex']=='Male') & (n_df['Genotype']=='Heterozygous') & (n_df['Status']=='Alive')])
df_FHet = pd.DataFrame(n_df.loc[(df['Sex']=='Female') & (n_df['Genotype']=='Heterozygous') & (n_df['Status']=='Alive')])
print(len(df_MWt))
print(len(df_FWt))
print(len(df_MHet))
print(len(df_FHet))
total_data = [df_MWt, df_FWt, df_MHet, df_FHet]

j=0 
d2 = []
for element in total_data:
    age_in_weeks = 0
    d1 = []
    for index, rows in total_data[j].iterrows():
        age_in_weeks = math.floor(n_df['Age_(days)'][index]/7)
        d1.append(age_in_weeks)
    j = j + 1    
    d1.sort(reverse=True)
    d2.append(d1)
print(d2)
d3 = [d2[0][0], d2[1][0], d2[2][0], d2[3][0]]
d4 = [d2[0], d2[1], d2[2], d2[3]]

b = int(max(d3))+1
i=0
listy = []
for i in range(b-1):
    
    listy.append(i)
    i =i+1


# Colors
colors = sns.color_palette("mako", 4)

#Plot
fig = plt.hist(d2, bins=listy, color=colors)
# axs.locator_params(axis='x', integer=True)
c = [0,1,2,3]