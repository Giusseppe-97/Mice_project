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


df = pd.read_excel("C:/Users/josep/Documents/GitHub/Mice_proyect/data/Mice_table.xlsx")
n_df = df            
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
q = []
for element in total_data:
    age_in_weeks = 0
    d1 = []
    cont = []
    for index, rows in total_data[j].iterrows():
        age_in_weeks = math.floor(n_df['Age_(days)'][index]/7)
        d1.append(age_in_weeks)
        # print(d1)
        if age_in_weeks not in total_data[j]: 
            cont.append(age_in_weeks)
    j = j + 1    
    d1.sort(reverse=True)
    d2.append(d1)
    q. append(cont)
print(d2)
print(cont)
print(q)
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
plt.show()









































"""
def import_excel_file_2(self):
""plot simple plot

Returns:Four figures sharing the same x axis
""
df = pd.read_excel(filepath)
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
q = []
for element in total_data:
    age_in_weeks = 0
    d1 = []
    cont = []
    for index, rows in total_data[j].iterrows():
        age_in_weeks = math.floor(n_df['Age_(days)'][index]/7)
        d1.append(age_in_weeks)
        # print(d1)
        if age_in_weeks not in total_data[j]: 
            cont.append(age_in_weeks)
    j = j + 1    
    d1.sort(reverse=True)
    d2.append(d1)
    q. append(cont)
print(d2)
print(cont)
print(q)
d3 = [d2[0][0], d2[1][0], d2[2][0], d2[3][0]]

# Colors
colors = sns.color_palette("mako", 4)

#Plot
fig, axs = plt.subplots(2, 2, figsize=(8, 8), sharey=True,  tight_layout=True)
# axs.locator_params(axis='x', integer=True)
b = int(max(d3))+1
print(b)
hist1, bins = np.histogram(d2[0],bins = b)
hist2, bins = np.histogram(d2[1],bins = b)
hist3, bins = np.histogram(d2[2],bins = b)
hist4, bins = np.histogram(d2[3],bins = b)

# I need to garantee its the same x axis quantity and integers
axs[0,0].bar(np.arange(0,b,1), hist1, color=colors[0], label='Male Wildtype', edgecolor='black')
axs[0,1].bar(np.arange(0,b,1), hist2, color=colors[1], label='Female Wildtype', edgecolor='black')
axs[1,0].bar(np.arange(0,b,1), hist3, color=colors[2], label='Male Heterozygous', edgecolor='black')
axs[1,1].bar(np.arange(0,b,1), hist4, color=colors[3], label='Female Heterozygous', edgecolor='black')

axs[0,0].xaxis.set_major_locator(MaxNLocator(integer=True))
axs[0,1].xaxis.set_major_locator(MaxNLocator(integer=True))
axs[1,0].xaxis.set_major_locator(MaxNLocator(integer=True))
axs[1,1].xaxis.set_major_locator(MaxNLocator(integer=True))

#plt.subplots_adjust(hspace=0)
axs[1,0].set_xlabel(r'Time (weeks)', fontsize=15)
axs[1,1].set_xlabel(r'Time (weeks)', fontsize=15)
axs[0,0].set_ylabel(r'Number of mice', fontsize =20, horizontalalignment='right')

plt.sca(axs[0,0])
plt.title('Male', fontsize=15)
plt.sca(axs[0,1])
plt.title('Female', fontsize=15)

axs[0,0].legend()
axs[0,1].legend()
axs[1,0].legend()
axs[1,1].legend()

plt.suptitle('MICE QUANTITY VS AGE', horizontalalignment='center', fontweight ="bold", fontsize=20)
"""
