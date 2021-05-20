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

# from PIL import ImageTk, Image


class Application(tk.Tk):
    """[Creating a main App class where all the frames are going to be set upon]

    The main App class must inherit from tk.Tk which is the root or main window
    """

    def __init__(self, *args, **kwargs):
        """init is the method for setting default state of the object
        """
        super().__init__(*args, **kwargs)

        # Set style of the GUI
        self.tk.call('source', r'../docs/style/azure.tcl')
        ttk.Style().theme_use('azure')

        # Create Main Frames
        self.mainFrame1 = tk.Frame(self)
        self.mainFrame2 = tk.Frame(self)
        
        # Create Calendar 
        self.cal = tkc.Calendar(
            self.mainFrame1, selectbackground="#120597", background = "#120597",
            selectmode="day", year=2021, month=5, day=1
            )

        # Call methods
        self.configure_basic_tk_properties()
        self.pack_all()

        # # Create scrollbar
        # self.scroll_frame_1 = ttk.Scrollbar(self.mainFrame1, orient="vertical", command=self.canvas1.yview)
        # self.scroll_frame_1.pack(side='right', fill='y')

    def configure_basic_tk_properties(self):
        """This method configures the basic tkinter esthetic properties for the GUI
        """
        self.title("Mice GUI")

        # Setting the main App in the center regardless to the window's size chosen by the user
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        # Setting a background for the main App (which is going to hold both frames)
        self.configure(bg="light blue")

        # Create and place Lables for each Frame
        self.lable = tk.Label(self.mainFrame1, text="RUN DATA", foreground="white",
                              background="#120597").place(x=0, width=1920)
        self.lable2 = tk.Label(self.mainFrame2, text="HISTOGRAM DISPLAY PREVIEW", foreground="white",
                               background="#120597").place(y=0, width=1920)

        # Create input and output lables
        self.label3 = ttk.Label(
            self.mainFrame1, text="Input: ", background="white")
        self.label4 = ttk.Label(
            self.mainFrame1, text="Output:", background="white")
        self.label5 = tk.Label(self.mainFrame1, text="Select Date Interval: ", font=(
            "Arial bold", 11), background="white")
        self.label6 = ttk.Label(
            self.mainFrame1, text="From:", background="white")
        self.label7 = ttk.Label(
            self.mainFrame1, text="To:", background="white")

        # Create textboxes
        self.textbox1 = ttk.Entry(self.mainFrame1, width=80)
        self.textbox2 = ttk.Entry(self.mainFrame1, width=80)
        self.textbox3 = ttk.Entry(self.mainFrame1, width=20)
        self.textbox4 = ttk.Entry(self.mainFrame1, width=20)

        # Create and inicialize buttons
        self.button1 = ttk.Button(self.mainFrame1, text="Select", command=lambda: [
                                  self.open_excel_file_location()])
        self.button2 = ttk.Button(
            self.mainFrame1, text="Save ", command=lambda: [self.save_results()])
        self.button3 = ttk.Checkbutton(
            self.mainFrame1, text="RUN", style='ToggleButton', command=self.display_plot)
        self.button5 = ttk.Button(
            self, text="Start date", command=self.grab_start_date)
        self.button4 = ttk.Button(
            self, text="End date", command=self.grab_end_date)
        self.button_quit = ttk.Button(
            master=self, text="Quit", command=self.quit)

        # Create Canvas (where Histograms are going to be placed as matplotlib Figures)
        self.canvas1 = tk.Canvas(self.mainFrame2)
        self.canvas2 = tk.Canvas(self.mainFrame2)

    def grab_start_date(self):
        self.textbox3.insert(tk.END, self.cal.selection_get())

    def grab_end_date(self):
        self.textbox4.insert(tk.END, self.cal.selection_get())

    def pack_all(self):

        self.mainFrame1.place(x=0, y=0, height=600, width=1950)
        self.mainFrame2.place(x=0, y=200, rely=0.05, height=1000, width=1950)

        self.label3.place(x=20, y=70, height=40, width=80)
        self.label4.place(x=20, y=140, height=40, width=80)
        self.label5.place(x=1250, y=30, height=30)
        self.label6.place(x=1250, y=70, height=30)
        self.label7.place(x=1250, y=140, height=30)

        self.textbox1.place(x=80, y=70, height=40, width=800)
        self.textbox2.place(x=80, y=140, height=40, width=800)
        self.textbox3.place(x=1300, y=70, height=40, width=99)
        self.textbox4.place(x=1300, y=140, height=40, width=99)

        self.button1.place(x=900, y=70, height=40, width=120)
        self.button2.place(x=900, y=140, height=40, width=120)
        self.button3.place(x=1100, y=100)
        self.button5.place(x=1450, y=70, height=40)
        self.button4.place(x=1450, y=140, height=40)
        self.button_quit.pack(side=tk.BOTTOM, pady=10)

        self.cal.place(x=1550, y=30, rely=0.005,
                       relx=0.02, height=210, width=300)

        self.canvas1.place(x=100, y=40, height=600, width=800)
        self.canvas2.place(x=1000, y=40, height=600, width=800)

    def open_excel_file_location(self):
        """Open the File Explorer to select desired excel file
        """
        global filepath

        filepath = askopenfilename(filetypes=[(
            "xlsx Files", "*.xlsx"), ("csv Files", "*.csv"), ("All Files", "*.*")])

        if not filepath:
            return
        with open(filepath, "r"):
            self.textbox1.insert(tk.END, filepath)

    def save_results(self):
        """Open the file Explorer to select desired location to save results
        """
        global filepath
        filepath = askdirectory()
        if not filepath:
            return
        self.textbox2.insert(tk.END, filepath)

    def display_plot(self):
        """plot function is created for plotting the graph in tkinter window
        """
        # creating the Tkinter canvas containing the Matplotlib figure
        fig = self.plot_individual_hist()
        fig2, ax = self.plot_4_hist()
        # fig2, ax = MyPlot.my_fxn2()

        self.canvas2 = FigureCanvasTkAgg(fig, master=self.canvas2)
        self.canvas1 = FigureCanvasTkAgg(fig2, master=self.canvas1)
        # self.canvas2 = FigureCanvasTkAgg(fig2, master=self.canvas2)
        self.canvas1.draw()
        self.canvas2.draw()

        # creating the Matplotlib toolbar
        self.toolbar1 = NavigationToolbar2Tk(
            self.canvas1, self, pack_toolbar=False)

        # toolbar = Embedding_in_Tk.NavigationToolbar2Tk(self.canvas, self)
        self.toolbar1.update()

        # placing the canvas on the Tkinter window
        self.toolbar1.pack(side=tk.BOTTOM, padx=400)

        self.canvas1.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        self.canvas2.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

        if self.textbox1 == None:
            messagebox.showinfo("Warning Message",
                                "Please select a valid Excel file path")

    def import_excel(self):
        self.df = pd.read_excel(filepath)
        i = 0
        self.init_date = self.textbox3.get()
        self.init_date[::-1]
        self.final_date = self.textbox4.get()
        self.final_date[::-1]

    def generate_data_from_excel(self):
        """plot simple plot

        Returns:Four figures sharing the same x axis
        """
        self.import_excel()
        print(self.init_date)
        print(self.final_date)

        n_df = self.df[(self.init_date <= self.df.Date_of_birth) &
                  (self.df.Date_of_birth <= self.final_date)]
        print(self.df['Date_of_birth'])
        print("")
        print(n_df['Date_of_birth'])
        print("")

        l = len(n_df['Date_of_birth'])

        df_MWt = pd.DataFrame(n_df.loc[(self.df['Sex'] == 'Male') & (
            n_df['Genotype'] == 'Wildtype') & (n_df['Status'] == 'Alive')])
        df_FWt = pd.DataFrame(n_df.loc[(self.df['Sex'] == 'Female') & (
            n_df['Genotype'] == 'Wildtype') & (n_df['Status'] == 'Alive')])
        df_MHet = pd.DataFrame(n_df.loc[(self.df['Sex'] == 'Male') & (
            n_df['Genotype'] == 'Heterozygous') & (n_df['Status'] == 'Alive')])
        df_FHet = pd.DataFrame(n_df.loc[(self.df['Sex'] == 'Female') & (
            n_df['Genotype'] == 'Heterozygous') & (n_df['Status'] == 'Alive')])
        print(len(df_MWt))
        print(len(df_FWt))
        print(len(df_MHet))
        print(len(df_FHet))
        total_data = [df_MWt, df_FWt, df_MHet, df_FHet]

        j = 0
        self.d2 = []
        for element in total_data:
            age_in_weeks = 0
            d1 = []
            for index, rows in total_data[j].iterrows():
                age_in_weeks = math.floor(n_df['Age_(days)'][index]/7)
                d1.append(age_in_weeks)
            j = j + 1
            d1.sort(reverse=True)
            self.d2.append(d1)
        print(self.d2)
        self.d3 = [self.d2[0][0], self.d2[1][0], self.d2[2][0], self.d2[3][0]]

        # Colors
        self.colors = sns.color_palette("mako", 4)
        self.listy = []
        self.b = int(max(self.d3))+1
        for i in range(self.b-1):

            self.listy.append(i)
            i = i+1

    # def receive_month_from_calendar(self):
        
    #     self.excel_sheet_vector_month = {
    #         (1,'January'), (2,'February'), (3,'March'), (4,'April'), (5,'May'), 
    #         (6,'June'), (7,'July'), (8,'August'), (9,'September'), (10,'October'),
    #         (11,'November'), (12,'December')
    #         }
    #     m = 0
    #     month_list = []
    #     for m in range(10):
    #         if int(self.init_date[6])==self.excel_sheet_vector_month[m][0]:
    #             month_list.append(self.excel_sheet_vector_month[m][1])
    #         else:
    #             print("Agghhh")
    #     n = 10
    #     for n in range(13):
    #         if int(self.init_date[6])==self.excel_sheet_vector_month[n][0][1]:
    #             month_list.append(self.excel_sheet_vector_month[n][1])
    #         else:
    #             print("Agghhh")
    #     print("Minth_list: ",month_list)

    def plot_individual_hist(self):
        self.generate_data_from_excel()
        # self.receive_month_from_calendar()
        # Plot
        fig = Figure(figsize=(5, 5), dpi=100)
        f = fig.gca()
        f.hist(self.d2, bins=self.listy, color=self.colors)
        f.set_xlabel(r'Time (weeks)', fontsize=15)
        f.set_ylabel(r'Number of mice', fontsize=15)
        f.set_title('MICE QUANTITY VS AGE',
                    horizontalalignment='center', fontweight="bold", fontsize=20)

        f.locator_params(axis='x', integer=True)

        return fig

    def plot_4_hist(self):

        self.import_excel()
        # Plot
        fig, axs = plt.subplots(2, 2, figsize=(
            8, 8), sharey=True,  tight_layout=True)

        # I need to garantee its the same x axis quantity and integers
        axs[0, 0].hist(self.d2[0], bins=self.listy, color=self.colors[0],
                       label='Male Wildtype', edgecolor='white')
        axs[0, 1].hist(self.d2[1], bins=self.listy, color=self.colors[1],
                       label='Female Wildtype', edgecolor='white')
        axs[1, 0].hist(self.d2[2], bins=self.listy, color=self.colors[2],
                       label='Male Heterozygous', edgecolor='white')
        axs[1, 1].hist(self.d2[3], bins=self.listy, color=self.colors[3],
                       label='Female Heterozygous', edgecolor='white')

        axs[0, 0].xaxis.set_major_locator(MaxNLocator(integer=True))
        axs[0, 1].xaxis.set_major_locator(MaxNLocator(integer=True))
        axs[1, 0].xaxis.set_major_locator(MaxNLocator(integer=True))
        axs[1, 1].xaxis.set_major_locator(MaxNLocator(integer=True))

        # plt.subplots_adjust(hspace=0)
        axs[1, 0].set_xlabel(r'Time (weeks)', fontsize=15)
        axs[1, 1].set_xlabel(r'Time (weeks)', fontsize=15)
        axs[0, 0].set_ylabel(r'Number of mice', fontsize=20,
                             horizontalalignment='right')

        plt.sca(axs[0, 0])
        plt.title('Male', fontsize=15)
        plt.sca(axs[0, 1])
        plt.title('Female', fontsize=15)

        axs[0, 0].legend()
        axs[0, 1].legend()
        axs[1, 0].legend()
        axs[1, 1].legend()

        plt.suptitle('MICE QUANTITY VS AGE',
                     horizontalalignment='center', fontweight="bold", fontsize=20)
        return fig, axs
        # # I need to garantee its the same x axis quantity and integers
        # axs[0,0].bar(np.arange(0,len(d2[0]),1), d2[0], color=colors[0], label='Male Wildtype', edgecolor='black')
        # axs[0,1].bar(np.arange(0,len(d2[1]),1), d2[1], color=colors[1], label='Female Wildtype', edgecolor='black')
        # axs[1,0].bar(np.arange(0,len(d2[2]),1), d2[2], color=colors[2], label='Male Heterozygous', edgecolor='black')
        # axs[1,1].bar(np.arange(0,len(d2[3]),1), d2[3], color=colors[3], label='Female Heterozygous', edgecolor='black')

        # axs[0,0].xaxis.set_major_locator(MaxNLocator(integer=True))
        # axs[0,1].xaxis.set_major_locator(MaxNLocator(integer=True))
        # axs[1,0].xaxis.set_major_locator(MaxNLocator(integer=True))
        # axs[1,1].xaxis.set_major_locator(MaxNLocator(integer=True))

        # #plt.subplots_adjust(hspace=0)
        # axs[1,0].set_xlabel(r'Time (weeks)', fontsize=15)
        # axs[1,1].set_xlabel(r'Time (weeks)', fontsize=15)
        # axs[0,0].set_ylabel(r'Number of mice', fontsize =20, horizontalalignment='right')

        # plt.sca(axs[0,0])
        # plt.title('Male', fontsize=15)
        # plt.sca(axs[0,1])
        # plt.title('Female', fontsize=15)

        # axs[0,0].legend()
        # axs[0,1].legend()
        # axs[1,0].legend()
        # axs[1,1].legend()


app = Application()
app.geometry("1900x990+0+0")
app.resizable(True, False)
app.iconbitmap(r'../docs/mickey.ico')
app.mainloop()
