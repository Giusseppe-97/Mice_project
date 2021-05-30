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
from pandas.core.series import Series
import tkcalendar as tkc
from tkinter import messagebox

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk)
import matplotlib.pyplot as plt

from matplotlib.ticker import MaxNLocator
import pandas as pd
import numpy as np
import math
import seaborn as sns

import os
from datetime import datetime as dt
import datetime


from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl

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
            self.mainFrame1, selectbackground="#120597", background="#120597",
            selectmode="day", year=2021, month=5, day=1
        )

        # Call methods
        self.configure_basic_tk_properties()
        self.pack_all()

# reset button not programmed yet
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
        self.lable = tk.Label(
            self.mainFrame1, text="RUN DATA", foreground="white",
            background="#120597").place(x=0, width=1920
                                        )
        self.lable2 = tk.Label(
            self.mainFrame2, text="HISTOGRAM DISPLAY PREVIEW", foreground="white",
            background="#120597").place(y=0, width=1920
                                        )

        # Create input and output lables
        self.label3 = ttk.Label(
            self.mainFrame1, text="Input: ", background="white"
        )
        self.label4 = ttk.Label(
            self.mainFrame1, text="Output:", background="white"
        )
        self.label5 = tk.Label(
            self.mainFrame1, text="Select Date Interval: ",
            font=("Arial bold", 11), background="white"
        )
        self.label6 = ttk.Label(
            self.mainFrame1, text="From:", background="white"
        )
        self.label7 = ttk.Label(
            self.mainFrame1, text="To:", background="white"
        )

        # Create textboxes
        self.textbox1 = ttk.Entry(self.mainFrame1, width=80)
        self.textbox2 = ttk.Entry(self.mainFrame1, width=80)
        self.textbox3 = ttk.Entry(self.mainFrame1, width=20)
        self.textbox4 = ttk.Entry(self.mainFrame1, width=20)

        # Create and inicialize buttons
        self.button1 = ttk.Button(
            self.mainFrame1, text="Select", command=lambda: [self.open_excel_file_location()]
        )
        self.button2 = ttk.Button(
            self.mainFrame1, text="Save ", command=lambda: [self.save_results()]
        )
        self.button3 = ttk.Checkbutton(
            self.mainFrame1, text="RUN", style='ToggleButton', command=self.display_plot
        )
        self.button5 = ttk.Button(
            self, text="Start date", command=self.grab_start_date
        )
        self.button4 = ttk.Button(
            self, text="End date", command=self.grab_end_date
        )
        self.button_quit = ttk.Button(
            master=self, text="Quit", command=self.quit
        )
        # self.button_reset = ttk.Button(
        #     master=self, text="Reset", command=lambda:[self.quit, self.reset_app])

        # Create Canvas (where Histograms are going to be placed as matplotlib Figures)
        self.canvas1 = tk.Canvas(self.mainFrame2)
        self.canvas2 = tk.Canvas(self.mainFrame2)

    def grab_start_date(self):
        self.textbox3.insert(tk.END, self.cal.selection_get())
        self.init_month = self.cal.selection_get().strftime("%B")

    def grab_end_date(self):
        self.textbox4.insert(tk.END, self.cal.selection_get())
        self.final_month = self.cal.selection_get().strftime("%B")
# pack and reset button not placed yet, just packed

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
        # self.button_reset.pack(side=tk.BOTTOM, pady=100)

        self.cal.place(x=1550, y=30, rely=0.005,
                       relx=0.02, height=210, width=300)

        self.canvas1.place(x=100, y=40, height=600, width=800)
        self.canvas2.place(x=1000, y=40, height=600, width=800)
# Data tale path reproducible to other devices, not only for this one

    def open_excel_file_location(self):
        """Open the File Explorer to select desired excel file
        """
        global filepath1
        filepath1 = askopenfilename(filetypes=[(
            "xlsx Files", "*.xlsx"), ("csv Files", "*.csv"), ("All Files", "*.*")])

        if not filepath1:
            filepath1 = "../data/Mice_table.xlsx"
        with open(filepath1, "r"):
            self.textbox1.insert(tk.END, filepath1)
# Save path now reproducible to other devices, not only for this one

    def save_results(self):
        """Open the file Explorer to select desired location to save results
        """
        global filepath2
        filepath2 = askdirectory()
        if not filepath2:
            filepath2 = "results/2021_monthly_results/plots_per_month"
        self.textbox2.insert(tk.END, filepath2)

    def display_plot(self):
        """plot function is created for plotting the graph in tkinter window
        """
        # create the Tkinter canvas containing the Matplotlib figures
        fig = self.plot_individual_hist()
        fig2, ax = self.plot_4_hist()

        # create canvas and draw figures into canvas
        self.canvas2 = FigureCanvasTkAgg(fig, master=self.canvas2)
        self.canvas1 = FigureCanvasTkAgg(fig2, master=self.canvas1)
        self.canvas1.draw()
        self.canvas2.draw()

        # create the Matplotlib toolbars
        self.toolbar1 = NavigationToolbar2Tk(
            self.canvas1, self, pack_toolbar=False)
        self.toolbar2 = NavigationToolbar2Tk(
            self.canvas2, self, pack_toolbar=False)
        self.toolbar1.update()
        self.toolbar2.update()

        # place the canvas on the Tkinter window
        self.toolbar1.place(x=400, y=900)
        self.toolbar2.place(x=1300, y=900)

# better if they where placed # pack the widgetsinside the Tkinter canvas
        self.canvas1.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        self.canvas2.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

    def import_excel_file(self):
        self.df = pd.read_excel(filepath1)

    def obtain_data_from_excel(self):
        self.import_excel_file()

# I could obtain the date in the correct format from the begining
        # Change date format to compare with excel
        self.init_date = self.textbox3.get()
        self.init_date[::-1]
        self.final_date = self.textbox4.get()
        self.final_date[::-1]

        # Limit data from the excel file for the chosen period of time
        n_df = self.df[(self.init_date <= self.df.Date_of_birth) &
                       (self.df.Date_of_birth <= self.final_date)]

        l = len(n_df['Date_of_birth'])

        # Selects data from the excel file for sex, genotype and status (only mice that are alive)
        df_MWt = pd.DataFrame(n_df.loc[(self.df['Sex'] == 'Male') & (
            n_df['Genotype'] == 'Wildtype') & (n_df['Status'] == 'Alive')])

        birth_MWt = []
        birth_FWt = []
        birth_MHet = []
        birth_FHet = []

        for index, rows in df_MWt.iterrows():
            a = str(n_df['Date_of_birth'][index])
            birth_MWt.append(a[:10])
        print(len(birth_MWt))

# There are 4 loops that are doing the same thing. There should be a way to reduce this
        df_FWt = pd.DataFrame(n_df.loc[(self.df['Sex'] == 'Female') & (
            n_df['Genotype'] == 'Wildtype') & (n_df['Status'] == 'Alive')])

        for index, rows in df_FWt.iterrows():
            a = str(n_df['Date_of_birth'][index])
            birth_FWt.append(a[:10])

        df_MHet = pd.DataFrame(n_df.loc[(self.df['Sex'] == 'Male') & (
            n_df['Genotype'] == 'Heterozygous') & (n_df['Status'] == 'Alive')])

        for index, rows in df_MHet.iterrows():
            a = str(n_df['Date_of_birth'][index])
            birth_MHet.append(a[:10])

        df_FHet = pd.DataFrame(n_df.loc[(self.df['Sex'] == 'Female') & (
            n_df['Genotype'] == 'Heterozygous') & (n_df['Status'] == 'Alive')])

        for index, rows in df_FHet.iterrows():
            a = str(n_df['Date_of_birth'][index])
            birth_FHet.append(a[:10])

        birthday = [birth_MWt, birth_FWt, birth_MHet, birth_FHet]
        # total_data = [df_MWt, df_FWt, df_MHet, df_FHet]
        print(n_df)

        self.d2 = []

        for i in range(len(birthday)):
            age_in_weeks = 0
            d1 = []
            j = 0
            for j in range(len(birthday[i])):

                a1 = str(self.final_date)
                aa = dt.strptime(a1, "%Y-%m-%d")
                b1 = str(birthday[i][j])
                bb = dt.strptime(b1, "%Y-%m-%d")
                bd = abs((bb - aa).days)
                age_in_weeks = bd//7

                d1.append(age_in_weeks)
                j = j + 1
            d1.sort(reverse=True)
            self.d2.append(d1)
        print(self.d2)

        self.d3 = []
        self.d4 = []
        es = 0
        for es in range(4):
            if self.d2[es] != []:
                self.d3.append(self.d2[es][0])
            else:
                self.d4.append(es)

        # Colors
        self.colors = sns.color_palette("mako", 4)
        self.listy = []

        if len(self.d4) == 4:
            self.b = 3
        else:
            self.b = int(max(self.d3))+1

        for i in range(self.b):
            self.listy.append(i)
            i = i+1

    def plot_individual_hist(self):

        self.obtain_data_from_excel()

        # Plot
        plt.ticklabel_format(style='sci')
        fig = plt.figure(figsize=(5, 5), dpi=100)
        f = fig.gca()
        # ax.xaxis.set_major_locator(MaxNLocator(integer=True))
        # ax.yaxis.set_major_locator(MaxNLocator(integer=True))

        f.hist(self.d2, bins=self.listy, color=self.colors)
        f.set_xlabel(r'Age (weeks)', fontsize=15)
        f.set_ylabel(r'Number of mice', fontsize=15)
        f.set_title('NUMBER OF MICE VS AGE',
                    horizontalalignment='center', fontweight="bold", fontsize=20)

        f.legend(['M_WT', 'F_WT', 'M_HET', 'F_HET'])

        if self.init_month != self.final_month:
            self.plot_name = str(self.init_month) + "-" + \
                str(self.final_month) + "_histogram"
        else:
            self.plot_name = str(self.init_month) + "_histogram"

        filepath_hist_plot = str(filepath2) + "/" + \
            str(self.plot_name) + ".png"
        hist_plot = fig.savefig(filepath_hist_plot)

        return fig

    def plot_4_hist(self):

        self.obtain_data_from_excel()
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
        axs[0, 0].yaxis.set_major_locator(MaxNLocator(integer=True))
        axs[0, 1].yaxis.set_major_locator(MaxNLocator(integer=True))
        axs[1, 0].yaxis.set_major_locator(MaxNLocator(integer=True))
        axs[1, 1].yaxis.set_major_locator(MaxNLocator(integer=True))

        # plt.subplots_adjust(hspace=0)
        axs[1, 0].set_xlabel(r'Age (weeks)', fontsize=15)
        axs[1, 1].set_xlabel(r'Age (weeks)', fontsize=15)
        axs[0, 0].set_ylabel(r'Number of mice', fontsize=20,
                             horizontalalignment='right')

        axs[0, 0].legend()
        axs[0, 1].legend()
        axs[1, 0].legend()
        axs[1, 1].legend()

        plt.suptitle('NUMBER OF MICE VS AGE',
                     horizontalalignment='center', fontweight="bold", fontsize=20)

        if self.init_month != self.final_month:
            self.plot_4_name = str(self.init_month) + \
                "-" + str(self.final_month)
        else:
            self.plot_4_name = str(self.init_month)

        self.filepath_4_plot = str(filepath2) + "/" + \
            str(self.plot_4_name)+".png"
        hist_4_plot = plt.savefig(self.filepath_4_plot)
        self.hist_4_plot = openpyxl.drawing.image.Image(self.filepath_4_plot)

        return fig, axs

    def reset_app(self):
        self.pop_up_message()


class FolderManager(Application):
    """
    Enables the creation and management of folder structures for multipurpose 
    projects with datetime configurations.
    :param project_name: name of the project for the FolderManager agent.
    """

    def __init__(self):
        self.current_datetime = dt.now()
        self.get_current_important_values()
        self.generate_folder_paths()
        self.create_folders()
        # self.create_excel_workbook_weeks()

    def get_current_important_values(self):
        self.year = self.current_datetime.strftime("%Y")
        self.month = self.current_datetime.strftime("%B")
        self.week = self.current_datetime.strftime("%W")

    def generate_folder_paths(self):
        self.current_directory = os.path.abspath(os.path.dirname(__file__))

        self.directory_per_month = os.path.join(
            self.current_directory,
            "results",
            "{}_{}".format(
                self.year,
                "monthly_results"
            )
        )
        self.directory_plots_month = os.path.join(
            "results",
            "2021_monthly_results",
            "{}".format("plots_per_month")
        )

    def create_folders(self):
        if not os.path.exists(self.directory_per_month):
            os.makedirs(self.directory_per_month)

        if not os.path.exists(self.directory_plots_month):
            os.makedirs(self.directory_plots_month)

    def create_jpeg_from_figures(self):
        pass


class ExcelDevelopement(Application):

    def __init__(self):

        self.create_excel_workbook_months()

    def create_excel_workbook_months(self):

        d2 = [[4, 4, 4, 4, 4, 4, 4], [4, 4, 4, 4], [4, 4, 4], [4, 4, 4, 4, 4, 4]]
        months_2021_excel_wb = Workbook()
        worksheet_month_ws = months_2021_excel_wb.active
        worksheet_month_ws.title = "Changed Sheet"

        longest_list_size = len(d2[0])
        list_dic = []
        i = 0
        j = 0
        k = 0
        l = 0

        for k in range(4):
            if len(d2[k]) > len(d2[0]):
                longest_list_size = len(d2[k])
            

        dic_mic = {}
        for longest_list_size in range(longest_list_size, -1, -1):
            dic_mic.setdefault(longest_list_size, []).append(0)

        list_of_dic = []
        data = ['Male Wildtype', 'Female Wildtype',
                'Male Heterozygous', 'Female Heterozygous']

        for i in range(4):

            dic_mice = dict()
            num_of_mice = []
            age_mice_weeks = []
            for j in range(len(d2[i])):

                if d2[i][j] not in age_mice_weeks:
                    age_mice_weeks.append(d2[i][j])
                    num_of_mice = d2[i].count(d2[i][j])
                    dic_mice.setdefault(d2[i][j], []).append(num_of_mice)

            dic_mice = {key: dic_mice.get(
                key, dic_mic[key]) for key in dic_mic}
            df = pd.DataFrame.from_dict(dic_mice)
            list_dic.append(df)

            print(data[i])
            print(df.to_string(index=False))
        print(list_dic)

        result = pd.concat(list_dic)
        print(result.to_string(index=False))

        rows = dataframe_to_rows(result, index=False)

        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                worksheet_month_ws.cell(row=r_idx, column=c_idx, value=value)

        months_2021_excel_wb.save(filename='results/sample_book.xlsx')

        # self.hist_4_plot.anchor(self.plot_name.cell('A20'))
        # self.plot_name.add_image(self.hist_4_plot)
        # months_2021_excel_wb.save("{}\\{}".format(
        #     self.directory_per_month, '{}.xlsx'.format("Months_2021")))


if __name__ == "__main__":
    app = Application()
    app.geometry("1900x990+0+0")
    app.resizable(True, False)
    app.iconbitmap(r'../docs/mickey.ico')
    app.mainloop()
    fm = FolderManager()
    ed = ExcelDevelopement()
