
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


a = datetime.strptime("2021-02-22", "%Y-%m-%d")
b = datetime.strptime("2021-02-22", "%Y-%m-%d")
c = abs((b-a).days)
d = abs((a-b).days)

print(c)
print(d)


# d2 = [[6 ,6 ,6 ,5 ,5, 2, 2, 2, 2, 1, 0], [4, 3, 3, 2, 2, 1, 1, 1, 1, 0, 0, 0, 0], [0], [3, 1, 0]]
# months_2021_excel_wb = Workbook()
# worksheet_month_ws =  months_2021_excel_wb.active
# worksheet_month_ws.title = "Changed Sheet"

# longest_list_size = 0
# list_dic = []
# i = 0
# j = 0
# k = 0
# l = 0

# for k in range(4):
#         if len(d2[k])>len(d2[0]):
#                 longest_list_size = len(d2[k]) 

# dic_mic={}
# for longest_list_size in range(longest_list_size,-1,-1):
#         dic_mic.setdefault(longest_list_size,[]).append(0)

# list_of_dic = []
# data = ['Male Wildtype', 'Female Wildtype', 'Male Heterozygous', 'Female Heterozygous']

# for i in range(4):
        
#         dic_mice = dict()
#         num_of_mice = []
#         age_mice_weeks = []
#         for j in range(len(d2[i])):

#                 if d2[i][j] not in age_mice_weeks:
#                         age_mice_weeks.append(d2[i][j])
#                         num_of_mice = d2[i].count(d2[i][j])
#                         dic_mice.setdefault(d2[i][j],[]).append(num_of_mice)

#         dic_mice = {key: dic_mice.get(key, dic_mic[key]) for key in dic_mic}
#         df = pd.DataFrame.from_dict(dic_mice) 
#         list_dic.append(df)
        
#         print(data[i])
#         print(df.to_string(index=False))
# print(list_dic)

# result = pd.concat(list_dic)
# print(result.to_string( index = False))

# rows = dataframe_to_rows(result,index = False)

# for r_idx, row in enumerate(rows, 1):
#         for c_idx, value in enumerate(row, 1):
#                 worksheet_month_ws.cell(row=r_idx, column=c_idx, value=value)

# months_2021_excel_wb.save(filename = 'results/sample_book.xlsx')
 


# birthdays = []
#         i = 0
#         j = 0
#         for index, rows in total_data[i].iterrows():
#             aa = str(total_data[i]['Date_of_birth'])
#             print(aa[0:9:-1])
#             bb = str(self.final_date)
#             # print(bb)
            
#             birthdays.append(aa)

#         print(birthdays)

# dt = pd.DataFrame({
#             'Male Wildtype': d2[0]},{
#                 'Female Wildtype': d2[1]},{
#                     'Male Heterozygous': d2[2]},{
#                         'Female Heterozygous': d2[3]})

# print(dt)

# self.pop_up_messages()

# # Create scrollbar
# self.scroll_frame_1 = ttk.Scrollbar(self.mainFrame1, orient="vertical", command=self.canvas1.yview)
# self.scroll_frame_1.pack(side='right', fill='y')




# def pop_up_messages(self):
#     if str(self.textbox1) == "" and str(self.textbox3) != "":
#         messagebox.showinfo("Warning Message: Empty input filepath",
#                             "Please select a valid Excel file or write its path before continuing")

#     if str(self.textbox2) == "" and str(self.textbox3) != "" & str(self.textbox4) != "":
#         messagebox.showinfo("Warning Message: Empty output filepath",
#                             "Please select a valid folder")

#     if str(self.textbox1)[-1] != "x":
#         messagebox.showinfo("Warning Message",
#                             "Please select a valid Excel file path before continuing")

#     if str(self.textbox3)[3] != "1" or str(self.textbox4)[3] != "1":
#         messagebox.showinfo("Warning Message: Invalid date or date period",
#                             "Please select a valid date (from 2021 further)")


# self.directory_per_week = os.path.join(
        #     self.current_directory,
        #     "results",
        #     "{}_{}".format(
        #         self.year,
        #         "weekly_results"
        #     )
        # )
        # self.directory_plots_week = os.path.join(
        #     "results",
        #     "2021_weekly_results",
        #     "{}".format("plots_per_week")
        # )

        # if not os.path.exists(self.directory_per_week):
        #     os.makedirs(self.directory_per_week)
        # if not os.path.exists(self.directory_plots_week):
        #     os.makedirs(self.directory_plots_week)




# def create_excel_workbook_weeks(self):
        # Create excel workbooks (xlsx)

        # self.weeks_2021_excel = xlsxwriter.Workbook("{}\\{}".format(
        #     self.directory_per_week, '{}.xlsx'.format("Weeks_2021")))

        # i = 1
        # self.excel_sheet_vector_week = []
        # for i in range(53):
        #     self.excel_sheet_vector_week.append("week " + str(i))

        # # Esthetic paramenters for Week's excel
        # normal_text = self.weeks_2021_excel.add_format(
        #     {'bold': False, "align": "center"}
        # )
        # bold = self.weeks_2021_excel.add_format(
        #     {'bold': True, "align": "center"}
        # )
        # bold_border = self.weeks_2021_excel.add_format(
        #     {"bold": True, "align": "center", "border": True}
        # )

        # self.excel_sheet_vector_week.append(self.weeks_2021_excel.add_worksheet(
        #     "{}".format(self.excel_sheet_vector_week[0])))

        # self.weeks_2021_excel.close()







        # worksheet_month.write('A2', mg_o.df_MWt)
        # worksheet_month.write('D2', mg_o.df_FWt)
        # worksheet_month.write('G2', mg_o.df_MHet)
        # worksheet_month.write('J2', mg_o.df_FHet)
        # filep = str(FigureManager.generate_figures)
        # worksheet_month.insert_image('D14', filep )

        # # Esthetic paramenters for Month's excel
        # normal_text = self.months_2021_excel.add_format(
        #     {'bold': False, "align": "center"}
        #     )
        # bold = self.months_2021_excel.add_format(
        #     {'bold': True, "align": "center"}
        #     )
        # bold_border = self.months_2021_excel.add_format(
        #     {"bold": True, "align": "center", "border": True}
        #     )

        # # time datetime for loop to add sheets in positions of the vector
        # self.excel_sheet_vector_month.append(self.months_2021_excel.add_worksheet(
        #         "{}".format(self.excel_sheet_vector_month[0][1])))