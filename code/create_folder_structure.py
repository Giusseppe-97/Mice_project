import os
import datetime
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
# from mice_gui import import_excel_file

class FolderManager:
    """
    Enables the creation and management of folder structures for multipurpose 
    projects with datetime configurations.
    :param project_name: name of the project for the FolderManager agent.
    """
    
    def __init__(self):
        self.current_datetime = datetime.datetime.now()
        self.get_current_important_values()
        self.generate_folder_paths()
        self.create_folders()
        self.create_excel_workbook_weeks()
        self.create_excel_workbook_months()

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

        self.directory_per_week = os.path.join(
            self.current_directory,
            "results",
            "{}_{}".format(
                self.year,
                "weekly_results"
            )
        )
        
        self.directory_plots_week = os.path.join(
            "results",
            "2021_weekly_results",
            "{}".format("plots_per_week")
        )

    def create_folders(self):
        if not os.path.exists(self.directory_per_month):
            os.makedirs(self.directory_per_month)

        if not os.path.exists(self.directory_per_week):
            os.makedirs(self.directory_per_week)

        if not os.path.exists(self.directory_plots_month):
            os.makedirs(self.directory_plots_month)

        if not os.path.exists(self.directory_plots_week):
            os.makedirs(self.directory_plots_week)

    def create_jpeg_from_figures(self):
        pass

    def create_excel_workbook_weeks(self):
        # Create excel workbooks (xlsx)
        
        self.weeks_2021_excel = xlsxwriter.Workbook("{}\\{}".format(
            self.directory_per_week, '{}.xlsx'.format("Weeks_2021")))
        
        i = 1
        self.excel_sheet_vector_week = []
        for i in range(53):
            self.excel_sheet_vector_week.append("week " + str(i))

        # Esthetic paramenters for Week's excel
        normal_text = self.weeks_2021_excel.add_format(
            {'bold': False, "align": "center"}
            ) 
        bold = self.weeks_2021_excel.add_format(
            {'bold': True, "align": "center"}
            )
        bold_border = self.weeks_2021_excel.add_format(
            {"bold": True, "align": "center", "border": True}
        )

        self.excel_sheet_vector_week.append(self.weeks_2021_excel.add_worksheet(
                "{}".format(self.excel_sheet_vector_week[0])))

        self.weeks_2021_excel.close()
        
    def create_excel_workbook_months(self):
        # Create excel workbooks (xlsx)
        self.months_2021_excel = xlsxwriter.Workbook("{}\\{}".format(
            self.directory_per_month, '{}.xlsx'.format("Months_2021")))

        self.excel_sheet_vector_month = [
            (1,'January'), (2,'February'), (3,'March'), (4,'April'), (5,'May'), 
            (6,'June'), (7,'July'), (8,'August'), (9,'September'), (10,'October'),
            (11,'November'), (12,'December')
            ]
        # Esthetic paramenters for Month's excel
        normal_text = self.months_2021_excel.add_format(
            {'bold': False, "align": "center"}
            ) 
        bold = self.months_2021_excel.add_format(
            {'bold': True, "align": "center"}
            )
        bold_border = self.months_2021_excel.add_format(
            {"bold": True, "align": "center", "border": True}
            )
        
        # time datetime for loop to add sheets in positions of the vector
        self.excel_sheet_vector_month.append(self.months_2021_excel.add_worksheet(
                "{}".format(self.excel_sheet_vector_month[0][1])))

        self.months_2021_excel.close()
        

if __name__ == "__main__":
    fm = FolderManager()
