import os
import math
import datetime


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

        self.directory_per_week = os.path.join(
            self.current_directory,
            "results",
            "{}_{}".format(
                self.year,
                "weekly_results"
            )
        )

    def create_folders(self):
        if not os.path.exists(self.directory_per_month):
            os.makedirs(self.directory_per_month)

        if not os.path.exists(self.directory_per_week):
            os.makedirs(self.directory_per_week)


if __name__ == "__main__":
    fm = FolderManager()
