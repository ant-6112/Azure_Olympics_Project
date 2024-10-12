import os
import pandas as pd

class FileHandler:
    def __init__(self, shared_drive_path):
        self.shared_drive_path = shared_drive_path

    def get_available_weeks(self):
        files = os.listdir(self.shared_drive_path)
        weeks = [f.split('_')[2].split('.')[0] for f in files if f.startswith('Risk_Week_')]
        return sorted(weeks, reverse=True)

    def read_risk_report(self, week):
        file_path = os.path.join(self.shared_drive_path, f"Risk_Week_{week}.xlsx")
        return pd.read_excel(file_path)

    def read_mapping_file(self):
        file_path = os.path.join(self.shared_drive_path, "Mapping_File.xlsx")
        return pd.read_excel(file_path)
