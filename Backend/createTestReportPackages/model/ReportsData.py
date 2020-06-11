import json
import os
import pandas as pd

from createTestReportPackages.parser import CONFIG


class ReportsData:
    def __init__(self):
        if os.path.exists("reportFileDataframe.pkl"):
            self.REPORT_FILE_DATA_FRAME = pd.read_pickle("reportFileDataframe.pkl")
        else:
            self.REPORT_FILE_DATA_FRAME = self.initialize_dataframe()
            pd.to_pickle(self.REPORT_FILE_DATA_FRAME, "reportFileDataframe.pkl")


    def initialize_dataframe(self):
        list_dict = []
        for report in os.listdir(CONFIG["tempFolder"]):
            status_file_name = os.path.join(CONFIG["tempFolder"], report, CONFIG["status_file_name"])
            status_text = open(status_file_name, "r").read()
            new_row = json.loads(status_text)
            list_dict.append(new_row)
        return pd.DataFrame(list_dict)
