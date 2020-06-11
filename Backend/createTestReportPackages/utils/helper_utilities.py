import os

from createTestReportPackages.parser import CONFIG


def write_report_in_dir(document_name, status_repost):
    status_repost_file_name = os.path.join(CONFIG["tempFolder"], document_name, CONFIG["status_file_name"])
    f = open(status_repost_file_name, "w")
    f.write(status_repost)
    f.close()
