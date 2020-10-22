# -*- coding: utf-8 -*-

from google_api.drive import Drive


drive = Drive()

folder_id = "1R_B5ZLcxWBmq6G9aGd0m2d_j2uPihXGQ"  # Google Drive Folder ID

test_file_name = "Test File"
test_sheet_name = "Test Sheet"

test_txt_filename = "samples/test_file.txt"
test_sheet_filename = "samples/test_sheet.xlsx"

# Upload txt file
file_id = drive.create_file(folder_id, test_file_name, test_txt_filename)
url = "https://drive.google.com/file/d/" + file_id +"/view?usp=sharing"
print(url)

# Upload xlsx
file_id = drive.create_sheet(folder_id, test_sheet_name, test_sheet_filename)
url = "https://drive.google.com/file/d/" + file_id +"/view?usp=sharing"
print(url)
