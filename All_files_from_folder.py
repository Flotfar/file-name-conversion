# this code is running the os-module which is a 
# part of the standard library, or stdlib, within Python 3.

import os
import xlsxwriter

folder = r'/Users/joakimbonde/Desktop/File_test/'   # file path, rapped with 'r'
file_num = "HBA.2.9_1_"                             # file initial numbering


#########################################################
#### Creating Excel file for exportation of results: ####
#########################################################

workbook = xlsxwriter.Workbook('File_info.xlsx')
worksheet = workbook.add_worksheet('Names')

#formats
title_format = workbook.add_format({'bold': True, 'font_size': '14', 'border': True})
name_format = workbook.add_format({'font_size': '11', 'border': True})

worksheet.write(0, 0, "Old_names:" , title_format)
worksheet.write(0, 1, "New names:" , title_format)


#########################################################
#### Extracting and renaming files:                  ####
#########################################################

row = 1
# count increase by 1 in each iteration
# iterate all files from a directory

for file_name in os.listdir(folder):
    # Constructing old file name and saving it to array:
    source = folder + file_name
    worksheet.write(row, 0, str(file_name), name_format)

    # Adding the numbering to the new file name:
    destination = folder + file_num + file_name
    worksheet.write(row, 1, str(file_num + file_name), name_format)

    # Renaming the file
    os.rename(source, destination)
    row += 1


workbook.close()
