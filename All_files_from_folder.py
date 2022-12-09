# this code is running the os-module which is a 
# part of the standard library, or stdlib, within Python 3.

import os
import shutil    
import xlsxwriter

folder = r'C:\Users\jbmk\OneDrive - Ramboll\Desktop\File Renaming PY\Data_input'   # file path, rapped with 'r'
dest_folder = r'C:\Users\jbmk\OneDrive - Ramboll\Desktop\File Renaming PY\Data_output'  # destination folder

panel_num = "HBA.2.9_1_"            # file initial numbering

#########################################################
#### Creating Excel file for exportation of results: ####
#########################################################

workbook = xlsxwriter.Workbook('File_info.xlsx')
worksheet = workbook.add_worksheet('Names')

#formats
title_format = workbook.add_format({'bold': True, 'font_size': '14', 'border': True})
name_format = workbook.add_format({'font_size': '11', 'border': True})

worksheet.write(0, 0, "Gammelt filnavn:" , title_format)
worksheet.write(0, 1, "Filnavn:" , title_format)
worksheet.write(0, 2, "Tegningsnummer:" , title_format)
worksheet.write(0, 3, "Tegningsnavn:" , title_format)


#########################################################
#### Extracting and renaming files:                  ####
#########################################################

row = 1
count_B = 1
count_C = 1

# count increase by 1 in each iteration
# iterate all files from a directory

####### Folder A #######

for fname_A in os.listdir(folder):

    path_A = os.path.join(folder, fname_A)
    if os.path.isfile(path_A):

        # Copying the file to output directory:
        dest_A = os.path.join(dest_folder, fname_A)  # creating filename in output folder 
        shutil.copyfile(path_A, dest_A)
        
        # Constructing old file name and saving it to excel:
        worksheet.write(row, 0, str(fname_A), name_format)

        # Constructing new file name with numbering and savin ít to excel:
        new_name = panel_num + fname_A
        destination = os.path.join(dest_folder, new_name)
        worksheet.write(row, 1, str(new_name), name_format)

        # Renaming the file
        os.rename(dest_A, destination)
        row += 1

####### If folder B #######

    if os.path.isdir(path_A):
        for fname_B in os.listdir(path_A):
            panel_num_B = panel_num + '_' + count_B + '_'
            
            path_B = os.path.join(path_A, fname_B)
            if os.path.isfile(path_B):

                # Copying the file to output directory:
                dest_B = os.path.join(dest_folder, fname_B)  # creating filename in output folder 
                shutil.copyfile(path_B, dest_B)
                
                # Constructing old file name and saving it to excel:
                worksheet.write(row, 0, str(fname_B), name_format)

                # Constructing new file name with numbering and savin ít to excel:
                new_name = panel_num_B + fname_B
                destination = os.path.join(dest_folder , new_name)
                worksheet.write(row, 1, str(new_name), name_format)

                # Renaming the file
                os.rename(dest_B, destination)
                row += 1

####### If folder C #######

             if os.path.isdir(path_B):
                for fname_C in os.listdir(path_B):
                    panel_num_C = panel_num + '_' + count_B + '_' + count_C + '_'

                    path_C = os.path.join(path_B, fname_C)
                    if os.path.isfile(path_C):

                        # Copying the file to output directory:
                        dest_C = os.path.join(dest_folder, fname_C)  # creating filename in output folder 
                        shutil.copyfile(path_C, dest_C)
                        
                        # Constructing old file name and saving it to excel:
                        worksheet.write(row, 0, str(fname_C), name_format)

                        # Constructing new file name with numbering and savin ít to excel:
                        new_name = panel_num_C + fname_C
                        destination = os.path.join(dest_folder , new_name)
                        worksheet.write(row, 1, str(new_name), name_format)

                        # Renaming the file
                        os.rename(dest_C, destination)
                        row += 1
                    
                    count_C += 1

            count_B += 1
    







workbook.close()
