# this code is running the os-module which is a 
# part of the standard library, or stdlib, within Python 3.


#########################################################
#### Rules of the script:                            ####
#########################################################

# All files belonging to vedligehold should be placed within that folder in input
# All files belonging to 
# All lose files should be placed in the open invironment of the input folder

#########################################################

import os
import shutil    
import xlsxwriter
import pandas as pd

folder = r'C:\Users\jbmk\OneDrive - Ramboll\Desktop\SoB - Doc\File Renaming PY\Data_input'   # file path, rapped with 'r'
dest_folder = r'C:\Users\jbmk\OneDrive - Ramboll\Desktop\SoB - Doc\File Renaming PY\Data_output'  # destination folder


panel_num = "HPA1.7"       # file initial numbering
main_folder = ["Vedligeh" , "Indstl.Opstil"]    #Main folder naming


###############################################################################
#### Extracting general File information from Collective.xlsx (File_load): ####
###############################################################################

# Importing the excel file, sheet.no 5 "File_load"
df = pd.read_excel(r'C:\Users\jbmk\OneDrive - Ramboll\Desktop\SoB - Doc\File Renaming PY\File_collective.xlsx', sheet_name='File_load')

# Refferencing columns to variables:
df_ofname = df['Gammelt Filnavn']
df_rev = df['Revision']
df_date = df['Udgavedato'].apply(lambda x: x.strftime('%d-%m-%Y'))   #Clearing out timestamps from the Date
df_dname = df['Tegningsnavn']


#########################################################
#### Creating Excel file for exportation of results: ####
#########################################################

workbook = xlsxwriter.Workbook('File_info.xlsx')
worksheet = workbook.add_worksheet('Names')

#formats
title_format = workbook.add_format({'bold': True, 'font_size': '14', 'border': False})
name_format = workbook.add_format({'font_size': '11', 'border': False})

worksheet.write(0, 0, "Gammelt filnavn:" , title_format)
worksheet.write(0, 1, "Tegningsnummer:" , title_format)
worksheet.write(0, 2, "Revision:" , title_format)
worksheet.write(0, 3, "Udgavedato:" , title_format)
worksheet.write(0, 4, "Tegningsnavn:" , title_format)
worksheet.write(0, 5, "Filnavn:" , title_format)




############################################################
#### Deleteing existing files within output directory:  ####
############################################################

for file in os.listdir(dest_folder):
    dest_folder_path = os.path.join(dest_folder, file)
    os.remove(dest_folder_path)



#########################################################
#### Extracting and renaming files:                  ####
#########################################################

row = 1
count = 0
# count_B = 1
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

        # Constructing new file name with numbering:
        new_name = panel_num + '_' + fname_A
        destination = os.path.join(dest_folder, new_name)

        # Renaming the file
        os.rename(dest_A, destination)

        # Extracting rev, date, drawing.no from File_collective.xlsx
        # ofname = Old File name,  rev = revision,  dname = Drawing name,  date = Issue date
        for index in range(len(df)):
            if df_ofname[index] == fname_A:
                dname = df_dname[index]
                rev = str(df_rev[index])
                date = str(df_date[index])
                break
            elif df_ofname[index].startswith("B 4.13.2.2 Transport, Idriftsætning og Vedligeholdelse") == True:
                dname = "Løgstrup. B 4.13.2.2 Transport, Idriftsætning og Vedligeholdelse"
                rev = "2"
                date = "28-06-2017"
                break
            else:
                dname = ""
                rev = ""
                date = ""

        # Adding information to Excel:
        worksheet.write(row, 0, str(fname_A), name_format)    # Old File Name
        worksheet.write(row, 1, str(panel_num), name_format)  # Panel number
        worksheet.write(row, 2, str(rev), name_format)        # Revision
        worksheet.write(row, 3, str(date), name_format)       # Issue date
        worksheet.write(row, 4, str(panel_num + '. ' + dname), name_format)  # Drawing name
        worksheet.write(row, 5, str(new_name), name_format)   # New File Name
        
        
        row += 1

####### If folder B #######

    if os.path.isdir(path_A):
        panel_num_B = panel_num + '_' + main_folder[count]
        # panel_num_B = panel_num + '_' + str(count_B)

        for fname_B in os.listdir(path_A):

            path_B = os.path.join(path_A, fname_B)
            if os.path.isfile(path_B):

                # Copying the file to output directory:
                dest_B = os.path.join(dest_folder, fname_B)  # creating filename in output folder 
                shutil.copyfile(path_B, dest_B)

                # Constructing new file name with numbering and savin ít to excel:
                new_name = panel_num_B + '_' + fname_B
                destination = os.path.join(dest_folder , new_name)

                # Renaming the file
                os.rename(dest_B, destination)

                # Extracting rev, date, drawing.no from File_collective.xlsx
                # ofname = Old File name,  rev = revision,  dname = Drawing name,  date = Issue date
                for index in range(len(df)):
                    if df_ofname[index] == fname_B:
                        dname = df_dname[index]
                        rev = str(df_rev[index])
                        date = str(df_date[index])
                        break
                    elif df_ofname[index].startswith("B 4.13.2.2 Transport, Idriftsætning og Vedligeholdelse") == True:
                        dname = "Løgstrup. B 4.13.2.2 Transport, Idriftsætning og Vedligeholdelse"
                        rev = "2"
                        date = "28-06-2017"
                        break
                    else:
                        dname = ""
                        rev = ""
                        date = ""

                # Adding information to Excel:
                worksheet.write(row, 0, str(fname_B), name_format)    # Old File Name
                worksheet.write(row, 1, str(panel_num), name_format)  # Panel number
                worksheet.write(row, 2, str(rev), name_format)        # Revision
                worksheet.write(row, 3, str(date), name_format)       # Issue date
                worksheet.write(row, 4, str(panel_num + '. ' + main_folder[count] + '. ' + dname), name_format)  # Drawing name
                worksheet.write(row, 5, str(new_name), name_format)   # New File Name

                row += 1

####### If folder C #######

            if os.path.isdir(path_B):

                count_D = 1    #folder count for level D

                for fname_C in os.listdir(path_B):
                    panel_num_C = panel_num_B + '_' + str(count_C)

                    path_C = os.path.join(path_B, fname_C)
                    if os.path.isfile(path_C):

                        # Copying the file to output directory:
                        dest_C = os.path.join(dest_folder, fname_C)  # creating filename in output folder 
                        shutil.copyfile(path_C, dest_C)
                        
                        # Constructing old file name and saving it to excel:
                        worksheet.write(row, 0, str(fname_C), name_format)

                        # Constructing new file name with numbering and savin ít to excel:
                        new_name = panel_num_C + '_' + fname_C
                        destination = os.path.join(dest_folder , new_name)
                        worksheet.write(row, 1, str(new_name), name_format)

                        # Renaming the file
                        os.rename(dest_C, destination)

                        # Extracting rev, date, drawing.no from File_collective.xlsx
                        # ofname = Old File name,  rev = revision,  dname = Drawing name,  date = Issue date
                        for index in range(len(df)):
                            if df_ofname[index] == fname_C:
                                dname = df_dname[index]
                                rev = str(df_rev[index])
                                date = str(df_date[index])
                                break
                            elif df_ofname[index].startswith("B 4.13.2.2 Transport, Idriftsætning og Vedligeholdelse") == True:
                                dname = "Løgstrup. B 4.13.2.2 Transport, Idriftsætning og Vedligeholdelse"
                                rev = "2"
                                date = "28-06-2017"
                                break
                            else:
                                dname = ""
                                rev = ""
                                date = ""

                        # Adding information to Excel:
                        worksheet.write(row, 0, str(fname_C), name_format)    # Old File Name
                        worksheet.write(row, 1, str(panel_num), name_format)  # Panel number
                        worksheet.write(row, 2, str(rev), name_format)        # Revision
                        worksheet.write(row, 3, str(date), name_format)       # Issue date
                        worksheet.write(row, 4, str(panel_num + '. ' + main_folder[count] + '. ' + fname_B + '. ' + dname), name_format)  # Drawing name
                        worksheet.write(row, 5, str(new_name), name_format)   # New File Name

                        row += 1

# ####### If folder D #######
                    if os.path.isdir(path_C):

                        for fname_D in os.listdir(path_C):
                            panel_num_D = panel_num_C + '.' + str(count_D)

                            path_D = os.path.join(path_C, fname_D)
                            if os.path.isfile(path_D):

                                # Copying the file to output directory:
                                dest_D = os.path.join(dest_folder, fname_D)  # creating filename in output folder 
                                shutil.copyfile(path_D, dest_D)

                                # Constructing new file name with numbering and savin ít to excel:
                                new_name = panel_num_D + '_' + fname_D
                                destination = os.path.join(dest_folder , new_name)

                                # Renaming the file
                                os.rename(dest_D, destination)

                                # Extracting rev, date, drawing.no from File_collective.xlsx
                                for index in range(len(df)):
                                    if df_ofname[index] == fname_D:
                                        dname = df_dname[index]
                                        rev = str(df_rev[index])
                                        date = str(df_date[index])
                                        break
                                    else:
                                        dname = ""
                                        rev = ""
                                        date = ""

                                # Adding information to Excel:
                                worksheet.write(row, 0, str(fname_D), name_format)    # Old File Name
                                worksheet.write(row, 1, str(panel_num), name_format)  # Panel number
                                worksheet.write(row, 2, str(rev), name_format)        # Revision
                                worksheet.write(row, 3, str(date), name_format)       # Issue date
                                worksheet.write(row, 4, str(panel_num + '. ' + main_folder[count] + '. ' + fname_B + '. ' + fname_C + '. ' + dname), name_format)  # Drawing name
                                worksheet.write(row, 5, str(new_name), name_format)   # New File Name

                                row += 1

                            if os.path.isdir(path_D):
                                print("OBS!! Level E folder structure")


                        count_D += 1
                        
                count_C += 1
        # count_B += 1
        count += 1



workbook.close()
