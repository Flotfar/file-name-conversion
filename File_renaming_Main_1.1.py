# this code is running the os-module which is a 
# part of the standard library, or stdlib, within Python 3.


#########################################################
#### Rules of the script:                            ####
#########################################################

# All files belonging to vedligehold should be placed within that folder in input
# All files belonging to 
# All lose files should be placed in the open invironment of the input folder

#########################################################
#########################################################
#### Future opdates:                                 ####
#########################################################

# Make a loop to capture all "mother-directories" and list them as it is for 'Main_folder' variable
# Squeeze all folder-loop-levels together with array input instead.
# draw information from "Tavle standardtekster" (input as panda) as output in excel

#########################################################



import os
import shutil    
import xlsxwriter
import pandas as pd
import re


folder = r'C:\Users\jbmk\OneDrive - Ramboll\Desktop\SoB - Doc\File Renaming PY\Data_input'   # file path, rapped with 'r'
dest_folder = r'C:\Users\jbmk\OneDrive - Ramboll\Desktop\SoB - Doc\File Renaming PY\Data_output'  # destination folder

from Input_menu import *


panel_num = output[0]
tiv = output[1]

main_folder = ["Vedligeh" , "Indstl.Opstil"]    #Main folder naming


###############################################################################
#### Extracting general File information from Collective.xlsx (File_load): ####
###############################################################################

# Importing the excel file, sheet.no 5 "File_load" and "Tavle_nr_Standardtekster"
df = pd.read_excel(r'C:\Users\jbmk\OneDrive - Ramboll\Desktop\SoB - Doc\File Renaming PY\File_naming.xlsx', sheet_name='File_load')
df_2 = pd.read_excel(r'C:\Users\jbmk\OneDrive - Ramboll\Desktop\SoB - Doc\File Renaming PY\File_naming.xlsx', sheet_name='Renaming')


# Refferencing columns to variables (File_load):
df_ofname = df['Gammelt Filnavn']
df_rev = df['Revision']
df_date = df['Udgavedato'].apply(lambda x: x.strftime('%d-%m-%Y'))   #Clearing out timestamps from the Date
df_dname = df['Tegningsnavn']

# Refferencing columns to variables (Renaming):
df_2_long_name = df_2['Original']
df_2_short_name = df_2['Omskrivning']

# Creating empty arrays
Old_fname = []
Panel = []
Rev_date = []
Ed_date = []
Draw_name = []
New_fname = []



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



#########################################################
#### Dname, rev and date extraction function:        ####
#########################################################

def drawing_namer(fname):
    for index in range(len(df)):
            if df_ofname[index] == fname:
                dname = df_dname[index]
                rev = str(df_rev[index])
                date = str(df_date[index])
                break
            #Driftsætning og vedligeholds doc:
            elif fname.startswith("B 4.13.2.2 Transport, Idriftsætning og Vedligeholdelse") == True:
                dname = "Løgstrup. B 4.13.2.2 Transport, Idriftsætning og Vedligeholdelse"
                rev = "2"
                date = "28-06-2017"
                break
            #03-833xx filer:
            elif "diagram" in fname.lower():
                numbers = re.findall(r"\d{4}\b|\d{2}\b", fname)
                dnum = ''.join(numbers)
                dname = "Løgstrup. " + dnum + ". Electrical Diagram"
                rev = "0"
                date = ""
                break
            elif "layout" in fname.lower():
                numbers = re.findall(r"\d{4}\b|\d{2}", fname)
                dnum = ''.join(numbers[2:5])
                dname = "Løgstrup. " + dnum + ". Layout"
                rev = ""
                date = ""
                break
            elif "test report" in fname.lower():
                numbers = re.findall(r"\d{8}\b", fname)
                dnum = ''.join(numbers)
                dname = "Løgstrup. " + dnum + ". Test report"
                rev = "0"
                date = ""
                break
            elif "settingsreport" in fname.lower():
                numbers = re.findall(r"\d{8}\b", fname)
                dnum = ''.join(numbers)
                dname = "ABB. " + "E-Hub 2.0. Settings Report. " + dnum
                rev = "0"
                date = "17-05-2022"
                break
            else:
                dname = ""
                rev = ""
                date = ""
    
    return dname, rev, date
    

#########################################################
#### short filename creation function:        ####
#########################################################

def file_namer(fname):
    for index in range(len(df_2)):
            if df_2_long_name[index] == fname:
                new_name = panel_num + '_' + df_2_short_name[index]
                break
            #Driftsætning og vedligeholds doc:
            elif fname.startswith("B 4.13.2.2 Transport, Idriftsætning og Vedligeholdelse") == True:
                new_name = panel_num + '_' + "B 4.13.2.2 Trans.Idrifts.Vedligeh_" + tiv + ".docx"
                break
            #03-833xx filer:
            elif "diagram" in fname.lower():
                numbers = re.findall(r"\d{4}\b|\d{2}\b", fname)
                dnum = ''.join(numbers)
                new_name = panel_num + '_' + tiv + '_' + "Diagram_" + '_' + dnum + ".pdf"
                break
            elif "layout" in fname.lower():
                numbers = re.findall(r"\d{4}\b|\d{2}", fname)
                dnum = ''.join(numbers[2:5])
                new_name = panel_num + '_' + tiv + '_' + "Layout_" + dnum + ".pdf"
                break
            elif "test report" in fname.lower():
                numbers = re.findall(r"\d{8}\b", fname)
                dnum = ''.join(numbers)
                new_name = panel_num + '_' + tiv + '_' + "Test report_" + dnum + ".pdf"
                break
            elif "settingsreport" in fname.lower():
                numbers = re.findall(r"\d{8}\b", fname)
                dnum = ''.join(numbers)
                new_name = panel_num + '_' + "Settingsreport_" + tiv + '_' + "E-hub_" + dnum + ".pdf"
                break
            else:
                new_name = panel_num + '_' + fname_A
    
    return new_name

# count increase by 1 in each iteration
# iterate all files from a directory

####### Folder A #######

for fname_A in os.listdir(folder):

    path_A = os.path.join(folder, fname_A)
    if os.path.isfile(path_A):

        # Copying the file to output directory:
        dest_A = os.path.join(dest_folder, fname_A)  # creating filename in output folder 
        shutil.copyfile(path_A, dest_A)

        # Constructing new file name with numbering + reference to standart file name list (df_2):
        new_name = file_namer(fname_A)
        
        destination = os.path.join(dest_folder, new_name)

        # Renaming the file
        os.rename(dest_A, destination)

        # Extracting rev, date, drawing.no from File_collective.xlsx
        # ofname = Old File name,  rev = revision,  dname = Drawing name,  date = Issue date
        dname, rev, date = drawing_namer(fname_A)

        # Adding information to Array:
        
        Old_fname.append(str(fname_A))
        Panel.append(str(panel_num))
        Rev_date.append(str(rev))
        Ed_date.append(str(date))
        Draw_name.append(str(panel_num + '. ' + dname))
        New_fname.append(str(new_name))
        
        row += 1

####### If folder B #######

    if os.path.isdir(path_A):
        panel_num_B = panel_num + '_' + main_folder[count]
        # panel_num_B = panel_num + '_' + str(count_B)

        for fname_B in os.listdir(path_A):

            path_B = os.path.join(path_A, fname_B)
            if os.path.isfile(path_B):

                # Copying the file to output directory:
                dest_B = os.path.join(dest_folder, fname_B)  # creating filepath with output folder 
                shutil.copyfile(path_B, dest_B)

                # Constructing new file name with numbering + reference to standart file name list (df_2):
                new_name = file_namer(fname_B)
                
                destination = os.path.join(dest_folder , new_name)

                # Renaming the file
                os.rename(dest_B, destination)

                # Extracting rev, date, drawing.no from File_naming.xlsx
                # ofname = Old File name,  rev = revision,  dname = Drawing name,  date = Issue date
                dname, rev, date = drawing_namer(fname_B)

                # Adding information to Array
                Old_fname.append(str(fname_B))
                Panel.append(str(panel_num_B))
                Rev_date.append(str(rev))
                Ed_date.append(str(date))
                Draw_name.append(str(panel_num + '. ' + main_folder[count] + '. ' + dname))
                New_fname.append(str(new_name))

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

                        # Constructing new file name with numbering + reference to standart file name list (df_2):
                        new_name = file_namer(fname_C)

                        destination = os.path.join(dest_folder , new_name)

                        # Renaming the file
                        os.rename(dest_C, destination)

                        # Extracting rev, date, drawing.no from File_collective.xlsx
                        # ofname = Old File name,  rev = revision,  dname = Drawing name,  date = Issue date
                        dname, rev, date = drawing_namer(fname_C)

                        # Adding information to Array
                        Old_fname.append(str(fname_C))
                        Panel.append(str(panel_num_C))
                        Rev_date.append(str(rev))
                        Ed_date.append(str(date))
                        Draw_name.append(str(panel_num + '. ' + main_folder[count] + '. ' + fname_B + '. ' + dname))
                        New_fname.append(str(new_name))
                        
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

                                # Constructing new file name with numbering + reference to standart file name list (df_2):
                                new_name = file_namer(fname_D)

                                destination = os.path.join(dest_folder , new_name)

                                # Renaming the file
                                os.rename(dest_D, destination)

                                # Extracting rev, date, drawing.no from File_collective.xlsx
                                dname, rev, date = drawing_namer(fname_D)

                                # Adding information to Array
                                Old_fname.append(str(fname_D))
                                Panel.append(str(panel_num_D))
                                Rev_date.append(str(rev))
                                Ed_date.append(str(date))
                                Draw_name.append(str(panel_num + '. ' + main_folder[count] + '. ' + fname_B + '. ' + fname_C + '. ' + dname))
                                New_fname.append(str(new_name))

                                row += 1

                            if os.path.isdir(path_D):
                                print("OBS!! Level E folder structure")


                        count_D += 1
                        
                count_C += 1
        # count_B += 1
        count += 1


df_3 = pd.read_excel(r'C:\Users\jbmk\OneDrive - Ramboll\Desktop\SoB - Doc\Tavle_nr_Standardtekster.xls')

df_place = df_3['Placering']
df_lok_des = df_3['Lokation beskrivelse']
df_lok_code = df_3['Lokationskode']
df_ass_des = df_3['Asset beskrivelse']
df_fys_name = df_3['Fysisk navn i marken']
df_lok = df_3['Lokation']


#########################################################
#### Creating Excel file for exportation of results: ####
#########################################################

workbook = xlsxwriter.Workbook('File_info.xlsx')
worksheet = workbook.add_worksheet('Names')
worksheet_2 = workbook.add_worksheet('Load_file')

#formats
title_format = workbook.add_format({'bold': True, 'font_size': '14', 'border': False})
name_format = workbook.add_format({'font_size': '11', 'border': False})

worksheet.write(0, 0, "Gammelt filnavn:" , title_format)
worksheet.write(0, 1, "Tegningsnummer:" , title_format)
worksheet.write(0, 2, "Revision:" , title_format)
worksheet.write(0, 3, "Udgavedato:" , title_format)
worksheet.write(0, 4, "Tegningsnavn:" , title_format)
worksheet.write(0, 5, "Filnavn:" , title_format)

worksheet_2.write(0, 0, "Placering" , title_format)
worksheet_2.set_column('B:B', None, None, {'hidden': True})
worksheet_2.write(0, 2, "Dok.Type" , title_format)
worksheet_2.write(0, 3, "Status" , title_format)
worksheet_2.write(0, 4, "Tegningsnr." , title_format)
worksheet_2.write(0, 5, "Revision" , title_format)
worksheet_2.write(0, 6, "Udgavedato" , title_format)
worksheet_2.write(0, 7, "Tegningsnavn" , title_format)
worksheet_2.write(0, 8, "Bemærkninger" , title_format)
worksheet_2.write(0, 9, "*Fagområde" , title_format)
worksheet_2.set_column('K:K', None, None, {'hidden': True})
worksheet_2.write(0, 11, "Filnavn" , title_format)
worksheet_2.set_column('M:O', None, None, {'hidden': True})
worksheet_2.write(0, 14, "Anlægs/tegn.type" , title_format)
worksheet_2.set_column('P:AA', None, None, {'hidden': True})
worksheet_2.write(0, 27, "Systemnummer" , title_format)
worksheet_2.write(0, 28, "System beskrivelse" , title_format)
worksheet_2.write(0, 29, "Lokations beskrivelse" , title_format)
worksheet_2.write(0, 30, "Lokationskode" , title_format)
worksheet_2.write(0, 31, "Asset beskrivelse" , title_format)
worksheet_2.write(0, 32, "Fysisk navn i marken" , title_format)
worksheet_2.set_column('AH:AH', None, None, {'hidden': True})
worksheet_2.write(0, 34, "Lokation" , title_format)


# Determaining values from "Tavle standardtekster"
for index in range(len(df_3)):
    if df_fys_name[index] == panel_num:
        place = df_place[index]
        lok_des = df_lok_des[index]
        lok_code = df_lok_code[index]
        ass_des = df_ass_des[index]
        fys_name = panel_num
        lok = df_lok[index]
        break
    else:
        continue


for index in range(len(Panel)):

    row = 1+index

    worksheet.write(row, 0, Old_fname[index], name_format)     # Old File Name
    worksheet.write(row, 1, Panel[index], name_format)         # Panel number
    worksheet.write(row, 2, Rev_date[index], name_format)      # Revision
    worksheet.write(row, 3, Ed_date[index], name_format)       # Issue date
    worksheet.write(row, 4, Draw_name[index], name_format)     # Drawing name
    worksheet.write(row, 5, New_fname[index], name_format)     # New File Name

    worksheet_2.write(row, 0, place, name_format)                # placering
    worksheet_2.write(row, 2, 'ANLÆGSDOKUMENTATION', name_format) 
    worksheet_2.write(row, 3, 'gældende', name_format) 
    worksheet_2.write(row, 4, Panel[index], name_format)         # Panel number
    worksheet_2.write(row, 5, Rev_date[index], name_format)      # Revision
    worksheet_2.write(row, 6, Ed_date[index], name_format)       # Issue date
    worksheet_2.write(row, 7, Draw_name[index], name_format)     # Drawing name
    worksheet_2.write(row, 8, 'Løgstrup', name_format)
    worksheet_2.write(row, 9, 'STÆRKSTRØM', name_format)
    worksheet_2.write(row, 11, New_fname[index], name_format)    # New File Name
    worksheet_2.write(row, 14, 'Dokumentation', name_format) 
    worksheet_2.write(row, 27, '152.25.07', name_format) 
    worksheet_2.write(row, 28, 'Lavspændingsanlæg', name_format) 
    worksheet_2.write(row, 29, lok_des, name_format)             # Lokations beskrivelse
    worksheet_2.write(row, 30, lok_code, name_format)            # Lokationskode
    worksheet_2.write(row, 31, ass_des, name_format)             # Asset beskrivelse
    worksheet_2.write(row, 32, fys_name, name_format)            # Fysisk navn i marken
    worksheet_2.write(row, 34, lok, name_format)                 # Lokation


workbook.close()

#from Excel_input import *

