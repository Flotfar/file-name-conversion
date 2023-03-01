import numpy as np
import pandas as pd

# # Importing the excel file, sheet.no 5 "File_load"
# df = pd.read_excel(r'C:\Users\jbmk\OneDrive - Ramboll\Desktop\SoB - Doc\File Renaming PY\File_collective.xlsx', sheet_name='File_load')


# # Refferencing columns to variables:
# df_ofname = df['Gammelt Filnavn']
# df_rev = df['Revision']
# df_date = df['Udgavedato'].apply(lambda x: x.strftime('%Y-%m-%d'))   #Clearing out timestamps from the Date
# df_dname = df['Tegningsnavn']

# F_name = 'iRCP_A9E21180.pdf'

# for index in range(len(df)):
#     if df_ofname[index] == F_name:
#         dname = df_dname[index]
#         rev = str(df_rev[index])
#         date = str(df_date[index])
#         break
# else:
#     dname = ""
#     rev = ""
#     date = ""

# print("name = " + dname)
# print("rev = " + rev)
# print("date = " + date)

df_2 = pd.read_excel(r'C:\Users\jbmk\OneDrive - Ramboll\Desktop\SoB - Doc\File Renaming PY\File_naming.xlsx', sheet_name='Renaming')

df_2_long_name = df_2['Original']
df_2_short_name = df_2['Omskrivning']

print(df_2_long_name)
print(df_2_short_name)
