
from easygui import *

# Title message
text = "Input document details"

# window title
title = "Input data"
 
# list of multiple inputs
input_list = ["Tavlenummer", "T.I.V.docx number"]
 
# list of default text
default_list = ["SPAx.x", "xx"]
 
# creating a integer box
output = multenterbox(text, title, input_list, default_list)

# title for the message box
title = "Message Box"


# creating a message
message = "Entered details are:   " + input_list[0] + ": " + output[0] + "  ,  " + input_list[1] + ": " + output[1]

# text of the Ok button
choices = ["Continue" , "Cancel"]

# Creating a continue cancel box
action = ccbox(message, title, choices)

if action == False:
    exit()






# creating a message box
# msg = msgbox(message, title, ok_btn_txt)





