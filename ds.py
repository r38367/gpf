

# History
#
# v1.0
#   03/12/23 add #1 open/edit jpg button  
#   04/12/23 add #5 clear button
#   05/12/23 add #4 open/edit word button
#   05/12/23 add #11 clear with double click on label
#
# v2.0
#   09/12/23 add #6 add file list in inout field
#

nVer = '2.0'

# Import Module
from tkinter import *
import tkinter.messagebox as mb
import os, sys
import time

# ==============================================================================================================
# create a timestamp in form dd.mm.yy hh:mm:ss 
# ==============================================================================================================

import datetime

def timestamp(): 
    # Get the current date and time
    now = datetime.datetime.now()

    return now.strftime("%d.%m.%y %H:%M:%S")

# ==============================================================================================================
# take a screenshot 
# ==============================================================================================================


from PIL import ImageGrab

def takescreenshoot( file ):
    # Capture the entire screen
    screenshot = ImageGrab.grab()

    
    # Save the screenshot to a file
    screenshot.save( file )

    # Close the screenshot
    screenshot.close()

def makescreenshotname( t, text="screen" ):
    subfolder=".\\screens" # if run not from folder - os.path.dirname(sys.argv[0]) +
    if not os.path.exists(subfolder):
        os.mkdir(subfolder)

    return   subfolder + '\\' + t.replace(':','').replace(' ','_') + "_" + text +".png"
 

# ==============================================================================================================
# working with word document 
# ==============================================================================================================

from docx import Document
from docx.shared import Mm
import re  

def get_text_width(document):
    """
    Returns the text width in mm.
    """
    section = document.sections[0]
    return (section.page_width - section.left_margin - section.right_margin) / 36000

def savetoword( file, text, jpg):
    
    try:
        # Load the existing document
        document = Document( file )
    except:
        # Create a new document if it does not exist
        document = Document()

    #document = Document('demo.docx')
    document.sections[0].left_margin = Mm(25)
    document.sections[0].right_margin = Mm(25)

    document.add_paragraph( text )
    document.paragraphs[-1].add_run().add_picture(jpg, Mm(get_text_width(document)) )
                        #width=Mm(document.sections[0].page_width - document.sections[0].left_margin - document.sections[0].right_margin)) 

    try: 
            document.save(file)
    except:
            mb.showerror(title="Error", message="Can't save " + file + ".\nClose if open.")

def makewordname( input ):
    if input == "" :
         return ""
    else: 
         
        # Find the first 4-digit number
        four_digit = re.findall(r'\b\d{4}\b', input)
        if not four_digit:
            testnum = ""
        else:
            testnum = four_digit[0]
        # Find the first 7-digit number
        seven_digit = re.findall(r'\b\d00\d{4}\b', input )
        if not seven_digit:
            policynum = ""
        else:
            policynum = seven_digit[0]
        
        #find name for file 
        #name can be anything not digits
        # for example:
        # 1234 6000313 this is a file name =this is a file name
        # 1234 filename2 6000313 =filename2
        # my file 6000313 2546 =my file
        # 2456 6000313 G02056 =G02056
        # 3333 anton 6000777 delete =
        #
        non_digit = re.findall(r'\b([^\d\s].*?)(?:\s+\d+)*$', input )
        if not non_digit:
            extratext = os.getlogin()
            ext = ""
        else:
            extratext = non_digit[0]
            ext = non_digit[0]
        
        # Concatenate the two numbers
        ret = testnum + policynum + ext  
        if ret == "":
            return ""
        else:
            return testnum +'_'+ extratext +'_' + policynum + '.docx'



# ==============================================================================================================
#  Call backs 
# ==============================================================================================================

test = 0
policy = 0
wordname = ""
jpg = ""

# ==============================================================================================================
#                   Input 
# pressed  
# ==============================================================================================================
   
def _input_pressed(event):
    global wordname 
    
    # hide listbox
    listbox.grid_remove()
    # get text from input 
    res = txt.get().strip()
    # make file name from input 
    wordname = makewordname( res )
    # set label to file name 
    lbl.configure(text = wordname)


# ==============================================================================================================
#                   Take button 
# 
# pressed/released/changed
# ==============================================================================================================
  
_pressed = 0  # 0-not pressed, 1-short, 2-long

def _change_label(event):
    global _pressed
    if _pressed == 1:
        _pressed = 2
        btn1.configure( text='Edit')

     
def _take_button_pressed( event ):
    global _pressed
    global _long

    if _pressed == 0 :
        _pressed = 1
        _long = root.after(700,_change_label,event)
        

def _take_button_released( event ):

    global _pressed
    global _long

    #if short press
    if _pressed == 1:
        if _long:
            root.after_cancel(_long)    # cancel long event

    # take screenshot
    global wordname 
    global jpg

    _input_pressed( 0 )
    t = timestamp()
    jpg = makescreenshotname( t, txt.get().strip() )
    takescreenshoot( jpg )


    # edit picture
    if _pressed == 2:
        os.system('"'+jpg+'"')
       
    _pressed = 0
    btn1.configure( text='Take')

    # save to word if needed 
    if not wordname == "":
        savetoword( wordname, t, jpg )
        res = 'Saved to ' + wordname  
    else:
        res = 'Saved to ' + jpg 

    lbl.configure(text = res)


# ==============================================================================================================
#                   Open button
# 
#            pressed/released/changed
# ==============================================================================================================

_pressed_open = 0  # 0-not pressed, 1-short, 2-long

def _open_change_label(event):
    global _pressed_open
    if _pressed_open == 1:
        _pressed_open = 2
        btn2.configure( text='Word')

     
def _open_button_pressed( event ):
    global _pressed_open
    global _long

    _input_pressed( 0 )
    
    if _pressed_open == 0 :
        _pressed_open = 1
        _long = root.after(700,_open_change_label,event)
        

def _open_button_released( event ):

    global _pressed_open
    global _long

    #if short press
    if _pressed_open == 1:
        if _long:
            root.after_cancel(_long)    # cancel long event

    # open file
    global wordname 
    global jpg

    # open picture/word
    if _pressed_open == 2:
        
        if os.path.isfile(wordname):
            os.system('"'+ wordname +'"')
        else:
            # Show an information message box
            mb.showinfo(title="Message", message="No word file to show")

    else:
         
        if jpg == "":
            # Show an information message box
            mb.showinfo(title="Message", message="No image file to show")
        else:
            os.system('"'+jpg+'"')

    _pressed_open = 0
    btn2.configure( text='Open')
    
   

# ==============================================================================================================
#                  Clear button 
# pressed 
# ==============================================================================================================

def _clear_button_pressed(event=None):
    
    txt.delete(0, END)
    lbl.configure(text = "")
    listbox.grid_remove()

# ==============================================================================================================
#                   Listbox - update
#
# update after input text in input field  
# ==============================================================================================================

def update_listbox (event=None):
    
    if event.keysym == "Return":
        # Move the focus to the listbox
        listbox.grid_remove()
        return
    
    # Clear the listbox
    listbox.delete (0, END)
    # Get the input text
    input_text = txt.get ().strip()
    #check if string is empty
    if not input_text:
        return 
    # Initialize a counter for the number of matching files
    count = 0
    # Loop through the files in the current folder
    for file in os.listdir ("."):
        # Check if the file name contains the input text
        if all(file.upper().find(str) >= 0  for str in input_text.upper().split()):
            # Insert the file name into the listbox
            listbox.insert (END, file)
            # Increment the counter
            count += 1
    # Resize the listbox height according to the counter
    listbox.config (height=min(count,5)) 
    if count < 1:
        listbox.grid_remove()
    else:
        listbox.grid(column=0, row=1)
    
        
# ==============================================================================================================
#                   Listbox - return from listbox
# 
# Create a function to update the entry when press Return in listbox
# ==============================================================================================================
def update_entry (event):
    
    # Check if the event is triggered by pressing ENTER
    if event.keysym == "Return":
        # Get the selected file name from the listbox
        file_name = listbox.get (listbox.curselection ())
        # hide listbox
        listbox.grid_remove()
        # Set the input text to the file name
        txt.delete(0, END)
        txt.insert (0, file_name.replace('_',' ').rsplit('.',1)[0])
        # Update the listbox
        #update_listbox ()
        
        # Hide the toplevel window 
        # Changing position of cursor to end 
        txt.icursor(END)
        # set focus to entry
        txt.focus()
        #top.withdraw ()
        _input_pressed(None)
        

# ==============================================================================================================
#                   Listbox - show file name 
# 
# Show the active entry from listbox in the input  
# ==============================================================================================================
def show_file_name (event):
    # Get the selected file name from the listbox
    file_name = listbox.get (listbox.curselection ())
    # Set the input text to the file name
    txt.delete(0, END)
    txt.insert (0, file_name.replace('_',' ').rsplit('.',1)[0])

# ==============================================================================================================
#                   Listbox - move focus 
# 
# Create a function to move the focus to the listbox
# ==============================================================================================================
def move_focus (event):
    # Check if the event is triggered by pressing DOWN
    if event.keysym == "Down":
        # Move the focus to the listbox
        listbox.focus ()
        # Select the first item in the listbox
        listbox.selection_set (0)
        
    if event.num == 1:
        index = listbox.nearest(event.y)
        value = listbox.get(index)
        listbox.selection_set (index)
        

# ==============================================================================================================
# Create GUI
# ==============================================================================================================

root = Tk()
 
#  windowSet  title 
root.title("Test DinSide - v." + nVer )

# Set geometry(widthxheight)
root.geometry('') 

# Set margins inside window
root['padx']=10
root['pady']=5

# make windows resizable 
root.resizable(False, False)
# handle resize event to re-allocate buttons
#root.bind("<Configure>", on_window_resize)

# Set always on top
root.attributes('-topmost',True)

# Set window bg color
#root.configure(bg='lightgray')

#
# widgets:
#      0                       1      2
# 0 [input text           ] [Take] [Open]
# 1 [label text           ]
#


#
#  create input widget
#

# Create an entry widget
txt = Entry(root, width=40)
# place top left
txt.grid(column =0, row =0, padx=5, pady=5)
# Bind to Return key
txt.bind('<Return>', _input_pressed)
# Bind the entry widget to the update function
txt.bind ("<KeyRelease>", update_listbox )
# Bind the entry widget to the move focus function
txt.bind ("<Down>", move_focus)
# Bind the root window to the show and hide functions
txt.bind ("<FocusIn>", update_listbox)

# Pack the entry widget
#txt.pack ()

#
# Create label 
#
lbl = Label(root, text = "", anchor="w", justify="left")
lbl.grid(column =0, row =1, sticky=W, padx=5)
#lbl.bind('<Double-1>', _clear_button_pressed) # Useless

#
# Create button Take
#
btn1 = Button(root, text = "Take" , width=6, fg = "black")
btn1.grid(column=1, row=0, padx=5)
btn1.bind('<Button-1>', _take_button_pressed)
btn1.bind('<ButtonRelease-1>', _take_button_released)

#
# Create button Open
#
btn2 = Button(root, text = "Open" , width=6, fg = "black")
btn2.grid(column=2, row=0, padx=5)
btn2.bind('<Button-1>', _open_button_pressed)
btn2.bind('<ButtonRelease-1>', _open_button_released)
#
# Create button Clear
#
btn3 = Button(root, text = "Clear" , width=6, fg = "black", command=_clear_button_pressed)
btn3.grid(column=1, row=1, columnspan=2, padx=5, sticky=N)

#
# Create listbox
#
listbox = Listbox (root, width=40, height=5)
# Pack the listbox widget
listbox.grid(column=0, row=1, sticky=NW, padx=5, pady=5)
# Bind the listbox widget to the update function
listbox.bind ("<Return>", update_entry)
# Bind the listbox widget to the update function
listbox.bind ("<Button-1>", move_focus)
# Bind the listbox widget to the show file name function when item is selected 
listbox.bind ("<<ListboxSelect>>", show_file_name)
# Bind the listbox widget to quit when item is selected 
listbox.bind ("<FocusOut>", lambda e: listbox.grid_remove ())
            
# Update the listbox initially
listbox.grid_remove ()

# Start the main loop
root.mainloop ()
