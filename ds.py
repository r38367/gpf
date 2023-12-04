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

nVer = '4.0'
test = 0
policy = 0
wordname = ""
jpg = ""

# ==============================================================================================================
#                   Input entered 
# ==============================================================================================================
   
def _input_pressed(event):
    global wordname 
    
    res = txt.get().strip()
    wordname = makewordname( res )
    
    lbl.configure(text = wordname)


# ==============================================================================================================
#                   Take button pressed 
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
        _long = root.after(1000,_change_label,event)
        

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
        res = 'Saved to ' + wordname [-20:] 
    else:
        res = 'Saved to ' + jpg  [-20:]

    lbl.configure(text = res)


# ==============================================================================================================
#                   Open button pressed 
# ==============================================================================================================

def _open_button_pressed():
    res = "No file to show" 
    
    if jpg == "":
        # Show an information message box
        mb.showinfo(title="Message", message=res)
    else:
        os.system('"'+jpg+'"')



# ==============================================================================================================
# Create GUI
# ==============================================================================================================

root = Tk()
 
#  windowSet  title 
root.title("Test DinSide - v." + nVer )

# Set geometry(widthxheight)
root.geometry('400x70')

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
txt = Entry(root, width=40)
txt.grid(column =0, row =0, padx=5, pady=5)
txt.bind('<Return>', _input_pressed)
 
#
# Create label 
#
lbl = Label(root, text = "Test:, Policy: ", anchor="w", justify="left")
lbl.grid(column =0, row =1,sticky=W, padx=5)

#
# Create button Take
#
btn1 = Button(root, text = "Take" ,
             fg = "black", width=6)
btn1.grid(column=1, row=0, padx=5)
btn1.bind('<Button-1>', _take_button_pressed)
btn1.bind('<ButtonRelease-1>', _take_button_released)
#
# Create button Open
#
btn2 = Button(root, text = "Open" , width=6,
             fg = "black", command=_open_button_pressed)
btn2.grid(column=2, row=0, padx=5)

# Execute Tkinter
root.mainloop()

