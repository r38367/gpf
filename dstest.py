
# Import Module
from tkinter import *
 

nVer = '4.0'

def _take_button_pressed():
 
    res = "take clicked" 
    lbl.configure(text = res)

def _open_button_pressed():
    res = "open clicked" 
    lbl.configure(text = res)

def _input_pressed(event):
    res = txt.get()
    lbl.configure(text = res)

def on_window_resize(event):
    width = event.width
    height = event.height
    print(f"Window resized to {width}x{height}")

# create window
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
btn = Button(root, text = "Take" ,
             fg = "black", width=6, command=_take_button_pressed)
btn.grid(column=1, row=0, padx=5)

#
# Create button Open
#
btn = Button(root, text = "Open" , width=6,
             fg = "black", command=_open_button_pressed)
btn.grid(column=2, row=0, padx=5)

# Execute Tkinter
root.mainloop()
