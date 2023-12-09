# Import tkinter module
import tkinter as tk
# Import os module to list files
import os, time

# Create a root window
root = tk.Tk ()
# Set the window title
root.title ("Listbox with Entry")
# Set the window size
root.geometry ("300x100")

# Create a variable to store the input text
input_var = tk.StringVar ()

# Create a function to update the listbox
def update_listbox ():
    # Clear the listbox
    listbox.delete (0, tk.END)
    # Get the input text
    input_text = input_var.get ().strip()
    # Initialize a counter for the number of matching files
    count = 0
    # Loop through the files in the current folder
    for file in os.listdir ("."):
        # Check if the file name contains the input text
        if all(file.upper().find(str) >= 0  for str in input_text.upper().split()):
            # Insert the file name into the listbox
            listbox.insert (tk.END, file)
            # Increment the counter
            count += 1
    # Resize the listbox height according to the counter
    if count < 1:
        listbox.pack_forget()
    else:
        listbox.pack()
    
    listbox.config (height=count,width=100)
    
# Create a function to update the entry when press Return in listbox
def update_entry (event):
    
    # Check if the event is triggered by pressing ENTER
    if event.keysym == "Return":
        # Get the selected file name from the listbox
        file_name = listbox.get (listbox.curselection ())
        # Set the input text to the file name
        input_var.set (file_name)
        # Update the listbox
        update_listbox ()
        # hide listbox
        listbox.pack_forget()
        # Hide the toplevel window 
        # Changing position of cursor to end 
        entry.icursor(tk.END)
        # set focus to entry
        entry.focus()
        top.withdraw ()
        time.sleep(1)
        
# Create a function to show the file name in the entry
def show_file_name (event):
    # Get the selected file name from the listbox
    file_name = listbox.get (listbox.curselection ())
    # Set the input text to the file name
    input_var.set (file_name)

# Create a function to move the focus to the listbox
def move_focus (event):
    # Check if the event is triggered by pressing DOWN
    if event.keysym == "Down":
        # Move the focus to the listbox
        listbox.focus ()
        # Select the first item in the listbox
        listbox.selection_set (0)

# Create an entry widget
entry = tk.Entry (root, textvariable=input_var, width=40)
# Bind the entry widget to the update function
entry.bind ("<KeyRelease>", lambda e: update_listbox ())
# Bind the entry widget to the move focus function
entry.bind ("<Down>", move_focus)
# Pack the entry widget
entry.pack ()

# Create a toplevel window to display the listbox
top = tk.Toplevel (root)
# Remove the title bar of the toplevel window
top.overrideredirect (True)
# Place the toplevel window below the entry widget
top.geometry (f"100x100")
# Make the toplevel window transparent
top.wm_attributes ("-alpha", 0.5)
# Make the toplevel window resizable
top.resizable (False, False)
# Make the toplevel window follow the root window
top.transient (root)

top.configure(bg="light green")
# Hide the toplevel window initially
top.withdraw ()

# Create a listbox widget
listbox = tk.Listbox (top)
# Bind the listbox widget to the update function
listbox.bind ("<Return>", update_entry)
# Bind the listbox widget to the show file name function when item is selected 
listbox.bind ("<<ListboxSelect>>", show_file_name)
# Pack the listbox widget
listbox.pack ()

# Update the listbox initially
update_listbox ()



# Create a function to show the toplevel window
def show_toplevel ():
    # Show the toplevel window
    top.deiconify ()
    # Update the toplevel window size and position
    top.update ()
    # Adjust the toplevel window width to match the listbox width
    top.geometry ("%dx%d+%d+%d" % (entry.winfo_width(), 100, root.winfo_rootx()+entry.winfo_x(), root.winfo_rooty()+entry.winfo_y()+entry.winfo_height()+1))

# Create a function to hide the toplevel window
def hide_toplevel ():
    # Hide the toplevel window
    top.withdraw ()

# Bind the root window to the show and hide functions
entry.bind ("<FocusIn>", lambda e: show_toplevel ())
listbox.bind ("<FocusOut>", lambda e: hide_toplevel ())

show_toplevel()

# Start the main loop
root.mainloop ()
