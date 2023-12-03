from PIL import ImageGrab
from timestamp import timestamp

def takescreenshoot( file ):
    # Capture the entire screen
    screenshot = ImageGrab.grab()

    # Save the screenshot to a file
    screenshot.save( file )

    # Close the screenshot
    screenshot.close()

def makescreenshotname( t, text="screen" ):
    return  t.replace(':','').replace(' ','_') + "_" + text +".png"

"""    
name = makescreenshotname( timestamp(), "my")
print (name )
takescreenshoot( name )

name = makescreenshotname( timestamp() )
print (name )
takescreenshoot( name )
"""