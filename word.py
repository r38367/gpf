from docx import Document
from docx.shared import Mm
  

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

    document.save(file)

def makewordname( input ):
    if input == "" :
         return ""
    else: 
        return 'Anton' + '.docx'
"""
from screen import takescreenshoot
from timestamp import timestamp 

text = timestamp()
jpg = text.replace(':','').replace(' ','_') + "_screen.jpg"
takescreenshoot( jpg )
savetoword( 'demo.docx', text, jpg )
"""