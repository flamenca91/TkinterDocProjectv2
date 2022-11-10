from tkinter import *
from DocExtract import *

root = Tk()                # Initiates form with size and background color
root.title("TARGEST_App")
root.geometry("800x400")
root['background'] = '#afeae6'
#root.configure(bg="")

# Creating the label widgets
tagLabel = Label(root, text="Tag Name").place(x=400, y=60)
fileLabel = Label(root, text="Document Name").place(x=580, y=60)
parentTagLbl = Label(root, text="Parent Tag").place(x=400, y=200)
childTagsLbl = Label(root, text="Child Tags").place(x=580, y=200)
pathLabel = Label(root, text="File Path").place(x=30, y=15)

# Creates the textbox widgets
txtTag = Entry(root, width=10, borderwidth=5)
txtTag.place(x=393, y=90)
#e.insert(0, "Tag: ")
txtDoc = Entry(root, width=40, borderwidth=5)
txtDoc.place(x=500, y = 90)
txtParent = Entry(root, width=10, borderwidth=5)
txtParent.place(x=393, y=230)
txtChild = Entry(root, width=40, borderwidth=5)
txtChild.place(x=500, y = 230)
txtDocPath = Entry(root, width=50, borderwidth=5)
txtDocPath.place(x=30, y=40)

#tagEntry = txtTag.get()
docFile = {}
docRelation = {}

def tagDocClick():              # Creates the docFile Dictionary (tags with corresponding file names)
    #myLabel2 = Label(root, text='The button was clicked')
    tagEntry = txtTag.get()
    docEntry = txtDoc.get()
    docFile[tagEntry] = docEntry
    # print(docFile)
    return docFile

def tagParentChildClick():     # Creates the docRelation Dictionary (parent with corresponding children)
    parentTag = txtParent.get()
    childTag = txtChild.get()
    childTag = childTag.split(',')
    childTag = tuple(childTag)
    docRelation[parentTag] = childTag
    print(docRelation)
    return(docRelation)

def GetParentTagsClick():      # Retrieves and runs the GetParentTags() function from the DocExtract.py file
    runner2 = paragraph.add_run("\n\nParent tag/tags\n\n")
    runner2.bold = True                              #make it bold
    GetParentTags()
    report3.save('report3.docx')

def GetChildTagsClick():       # Retrieves and runs the GetChildTags() function from the DocExtract.py file
    runner2 = paragraph.add_run("\n\nChild tag/tags\n\n")
    runner2.bold = True
    GetChildTags()
    report3.save('report3.docx')


# Creates the button widgets with corresponding method call commands
tagFileBtn = Button(root, text='Enter', padx=40, pady=10, command=tagDocClick, bg = 'white')
tagFileBtn.place(x=640,y=140)
parentChildBtn = Button(root, text='Enter', padx=40, pady=10, command=tagParentChildClick, bg = 'white')
parentChildBtn.place(x=640, y=280)
createParentBtn = Button(root, text='Extract Parent Tags', padx=50, pady=20, command=GetParentTagsClick,  bg = 'white')
createParentBtn.place(x=30, y=80)
createChildBtn = Button(root, text='Extract Child Tags', padx=53, pady=20, command=GetChildTagsClick,  bg = 'white')
createChildBtn.place(x=30, y=150)

# Create an event loop
root.mainloop()


