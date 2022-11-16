from tkinter.ttk import *
from tkinter import *
from tkinter.filedialog import askopenfilename
import pandas as pd

#Create and configure mainWindow using Tkinter
mainWindow = Tk()
mainWindow.title('Excel-Scraper')
mainWindow.geometry('1000x500')
mainWindow.configure(bg='#FAFAFA')


#Create (frame0) in (mainWindow)
frame0 = Frame(mainWindow, bd=5, bg='#FAFAFA')
frame0.pack(side=TOP, fill='both', padx=3, pady=0)

#Create (colorsFrame) in (mainWindow)
colorsFrame = Frame(mainWindow, bd=2, bg='#FAFAFA')
colorsFrame.pack(side=TOP)


#mainScreen displays initial welcome and instruction text
def mainScreen(frame = None):
    #Destroy (frame) if (frame) is present
    try:
        frame.destroy()
    except:
        pass
    
    #Delete widgets from window
    resetWindow()
    
    #(frame0) Display welcome text and instructions
    Label(frame0, text = "Welcome to Excel-Scraper!", font = ('Comic Sans MS', 12), bg='#FAFAFA').pack(side=TOP)
    Label(frame0, text = 'Please upload file for Scraping:\n(File must be Excel)', font = ('Comic Sans MS',10), bg = '#FAFAFA').pack()

    #(frame0) Display button to load openFile function
    Button(frame0, text = 'Browse Files', command = openFile).pack(side=BOTTOM)

#First color option
def colorsOp0(color0, color1, filePath, frame):
    #Black and green combination
    color0 = '#0D0208'
    color1 = '#00FF41'
    firstOpen= False
    frame.destroy()
    openFile(color0, color1, open, filePath)

#Second color option
def colorsOp1(color0, color1, filePath, frame):
    #White and black combination
    color0 = '#FAFAFA'
    color1 = 'black'
    firstOpen= False
    frame.destroy()
    openFile(color0, color1, open, filePath)

#Third color option
def colorsOp2(color0, color1, filePath, frame):
    #Blue and white combination
    color0 = '#012456'
    color1 = '#FFFFF9'
    firstOpen= False
    frame.destroy()
    openFile(color0, color1, open, filePath)

#openFile opens user file and reads its contents
def openFile(color0 = '#0D0208',color1 = '#00FF41', firstOpen= True, filePath = None, frame = None):
    #Display file search window if True
    if firstOpen== True:
        filePath = askopenfilename(filetypes=[('Excel file', '*.xls *.xlsx *.csv')])

    #Create frame in (mainWindow)
    frame = Frame(mainWindow, bg='#FAFAFA')
    frame.pack(fill="both", expand=100, padx=3, pady=0)   
    
    resetWindow()
    #Opens file 
    try:
        outputFile = pd.read_excel(filePath, sheet_name=None)
    except:
        try:
            outputFile = pd.read_csv(filePath)
        #If file extension invalid or no file chosen return to main screen
        except Exception as e:
            mainScreen()
            #(frame0) Display error message on main screen
            Label(frame0, text = 'ERROR: Failed to open file!', background = '#FAFAFA', foreground = 'red').pack(side=BOTTOM)
            return

    #Create the frame, canvas, and scroll bar to be used by outputFile
    canvas = Canvas(frame, bg = color0)
    canvas.pack(side = LEFT, fill=BOTH, expand=1000)

    scrollbar = Scrollbar(frame, orient=VERTICAL , command = canvas.yview)
    scrollbar.pack(side=RIGHT, fill=Y)

    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion = canvas.bbox("all")))

    frame1 = Frame(canvas)

    canvas.create_window((0,0), window = frame1, anchor="nw")   

    #pd.option_context displays full file
    with pd.option_context('display.max_rows', None, 'display.max_columns', None):
        #(colorsFrame) Display color changer buttons
        Button(colorsFrame, text = 'Color Combo 1', font = 10, foreground='#00FF41', background='#0D0208',command = lambda: colorsOp0(color0, color1, filePath, frame)).grid(row=0, column=2, padx=10, pady=5)
        Button(colorsFrame, text = 'Color Combo 2', font = 10, foreground='black', background='#FAFAFA',command = lambda: colorsOp1(color0, color1, filePath, frame)).grid(row=0, column=3, padx=10, pady=5)
        Button(colorsFrame, text = 'Color Combo 3', font = 10, foreground='#FFFFF9', background='#012456',command = lambda: colorsOp2(color0, color1, filePath, frame)).grid(row=0, column=4, padx=10, pady=5)

        #(frame0) Display return to main screen button and file uploaded successfully label
        Button(frame0, text = 'Press to Return to Main Screen', command = lambda: mainScreen(frame)).pack(side=TOP)
        Label(frame0, text = 'File Uploaded Successfully!', foreground='green',background='#FAFAFA').pack(pady = 5, padx =5)
        Label(frame0, text = 'Click below to change color of results:', background = '#FAFAFA', font = 12).pack(side = BOTTOM, padx = 5)

        #(frame1) Display outputFile contents
        Label(frame1, text = outputFile, font=('Arial', 10), foreground = color1, background = color0,underline = 1).pack(fill=BOTH, expand=1000)
        
#resetWindow() clears window of all widgets
def resetWindow():
    #for Loops destroys all widgets in frames
    for widgets in frame0.winfo_children():
        widgets.destroy()
    for widgets in colorsFrame.winfo_children():
        widgets.destroy()
    return

#Run mainScreen function
mainScreen()

mainWindow.mainloop()