
# importing modules for tk 
import array
from ntpath import join
import tkinter as tk
from tkinter import filedialog, Text
import os

# importing modules for reading excel files
import openpyxl

# importing modules for pdf
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors

import pandas as pd
import numpy as nm


root = tk.Tk()
exfiles = []

def addFile():

    for widget in frame.winfo_children():
        widget.destroy()
    
    filename = filedialog.askopenfilename(initialdir="/", title="Select File", filetypes=(("Excel","*.xlsx"),("all files","*.*")))
    exfiles.append(filename)
    for exfile in exfiles:
        label = tk.Label(frame, text=exfile, bg="gray")
        label.pack()


def runCon():
    for exfile in exfiles:
        path =  exfile
        print (path) 
        df = pd.read_excel(path, skiprows=1)
        

        for i in range(1,6):
            fname=df['First Name '].where(df['Candidate No. (Need not appear on the scorecard)']==i).dropna().unique()
            lname=df['Last Name '].where(df['Candidate No. (Need not appear on the scorecard)']==i).dropna().unique()
            fulname=fname+' '+lname
            reg_numb =df['Registration Number'].where(df['Candidate No. (Need not appear on the scorecard)']==i).dropna().unique()
            round= df['Round'].where(df['Candidate No. (Need not appear on the scorecard)']==i).dropna().unique()
            grade=df['Grade '].where(df['Candidate No. (Need not appear on the scorecard)']==i).dropna().unique()
            name_sch=df['Name of School '].where(df['Candidate No. (Need not appear on the scorecard)']==i).dropna().unique()
            gender = df['Gender'].where(df['Candidate No. (Need not appear on the scorecard)']==i).dropna().unique()
            dateofbirth=df['Date of Birth '].where(df['Candidate No. (Need not appear on the scorecard)']==i).dropna().unique()
            country = df['Country of Residence'].where(df['Candidate No. (Need not appear on the scorecard)']==i).dropna().unique()
            max_marks = str(100)
            city=df['City of Residence'].where(df['Candidate No. (Need not appear on the scorecard)']==i).dropna().unique()
            scored_marks=df['Your score'].where(df['Candidate No. (Need not appear on the scorecard)']==i).dropna().sum()
            result = df['Final result'].where(df['Candidate No. (Need not appear on the scorecard)']==i).dropna().unique()
            dateoftest=df['Date and time of test'].where(df['Candidate No. (Need not appear on the scorecard)']==i).dropna().unique()

            img = 'Pics for assignment\\'+str(fulname)[1:-1]+'.PNG'

            
            print(str(fulname)[1:-1])
            print(int(reg_numb))
            print(str(fname)[1:-1])
            print(str(lname)[1:-1])
            print(int(round))
            print(int(grade))
            print(str(name_sch)[1:-1])
            print(str(dateofbirth)[1:-1])
            print(str(city)[1:-1])
            print(str(dateoftest)[1:-1])
            print(str(gender)[1:-1])
            print(str(country)[1:-1])
            print(max_marks)
            print(scored_marks)
            print(str(result)[1:-1])
            print(img)
            print("\n\n")

            # ###################################
            # Content
            fileName = str(fulname)[1:-1].replace("'", "")+'.pdf'
            documentTitle = 'Document title!'
            title = 'Competative examination'

            subTitle = 'Report Card'

            textLines = [
            'Full Name :'+ str(fulname)[1:-1].replace("'", ""),
            'Round :'+ str(round)[1:-1].replace("'", ""),
            'F Name :'+ str(fname)[1:-1].replace("'", ""),
            'L Name :'+ str(lname)[1:-1].replace("'", ""),
            'Registration Number :'+ str(int(reg_numb))[1:-1],
            'Grade :'+ str(grade)[1:-1],
            'Name of School :'+ str(name_sch)[1:-1].replace("'", "") ,
            'Gender :'+ str(gender)[1:-1].replace("'", ""),
            'Date of Birth :'+ str(dateofbirth)[1:-1].replace("'", ""),
            'City of Residence :'+ str(city)[1:-1].replace("'", ""),
            'Country of Residence :'+ str(country)[1:-1].replace("'", ""),
            'Date and Time of Test:'+ str(dateoftest)[1:-1].replace("'", ""),
            'Total Marks :'+ str(max_marks).replace("'", ""),
            'Marks Scored :'+ str(scored_marks).replace("'", ""),
            'Result :'+ str(result)[1:-1].replace("'", ""),
            ]

            image='image.jpg'

            import os
            im=os.path.join('Pics for assignment/', str(fulname)[1:-1])
            img1=im+'.PNG'
            image1=img1.replace("'", "")

            #image1='Pics for assignment/' + str(fulname)[1:-1] +'.PNG'
            


            # ###################################
            # 0) Create document 
            from reportlab.pdfgen import canvas 

            pdf = canvas.Canvas(fileName)
            pdf.setTitle(documentTitle)


            #drawMyRuler(pdf)

            # Register a new font
            from reportlab.pdfbase.ttfonts import TTFont
            from reportlab.pdfbase import pdfmetrics

            pdf.setFont("Courier-Bold", 36)
            pdf.drawCentredString(300, 770, title)

            # ###################################
            # 5) Draw a image
            pdf.drawInlineImage(image, 250, 700)

            # ###################################
            # 3) Draw a li
            pdf.line(30, 680, 550, 680)

            # ###################################
            # 2) Sub Title 
            # RGB - Red Green and Blue
            pdf.setFillColorRGB(0, 0, 255)
            pdf.setFont("Courier-Bold", 24)
            pdf.drawCentredString(290,650, subTitle)

            # ###################################
            # 5) Draw a image
            pdf.drawInlineImage(image1, 250, 500, width=90, height=90)

            # ###################################
            # 4) Text object :: for large amounts of text
            from reportlab.lib import colors

            text = pdf.beginText(40, 400)
            text.setFont("Courier", 14)
            text.setFillColor(colors.green)
            for line in textLines:
                text.textLine(line)

            pdf.drawText(text)

            pdf.save()

    exit()



canva = tk.Canvas(root, height=200,width=500, bg="#ccffcc")
canva.pack()

frame = tk.Frame(root,bg="white")
frame.place(relwidth=0.8, relheight=0.8, relx=0.1, rely=0.1)

openFile = tk.Button(root, text="Open File", padx=10, pady=5, fg="Black", bg="#ccffcc", command=addFile)
openFile.pack()


runConvert = tk.Button(root, text="Convert", padx=10, pady=5, fg="Black", bg="#ccffcc", command=runCon)
runConvert.pack()


root.mainloop()