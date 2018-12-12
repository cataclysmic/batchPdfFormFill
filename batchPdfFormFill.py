from tkinter import filedialog, Frame, Button, Entry, Tk, StringVar, E, Label, messagebox
from pandas import read_excel, DataFrame, ExcelFile, ExcelWriter
import pypdftk
from operator import itemgetter
from os import listdir, path
from collections import defaultdict
from html.parser import HTMLParser

# class for window content
class CoreGui(Frame):
    # initialize and define window content
    def __init__(self, master=None):
        Frame.__init__(self, master)

        self.fillPdflabel = Label(master, text="PDF Formulare befüllen mit XLSX Daten",
                                  font='Helvetica 18 bold')
        # button pdf select
        self.selectPdfFormBut = Button(master, text="PDF Formular",
                                       command=self.selectPdfForm)

        # button excel select
        self.selectExcelButton = Button(master, text="Excel Tabelle",
                                        command=self.selectExcel)

        # button target folder select
        self.selectTargetButton = Button(master, text='Ordner',
                                         command=self.selectTargetFolder)

        # button create multiple pdfs
        self.createLottaPdfButton = Button(master, text='PDF Formulare befüllen',
                                           command=self.createMultiplePDF)
        
        self.IDlabel = Label(master, text="ID-Spalten (ID1,ID2,ID3)")
        # file path for selected pdf
        self.inputPdfStr = StringVar()

        # file path for selected excel
        self.inputExcelPath = StringVar()

        # path for selected target folder
        self.inputTargetPath = StringVar()

        # text field readonly for PDF path
        self.inputPdf = Entry(master, width=40, state='readonly',
                              textvariable=self.inputPdfStr)

        # text field read only for Excel path
        self.inputID = Entry(master, width=40)

        # text field read only for Excel path
        self.inputExcel = Entry(master, width=40, state='readonly',
                                textvariable=self.inputExcelPath)

        # text field read only for target folder path
        self.inputTarget = Entry(master, width=40, state='readonly',
                                 textvariable=self.inputTargetPath)

    # ---------------------------------------------------------------------------------------
        self.extractPdflabel = Label(master, text="PDF Formulare zu XLSX extrahieren",
                                     font='Helvetica 18 bold')

        # button input folder select
        self.selectBaseButton = Button(master, text='PDF Ordner',
                                         command=self.selectBaseFolder)
        # input path folder
        self.inputBasePath = StringVar()
        self.inputBase = Entry(master, width=40, state='readonly',
                               textvariable=self.inputBasePath)

        # button input folder select
        self.selectOutputExcel = Button(master, text='Ausgabe Excel',
                                        command=self.selectOutExcel)
        # text field read only for Excel path
        self.outputExcelPath = StringVar()
        self.outputExcel = Entry(master, width=40, state='readonly',
                                 textvariable=self.outputExcelPath)

        # button extract multiple pdfs
        self.extractLottaPdfButton = Button(master, text='PDF zu XLSX extrahieren',
                                            command=self.extractMultiplePDF)
        
    # ---------------------------------------------------------------------------------------
        # initialize graphical objects in frame (order in grid)
        self.fillPdflabel.grid(row=0,column=0,columnspan=2)
        self.selectPdfFormBut.grid(row=1, column=0, sticky=E)
        self.selectExcelButton.grid(row=2, column=0, sticky=E)
        self.IDlabel.grid(row=3,column=0,sticky=E)
        self.selectTargetButton.grid(row=4, column=0, sticky=E)
        self.createLottaPdfButton.grid(row=5, column=1, sticky=E)

        self.extractPdflabel.grid(row=6,column=0,columnspan=2)
        self.selectBaseButton.grid(row=7,column=0,sticky=E)
        self.selectOutputExcel.grid(row=8,column=0,sticky=E)
        self.extractLottaPdfButton.grid(row=9, column=1, sticky=E)

        self.inputPdf.grid(row=1, column=1, sticky=E)
        self.inputExcel.grid(row=2, column=1, sticky=E)
        self.inputID.grid(row=3, column=1, sticky=E)
        self.inputTarget.grid(row=4, column=1, sticky=E)
        self.inputBase.grid(row=7, column=1, sticky=E)
        self.outputExcel.grid(row=8, column=1, sticky=E)


    def selectPdfForm(self):
        filename = filedialog.askopenfilename(initialdir="./",
                                              title="Select file",
                                              filetypes=(("PDF Dateien", "*.pdf"),
                                                       ("Alle Dateien", "*.*")))
        self.inputPdfStr.set(filename)

    def selectExcel(self):
        filename = filedialog.askopenfilename(initialdir='./',
                                              title='Select excel',
                                              filetypes=(
                                                  ('Excel Dateien', '*.xlsx'),
                                                  ('Alle Dateien', '*.*')))
        self.df = read_excel(filename)
        print("###### Gefundene Tabelle #####")
        print(self.df.head())
        print("")
        print("##### Verfügbare Spalten IDs #####")
        print("\n".join(list(self.df)))
        self.inputExcelPath.set(filename)

    def selectTargetFolder(self):
        foldername = filedialog.askdirectory()
        self.inputTargetPath.set(foldername)

    def selectBaseFolder(self):
        foldername = filedialog.askdirectory()
        self.inputBasePath.set(foldername)

    def selectOutExcel(self):
        filename = filedialog.asksaveasfilename(initialdir='./',
                                                title='Speichern unter',
                                                filetypes=(('Excel Dateien', '*.xlsx'),
                                                         ('Alle Dateien', '*.*')))
        self.outputExcelPath.set(filename)

    def createMultiplePDF(self):
        #create dataframe from excel upload
        tP = self.inputTarget.get()
        # generate output file base string
        OutPDF = path.join(self.inputPdf.get()).split('/')[-1].split(".")[0]
        ID = self.inputID.get().split(',')
        # iterate by row of data frame
        for index, row in self.df.iterrows():
            getRow = row.to_dict()
            # remove nan with empty string ''
            for el in getRow.keys():
                if getRow[el] == 'nan':
                    getRow[el] = ''
            print(getRow)
            # get IDs for name
            IDval = itemgetter(*ID)(getRow)
            if all("" == x for x in IDval):
                return(print("Fertig"))
            # if multiple ID parts join them into single string
            if type(IDval) is tuple:
                IDval = "_".join(map(str,IDval))
            filename = str(IDval)+'_'+OutPDF+'.pdf'
            createdPdfList = listdir(tP)
            # add counter if filename already exists
            if filename in createdPdfList:
                count = 1
                while filename in createdPdfList:
                    filename = IDval+'_'+OutPDF+'_'+str(count)+'.pdf'
                    count += 1
            pypdftk.fill_form('"'+path.join(self.inputPdf.get())+'"', datas=getRow,
                              out_file='"'+path.join(tP, filename)+'"',
                              flatten=False, need_appearances=True)
            print('Created: '+filename)


    def extractMultiplePDF(self):
        parser = HTMLParser()
        pdfPath = path.join(self.inputBase.get())
        xlsxOut = path.join(self.outputExcel.get())
        pdfFiles = [x for x in listdir(pdfPath) if ".pdf" in x.lower()]
        print("Gefundene PDFs:")
        print(pdfFiles)
        rows = []
        for el in pdfFiles:
            print(el) # print next filename
            raw = pypdftk.dump_data_fields('"'+path.join(pdfPath, el)+'"')
            if len(raw) > 0:
                for x in raw:
                    print(x) # print current field output
                    if b"FieldValue" in x.keys():
                        #print(x[b'FieldValue'])
                        rows = rows + [("\r".join(parser.unsecape(str(x[b'FieldName'])).split("\r")[:-1])),
                                        ("\r".join(parser.unsecape(str(x[b'FieldValue'])).split("\r")[:-1]))]
                    else:
                        rows = rows + [("\r".join(parser.unsecape(str(x[b'FieldName'])).split("\r")[:-1]), '')]
        print("Formulardaten:")
        print(rows)

        if len(rows) == 0:
            messagebox.showinfo("Hinweis","Es wurden keine Formulardaten gefunden.\nProzess abgebrochen.")
            return(0)

        drows = defaultdict(list)
        for k, v in rows:
            drows[k].append(v)
        try:
            df = DataFrame.from_dict(drows)
            print(df)
            writer = ExcelWriter(path.join(self.outputExcel.get()))
            df.to_excel(writer, 'Sheet1', index=False)
            writer.save()
        except:
            messagebox.showinfo("Hinweis","Die Struktur der Formulardaten.\nDie Anwendung verarbeitet jeweils nur ein Set Formulardaten.\nProzess abgebrochen.")


root = Tk()
root.title('Massenverarbeitung von PDF-Formularen')
#root.configure(background='gray1')
my_gui = CoreGui(root)
root.mainloop()

