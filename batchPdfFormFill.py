from tkinter import filedialog, Frame, Button, Entry, Tk, StringVar, E, Label
from pandas import read_excel, DataFrame, ExcelFile, ExcelWriter
import pypdftk
from operator import itemgetter
from os import listdir
from collections import defaultdict

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
        filename = filedialog.askopenfilename(initialdir = './',
                                              title = 'Select excel',
                                              filetypes =(('Excel Dateien','*.xlsx'),('Alle Dateien','*.*')))
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
        df = read_excel(self.inputExcel.get())
        tP = self.inputTarget.get()
        # generate output file base string
        OutPDF = self.inputPdf.get().split('/')[-1].split(".")[0]
        ID = self.inputID.get().split(',')
        for index, row in df.iterrows():
            getRow = row.to_dict()
            IDval = itemgetter(*ID)(getRow)
            if type(IDval) is tuple:
                IDval = "_".join(IDval)
            count = index + 1
            filename = OutPDF+'_'+IDval+'.pdf'
            createdPdfList = listdir(tP)
            if filename in createdPdfList:
                count = 1
                while filename in createdPdfList:
                    filename = OutPDF+'_'+IDval+'_'+str(count)+'.pdf'
                    count += 1
            pypdftk.fill_form('"'+self.inputPdf.get()+'"', datas=getRow,
                              out_file='"'+tP+'/'+filename+'"',
                              flatten=False, need_appearances=True)
            print('Created: '+filename)


    def extractMultiplePDF(self):
        pdfPath = self.inputBase.get()
        xlsxOut = self.outputExcel.get()
        pdfFiles = [x for x in listdir(pdfPath) if ".pdf" in x.lower()]
        rows = []
        for el in pdfFiles:
            raw = pypdftk.dump_data_fields(pdfPath+'/'+el)
            rows = rows + [(x['FieldName'], x['FieldValue']) for x in raw]
        drows = defaultdict(list)
        for k, v in rows:
            drows[k].append(v)
        df = DataFrame.from_dict(drows)
        print(df)
        writer = ExcelWriter(self.outputExcel.get())
        df.to_excel(writer, 'Sheet1', index=False)
        writer.save()


root = Tk()
root.title('Massenverarbeitung von PDF-Formularen')
#root.configure(background='gray1')
my_gui = CoreGui(root)
root.mainloop()

