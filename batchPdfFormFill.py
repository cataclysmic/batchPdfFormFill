from tkinter import filedialog, Frame, Button, Entry, Tk, StringVar, E, Label
from pandas import read_excel
import pypdftk 
from operator import itemgetter


# class for window content
class CoreGui(Frame):
    # initialize and define window content
    def __init__(self, master=None):
        Frame.__init__(self, master)

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
        self.createLottaPdfButton = Button(master, text='LOS',
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

        # initialize graphical objects in frame (order in grid)
        self.selectPdfFormBut.grid(row=0, column=0, sticky=E)
        self.selectExcelButton.grid(row=1, column=0, sticky=E)
        self.IDlabel.grid(row=2,column=0,sticky=E)
        self.selectTargetButton.grid(row=3, column=0, sticky=E)

        self.inputPdf.grid(row=0, column=1, sticky=E)
        self.inputExcel.grid(row=1, column=1, sticky=E)
        self.inputID.grid(row=2, column=1, sticky=E)
        self.inputTarget.grid(row=3, column=1, sticky=E)

        self.createLottaPdfButton.grid(row=4, column=1, sticky=E)

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
            IDname = "_".join(IDval)
            count = index + 1
            pypdftk.fill_form('"'+self.inputPdf.get()+'"', datas=getRow,
                              out_file='"'+tP+'/'+
                              OutPDF+'_'+IDname+'_'+str(count)+'.pdf'+'"', flatten=False, need_appearances=True)
            print('Created: '+OutPDF+'_'+IDname+'_'+str(count)+'.pdf')

root = Tk()
root.title('Massengenerierung PDF aus Excel')
#root.configure(background='gray1')
my_gui = CoreGui(root)
root.mainloop()

