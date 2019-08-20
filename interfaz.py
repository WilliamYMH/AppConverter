from tkinter import filedialog, messagebox
from tkinter import Button, mainloop, Label
from File import File
from converters import *
import tkinter


class Controller:

    ruta = ''
    file = None
    list = None
    optDropdown = None
    variable = None
    labelDropdown = None

    def addDropdown(self):

        self.variable = tkinter.StringVar(root)
        self.variable.set(controller.list[0])
        self.optDropdown = tkinter.OptionMenu(
            root, self.variable, *controller.list)
        #opt.config(width=90, font=('Helvetica', 12))
        self.labelDropdown = Label(
            root, text="Tienes que elegir a que formato lo convertiras: ")
        self.labelDropdown.pack(anchor=tkinter.NW)
        self.optDropdown.pack()

    def getOptionConverter(self):
        self.list = None
        if(self.file.extension == 'odt'):
            self.list = ['doc', 'docx']
            self.addDropdown()
        elif(self.file.extension == 'odp'):
            self.list = ['ppt', 'pptx']
            self.addDropdown()
        elif(self.file.extension == 'ods'):
            self.list = ['xls', 'xlsx']
            self.addDropdown()
        self.list = None

    def cargarFile(self):
        self.ruta = filedialog.askopenfilename(parent=root)
        self.file = File(self.ruta)
        if(self.file.extension == 'odt' or self.file.extension == 'ods' or self.file.extension == 'odp'
                or self.file.extension == 'doc' or self.file.extension == 'xls' or self.file.extension == 'ppt'
                or self.file.extension == 'docx' or self.file.extension == 'xlsx' or self.file.extension == 'pptx'):
            messagebox.showinfo(title='Archivo cargado',
                                message='Archivo cargado con exito')
            if(self.optDropdown is not None):
                self.optDropdown.forget()
                self.labelDropdown.forget()
            self.getOptionConverter()
        else:
            self.ruta = ''
            return messagebox.showwarning(title='Error',
                                          message='Extension de archivo no valido')

    def convertFile(self):
        if(self.ruta == ''):
            return messagebox.showerror(title='Error',
                                        message='Verifique que primero haya cargado un archivo')
        if(self.file.extension == 'odt' or self.file.extension == 'ods' or self.file.extension == 'odp'):

            if(self.variable.get() == 'doc' or self.variable.get() == 'xls' or self.variable.get() == 'ppt'):
                messagebox.showinfo(title='Info',
                                    message=convertOdfToMso(self.file))
            elif(self.variable.get() == 'docx' or self.variable.get() == 'xlsx' or self.variable.get() == 'pptx'):
                messagebox.showinfo(title='Info',
                                    message=convertOdfToMso2(self.file))

            self.ruta = ''
        elif(self.file.extension == 'doc' or self.file.extension == 'xls' or self.file.extension == 'ppt'):
            messagebox.showinfo(title='Info',
                                message=convertMsoToOdf(self.file))
            self.ruta = ''
        elif(self.file.extension == 'docx' or self.file.extension == 'xlsx' or self.file.extension == 'pptx'):

            messagebox.showinfo(title='Info',
                                message=convertMso2toOdf(self.file))
            self.ruta = ''
        else:
            self.ruta = ''
            return messagebox.showwarning(title='Error',
                                          message='Extension de archivo no valido')


root = tkinter.Tk()
root.geometry("300x220")
root.title('Converter')
controller = Controller()
button_cargar = Button(text='cargar archivo', command=controller.cargarFile)
button_convertir = Button(text='convertir archivo',
                          command=controller.convertFile)
label = Label(root, text="Cargue el archivo a convertir: ")
label.pack(anchor=tkinter.NW)
button_cargar.pack()
label = Label(root, text="Conviertelo!: ")
label.pack(anchor=tkinter.NW)
button_convertir.pack()
mainloop()
