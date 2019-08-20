class File:

    extension = ''
    name = ''
    directorio = ''
    ruta = ''

    def __init__(self, ruta):
        self.ruta = ruta
        self.getExtension()
        self.getName()
        self.getDirectory()

    def getExtension(self):  # obtener extension de un archivo
        if(self.ruta == ''):
            return
        i = len(self.ruta)-1
        aux = self.ruta[i]
        ext = ''
        while(aux != '.'):
            ext += aux
            i -= 1
            aux = self.ruta[i]
        self.extension = ext[::-1]

    def getName(self):  # obtener el nombre de un archivo
        if(self.ruta == ''):
            return
        i = len(self.ruta)-1
        aux = self.ruta[i]
        name = ''
        extension = True
        while(aux != '/'):
            if(not extension):
                name += aux
            i -= 1
            if(aux == '.'):
                extension = False
            aux = self.ruta[i]

        self.name = name[::-1]

    def getDirectory(self):  # obtener la ruta raiz de un archivo
        if(self.ruta == ''):
            return
        self.directorio = self.ruta[slice(
            len(self.ruta)-(len(self.name)+len(self.extension))-1)]
