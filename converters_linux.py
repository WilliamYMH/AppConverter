import subprocess
import shutil
import os


def convertOdfToMso(file):  # convertir archivo odf a msOffice formato antiguo

    if(file.ruta != ''):
        if(file.extension == 'odt'):
            shutil.copy2(os.path.abspath(file.ruta), os.path.abspath(
                file.directorio+file.name+'_converted.doc'))
        elif(file.extension == 'odp'):
            shutil.copy2(os.path.abspath(file.ruta), os.path.abspath(
                file.directorio+file.name+'_converted.ppt'))
        elif(file.extension == 'ods'):
            shutil.copy2(os.path.abspath(file.ruta), os.path.abspath(
                file.directorio+file.name+'_converted.xls'))
        else:
            return 'Extension de archivo no valido'
        return 'Archivo convertido satisfactoriamente'

    else:
        return 'Ruta invalida'


def convertMsoToOdf(file):  # convertir archivo msofice formato antiguo a odf

    if(file.ruta != ''):
        if(file.extension == 'doc'):
            shutil.copy2(os.path.abspath(file.ruta), os.path.abspath(
                file.directorio+file.name+'_converted.odt'))
        elif(file.extension == 'ppt'):
            shutil.copy2(os.path.abspath(file.ruta), os.path.abspath(
                file.directorio+file.name+'_converted.odp'))
        elif(file.extension == 'xls'):
            shutil.copy2(os.path.abspath(file.ruta), os.path.abspath(
                file.directorio+file.name+'_converted.ods'))
        else:
            return 'Extension de archivo no valido'
        return 'Archivo convertido satisfactoriamente'

    else:
        return 'Ruta invalida'


def convertOdfToMso2(file):  # convertir archivo odf a msOffice formato actual
    if(file.ruta != ''):
        if(file.extension == 'odt'):
            subprocess.call(
                ['soffice', '--headless', '--convert-to', 'docx', os.path.abspath(file.ruta)])

        elif(file.extension == 'odp'):
            subprocess.call(
                ['soffice', '--headless', '--convert-to', 'pptx', os.path.abspath(file.ruta)])
        elif(file.extension == 'ods'):
            subprocess.call(
                ['soffice', '--headless', '--convert-to', 'xlsx', os.path.abspath(file.ruta)])
        else:
            return 'Extension de archivo no valido'
        return 'Archivo convertido satisfactoriamente'

    else:
        return 'Ruta invalida'


def convertMso2toOdf(file):  # convertir archivo msoffice actual a odf
    if(file.ruta != ''):
        if(file.extension == 'docx'):
            subprocess.call(
                ['soffice', '--headless', '--convert-to', 'odt', os.path.abspath(file.ruta)])

        elif(file.extension == 'pptx'):
            subprocess.call(
                ['soffice', '--headless', '--convert-to', 'odp', os.path.abspath(file.ruta)])
        elif(file.extension == 'xlsx'):
            subprocess.call(
                ['soffice', '--headless', '--convert-to', 'ods', os.path.abspath(file.ruta)])
        else:
            return 'Extension de archivo no valido'
        return 'Archivo convertido satisfactoriamente'

    else:
        return 'Ruta invalida'
