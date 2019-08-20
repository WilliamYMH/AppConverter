import subprocess
import shutil
import win32com.client
import os
import gc


def convertOdfToMso(file):  # convertir archivo odf a msOffice formato antiguo

    if(file.ruta != ''):
        if(file.extension == 'odt'):
            shutil.copy2(file.ruta, file.directorio+file.name+'_converted.doc')
        elif(file.extension == 'odp'):
            shutil.copy2(file.ruta, file.directorio+file.name+'_converted.ppt')
        elif(file.extension == 'ods'):
            shutil.copy2(file.ruta, file.directorio+file.name+'_converted.xls')
        else:
            return 'Extension de archivo no valido'
        return 'Archivo convertido satisfactoriamente'

    else:
        return 'Ruta invalida'


def convertMsoToOdf(file):  # convertir archivo msofice formato antiguo a odf

    if(file.ruta != ''):
        if(file.extension == 'doc'):
            shutil.copy2(file.ruta, file.directorio+file.name+'_converted.odt')
        elif(file.extension == 'ppt'):
            shutil.copy2(file.ruta, file.directorio+file.name+'_converted.odp')
        elif(file.extension == 'xls'):
            shutil.copy2(file.ruta, file.directorio+file.name+'_converted.ods')
        else:
            return 'Extension de archivo no valido'
        return 'Archivo convertido satisfactoriamente'

    else:
        return 'Ruta invalida'


def convertOdfToMso2(file):  # convertir archivo odf a msOffice formato actual
    if(file.ruta != ''):
        if(file.extension == 'odt'):
            gc.collect()
            word = win32com.client.Dispatch("Word.Application")
            word.visible = 0

            wb = word.Documents.Open(os.path.abspath(file.ruta))
            # file format for docx
            wb.SaveAs2(os.path.abspath(file.directorio+file.name +
                                       '_converted.docx'), FileFormat=16)
            wb.Close()
            wb = None
            word = None
            # word.Quit()
            # subprocess.call(
            #  ['soffice', '--headless', '--convert-to', 'odt', file.ruta])

        elif(file.extension == 'odp'):
            gc.collect()
            power = win32com.client.Dispatch('Powerpoint.Application')
            wb = power.Presentations.Open(os.path.abspath(
                file.ruta), WithWindow=0)
            # wb.Activate()
            wb.SaveAs(os.path.abspath(file.directorio+file.name +
                                      '_converted.pptx'), FileFormat=24)
            wb.Close()
            wb = None
            # power.Quit()
            power = None
        elif(file.extension == 'ods'):
            gc.collect()
            excel = win32com.client.Dispatch(
                'Excel.Application')
            excel.visible = 0

            ex = excel.Workbooks.Open(os.path.abspath(file.ruta))
            ex.Activate()
            ex.SaveAs(os.path.abspath(file.directorio +
                                      file.name+'_converted.xlsx'), FileFormat=51)
            ex.Close()
            ex = None
            # excel.Quit()
            excel = None
        else:
            return 'Extension de archivo no valido'
        return 'Archivo convertido satisfactoriamente'

    else:
        return 'Ruta invalida'


def convertMso2toOdf(file):  # convertir archivo msoffice actual a odf
    if(file.ruta != ''):
        if(file.extension == 'docx'):
            gc.collect()
            word = win32com.client.Dispatch("Word.Application")
            word.visible = 0

            wb = word.Documents.Open(os.path.abspath(file.ruta))
            # file format for docx
            wb.SaveAs2(os.path.abspath(file.directorio+file.name +
                                       '_converted.odt'), FileFormat=23)
            wb.Close()
            wb = None
            word = None
            # word.Quit()
            # subprocess.call(
            #  ['soffice', '--headless', '--convert-to', 'odt', file.ruta])

        elif(file.extension == 'pptx'):
            gc.collect()
            power = win32com.client.Dispatch('Powerpoint.Application')
            wb = power.Presentations.Open(os.path.abspath(
                file.ruta), WithWindow=0)
            # wb.Activate()
            wb.SaveAs(os.path.abspath(file.directorio+file.name +
                                      '_converted.odp'), FileFormat=35)
            wb.Close()
            wb = None
            # power.Quit()
            power = None
        elif(file.extension == 'xlsx'):
            gc.collect()
            excel = win32com.client.Dispatch(
                'Excel.Application')
            excel.visible = 0

            ex = excel.Workbooks.Open(os.path.abspath(file.ruta))
            ex.Activate()
            ex.SaveAs(os.path.abspath(file.directorio +
                                      file.name+'_converted.ods'), FileFormat=60)
            ex.Close()
            ex = None
            # excel.Quit()
            excel = None
        else:
            return 'Extension de archivo no valido'
        return 'Archivo convertido satisfactoriamente'

    else:
        return 'Ruta invalida'
