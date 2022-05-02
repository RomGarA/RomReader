#Librerias necesarias para el importe de datos
from sre_constants import SUCCESS
from cv2 import BORDER_WRAP
from pdf2image import convert_from_path
import cv2
import os
import pytesseract
from PIL import Image
import xlsxwriter
import fitz

libro = xlsxwriter.Workbook('Resultados.xlsx')    
hoja = libro.add_worksheet()

def extractiontext():
    try:
        os.chdir(r'inputFiles')
    except:
        os.chdir(r'..\inputFiles')

    images = convert_from_path(str(counter) + '.pdf')
    for i in range(len(images)):
        os.chdir(r'..\cacheimg')
        images[i].save('page.jpg', 'JPEG')

    image = cv2.imread('page.jpg')

    Empresa     =   image[188:260,200:900]
    Cotizantes  =   image[520:570,1535:1605]
    Serial      =   image[1855:1895,385:1445]
    Total       =   image[1737:1790,1310:1610]
    Patronal    =   image[270:310,390:580]

    cv2.imwrite("Empresa.jpg",Empresa)
    cv2.imwrite("Cotizantes.jpg",Cotizantes)
    cv2.imwrite("Serial.jpg",Serial)
    cv2.imwrite("Total.jpg",Total)
    cv2.imwrite("Patronal.jpg",Patronal)

#Extraer texto de cada imagen
def imagetotxt(i):
    os.chdir(r'..\inputFiles')

    Id = 'TRANSFERENCIA ELECTRÓNICA'
    id = ''

    pdf_documento = str(i) + ".pdf"
    documento = fitz.open(pdf_documento)
    pagina = documento.load_page(0)
    text = pagina.get_text("text") 
    m = int(text.find(Id)) + 26
    b = 0

    while b<62:    
        id += text[m]
        m+=1
        b+=1

    os.chdir(r'..\cacheimg')

    txtEmpresa = str(((pytesseract.image_to_string(Image.open("Empresa.jpg")))))
    txtEmpresa = txtEmpresa.replace('-\n','')
    txtCotizantes = (pytesseract.image_to_string("Cotizantes.jpg",lang='eng', config='--psm 10 --oem 3 -c tessedit_char_whitelist=0123456789'))
    txtCotizantes = txtCotizantes.replace('-\n','')
    txtSerial = (pytesseract.image_to_string("Serial.jpg", config='--psm 10 --oem 3 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLAMÑNOPQRSTUVWXYZ-'))
    txtSerial = txtSerial.replace('-\n','')
    txtTotal = (pytesseract.image_to_string("Total.jpg",lang='eng', config='--psm 10 --oem 3'))
    txtTotal = txtTotal.replace('-\n','')
    txtPatronal = (pytesseract.image_to_string("Patronal.jpg", lang='eng', config='--psm 10 --oem 3 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLAMÑNOPQRSTUVWXYZ-'))
    txtPatronal = txtPatronal.replace('-\n','')   

    hoja.write(i,0,txtEmpresa)
    hoja.write(i,1,txtCotizantes)
    hoja.write(i,2,id)
    hoja.write(i,3,txtTotal)
    hoja.write(i,4,txtPatronal)

#Para borrar los residuos
def clear():
    os.chdir(r'..\cacheimg')
    os.remove("Empresa.jpg")
    os.remove("Cotizantes.jpg")
    os.remove("Serial.jpg")
    os.remove("Total.jpg")
    os.remove("Patronal.jpg")
    os.remove("page.jpg")
    #os.remove("Resultados.xlsx")

#Empezar hasta terminar
while True:
    x = int(input("Numero de archivos: ")) 
    x += 1
    counter = 0
    for counter in range(1,x):
        extractiontext()
        imagetotxt(counter)
        clear()
        os.system ("cls")
        print("Proceso:\t" + str(counter*100/x) + "% [" + ((int((counter*100/x)/10))*'■') + ((10-(int((counter*100/x)/10)))*'-') + ']' )

    os.system ("cls")
    print('Proceso: Finalizado [■■■■■■■■■■]')
    libro.close()
    break