from pyhunter import PyHunter
from openpyxl import Workbook
from openpyxl.styles import Font
import getpass


def Busqueda(organizacion):
    #Cantidad de resultados esperados de la búsqueda
    #El límite MENSUAL de Hunter es 50, cuidado!
    resultado = hunter.domain_search (company = organizacion, limit = 1, emails_type = 'personal')
    return resultado


def GuardarInformacion(datosEncontrados,organizacion):
    libro = Workbook()
    hoja = libro.create_sheet(organizacion)
    libro.save("Hunter" + organizacion + ".xlsx")
    hoja["A1"].font = Font(color="FF0000", bold=True)
    hoja["A1"] = "Datos"
    hoja["D1"] = (datosEncontrados["emails"][0]["value"])
    hoja["A2"].font = Font(color="FF0000", bold=True)
    hoja["A2"] = "Type"
    hoja["D2"] = (datosEncontrados["emails"][0]["type"])
    hoja["A3"].font = Font(color="FF0000", bold=True)
    hoja["A3"] = "confidence"
    hoja["D3"] = (datosEncontrados["emails"][0]["confidence"])
    #Agrega el codigo necesario para guardar en formato tabla
    #dentro del libro de Excel, información que consideres relevante
    #de lo obtenido en la búsqueda.
    libro.save("Hunter" + organizacion + ".xlsx")
    

print("Script para buscar información")
apikey = getpass.getpass("Ingresa tu API key: ")
hunter = PyHunter (apikey)
orga = input("Dominio a investigar: ")
datosEncontrados = Busqueda(orga)
if datosEncontrados == None:
    exit()
else:
    print(datosEncontrados)
    print(type(datosEncontrados))
    GuardarInformacion(datosEncontrados,orga)

print ("------------------------------------")
print (datosEncontrados["emails"][0]["value"])
print ("------------------------------------")

for x,y  in datosEncontrados.items():
        print(x,y,type(y))
