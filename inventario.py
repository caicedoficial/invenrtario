import platform
import psutil
import socket
import uuid
import openpyxl
import os
from tkinter import Tk, Label, Entry, Button

def obtener_datos():
    responsable = entry_responsable.get()
    nombre_equipo = platform.node()
    sistema_operativo = platform.system() + " " + platform.release()
    procesador = platform.processor()
    ram = str(round(psutil.virtual_memory().total / (1024.0 **3))) + " GB"
    almacenamiento = str(round(psutil.disk_usage('/').total / (1024.0 **3))) + " GB"
    direccion_mac = ':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff) for elements in range(0,2*6,2)][::-1])
    direccion_ip = socket.gethostbyname(socket.gethostname())
    estado = "Activo"

    ruta_excel = "./inventario.xlsx"

    if os.path.exists(ruta_excel):
        wb = openpyxl.load_workbook(ruta_excel)
        hoja = wb.active
    else:
        wb = openpyxl.Workbook()
        hoja = wb.active
        encabezados = ["Responsable", "Nombre del equipo", "Sistema Operativo", "Procesador", "RAM", "Almacenamiento", "Direccion MAC", "Direccion IP", "Estado"]
        hoja.append(encabezados)

    datos = [responsable ,nombre_equipo, sistema_operativo, procesador, ram, almacenamiento, direccion_mac, direccion_ip, estado]
    hoja.append(datos)

    wb.save(ruta_excel)
    root.destroy()

root = Tk()
root.title("Inventario de Equipos")

Label(root, text="Ingrese el nombre del responsable:").grid(row=0, column=0)
entry_responsable = Entry(root)
entry_responsable.grid(row=0, column=1)

Button(root, text="Guardar", command=obtener_datos).grid(row=1, columnspan=2)

root.mainloop()