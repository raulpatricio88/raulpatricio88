
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Border, Side
import time

#Producto que reune y calcula sueldos de colaboradores

book = Workbook()
sheet = book.active

top = Side(border_style='medium', color="000000")
bottom = Side(border_style='medium', color="000000")
left = Side(border_style='medium', color="000000")
right = Side(border_style='medium', color="000000")

border = Border(top=top, bottom=bottom, left=left, right=right)

jornada1 = 0
opcion = 's'
print("Buen día!")

sheet['B2'] = 'Nombre'
sheet['B2'].border = border
sheet['C2'] = 'Rut'
sheet['C2'].border = border
sheet['D2'] = 'Cargo'
sheet['D2'].border = border
sheet['E2'] = 'Sueldo base'
sheet['E2'].border = border
sheet['F2'] = 'Fecha de ingreso'
sheet['F2'].border = border
sheet['G2'] = 'Jornada'
sheet['G2'].border = border
sheet['H2'] = 'Años de antiguedad'
sheet['H2'].border = border
sheet['I2'] = 'Gratificacion legal'
sheet['I2'].border = border


for a in range (3, 71):
    while (opcion=='s'):
        print("Ingrese los datos del colaborador: ")
        nombre = input("Nombre: ")
        sheet[f'B{a}'] = nombre
        sheet[f'B{a}'].border = border
        rut = input("Rut: ")
        sheet[f'C{a}'] = rut
        sheet[f'C{a}'].border = border
        cargo = input("Cargo: ")
        sheet[f'D{a}'] = cargo
        sheet[f'D{a}'].border = border

        fechaIng = input("Fecha de ingreso: (formato dd-mm-aa) ")
        sheet[f'F{a}'] = fechaIng
        sheet[f'F{a}'].border = border

        jornada = input("Tipo de jornada: c: completa / m: media ")
        jornada = jornada.lower()
        if(jornada=='c'):
            jornada1='Completa'
        else:
            jornada1='Media'
        sheet[f'G{a}'] = jornada1
        sheet[f'G{a}'].border = border
        antiwedad = input("Años de antiguedad: ")
        sheet[f'H{a}'] = antiwedad
        sheet[f'H{a}'].border = border
        sueldoB = float(input("Sueldo base: "))
        sheet[f'E{a}'] = sueldoB
        sheet[f'E{a}'].border = border
        grat = float(input("Gratificacion legal: $"))
        sheet[f'I{a}'] = grat
        sheet[f'I{a}'].border = border

        sueldoN= sueldoB + grat - (sueldoB)*0.19
        print("El sueldo neto es: $", + float(sueldoN))
        a+=1
        opcion = input("¿Desea agregar otro colaborador? s: si / n: no \n\n")

print("Documento Excel generado en directorio del programa. Que tenga un buen día!")
book.save('Prueba_producto.xlsx')


