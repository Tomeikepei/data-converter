# -*- coding: utf-8 -*-

from tkinter import *
from tkinter import filedialog
import re
import xlwt
import os


def abrir():
    archivo = filedialog.askopenfile()

    col = 1
    hoja = xlwt.Workbook()
    hoja_libro = hoja.add_sheet('BP')
    for line in archivo:
        new_group = line.split()

        try:
            item1 = (new_group[0])
        except:
            continue
        if len(item1) > 19:
            continue
        else:
            t = list(item1)

            fecha = (t[5] + t[6] + t[7] + t[8] + t[9] + t[10] + t[11] + t[12])

            hoja_libro.write(col, 0, fecha)

        line = line.rstrip()
        combustible = re.findall('.[0-9]+00000([0-9]+)', line)


        hoja_libro.write(col, 1, combustible[0])

        try:
            item3 = (new_group[3])
        except:
            continue

        hoja_libro.write(col, 2, item3)

        try:
            item4 = (new_group[4])
        except:
            continue

        hoja_libro.write(col, 3, item4)

        col = col + 1

    hoja_libro.write(0, 0, 'Fecha')
    hoja_libro.write(0, 1, 'Combustible')
    hoja_libro.write(0, 2, 'Matricula')
    hoja_libro.write(0, 3, 'Numero de Vuelo')

    hoja.save('C:\BP.xls')
    os.startfile('C:\BP.xls')

ventana = Tk()
ventana.config(bg="black")
ventana.geometry("200x200")
ventana.title("Conversor datos BP")
foto = PhotoImage(file="bp.gif")
cuadro = Label(ventana, bd = 0, image = foto) .place(x = 10, y = 10)
botonAbrir = Button(ventana, text = "Abrir y generar", command = abrir)
botonAbrir.grid(padx = 150, pady = 100)
botonAbrir.pack()
ventana.mainloop()
