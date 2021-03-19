from tkinter import *
from tkinter import ttk
from tkinter.messagebox import *
from PIL import Image, ImageTk
import xlsxwriter as xw 

wbcompt = xw.Workbook("Listado_de_Competidores.xlsx")

class App(Tk):

    def __init__(self):
        Tk.__init__(self)
        self.geometry("1000x450")
        self.title('Tkd Competitors')
        self.iconbitmap(r'.\Proyecto\tkd.ico')
        self.resizable(0,0)

        self._frame = None
        self.switch_frame(Competidor)

    def switch_frame(self, frame_class):
        new_frame = frame_class(self)
        if self._frame is not None:
            self._frame.destroy()
        self._frame = new_frame
        self._frame.pack()
        
# Agregar Competidor
class Competidor(Frame):
    def __init__(self, master):

        Frame.__init__(self,master)

        # Frame buttons left
        btn_frame = Frame(self, bg='#868B8E', width=200, height=800)
        btn_frame.pack(padx=0, pady=0, side='left')
        btn_frame.grid_propagate(0)

        # Frame Labels and buttons
        self.label_frame = Frame(self, bg='#145DA0', width=800, height=800)
        self.label_frame.pack(padx=0, pady=0, side='left')
        self.label_frame.grid_propagate(0)   

        #Buttons Opntions
        btn_agr = Button(btn_frame, text='Agregar Competidor/a', bd=5, command=lambda: master.switch_frame(Competidor))
        btn_agr.grid(row=0, column=0, padx=25, pady=80, ipadx=5, ipady= 5)

        btn_list = Button(btn_frame, text='Listar Competidores/as', bd=5, command=lambda: master.switch_frame(ListCompet))
        btn_list.grid(row=1, column=0, padx=25, pady=(0,40), ipadx=5, ipady= 5)

        btn_word = Button(btn_frame, text='Crear Llaves', bd=5, command=lambda: master.switch_frame(CreateKeys))
        btn_word.grid(row=2, column=0, padx=25, pady=30, ipadx=5, ipady= 5)

        self._createWidget()

 
    def _createWidget(self):

        self.optSex = StringVar()

        def only_numbers(char):
            return char.isdigit()

        validation = self.label_frame.register(only_numbers)

        #Nombre
        nmc= Label(self.label_frame,text="Nombre Competidor/a: ", bg='#145DA0', font=(10))
        nmc.grid(row=0, column=0, pady=50, padx=(50,0))

        self.nmcE = Entry(self.label_frame,font=(10))
        self.nmcE.grid( row=0, column=1, padx=20)

        #Sexo
        sxc = ttk.LabelFrame(self.label_frame,text="Sexo")
        sxc.grid(row=0, column=2, pady=(60,0) ,padx=(100,0))

        self.rbSex = ttk.Radiobutton(sxc, variable=self.optSex, value='M' ,text='Masculino')
        self.rbSex.grid(row=0, column=3, padx=20, pady=10)

        self.rbSex2 = ttk.Radiobutton(sxc, variable=self.optSex, value='F',text='Femenino')
        self.rbSex2.grid(row=1, column=3, padx=20, pady=10)

        #Edad
        edc = Label(self.label_frame, text="Edad Competidor/a: ", bg='#145DA0', font=(10))
        edc.grid(row=1, column=0, pady=0, padx=(50,0))

        self.edadE = Entry(self.label_frame,font=(10), width=8, validate='key', validatecommand=(validation, '%S'))
        self.edadE.grid( row=1, column=1, padx=20)
        
        #Peso
        psc = Label(self.label_frame, text="Peso Competidor/a:", bg='#145DA0', font=(10))
        psc.grid(row=2, column=0, pady=60, padx=25)

        self.pscE = Entry(self.label_frame, font=(10), width=8)
        self.pscE.grid(row=2, column=1, padx=20)
        
        #Color Cinturon
        cinturones = ['Blanco','Blanco-Amarillo','Amarillo','Amarillo-Verde',
                     'Verde','Verde-Azul','Azul','Azul-Rojo','Rojo','Rojo-Negro',
                     'Negro']

        
        clc = Label(self.label_frame,text="Color Cinturon: ", bg='#145DA0', font=(10))
        clc.grid(row=3,column=0,pady=(0,10))

        self.clcbx = ttk.Combobox(self.label_frame,values=cinturones, state='readonly',font=(10), width=14)
        self.clcbx.grid(row=3, column=1, padx=20)
        self.clcbx.current(0)

        #Button
        btn_sv = Button(self.label_frame,text="Guardar", bd=5, command=self.save_data)
        btn_sv.grid(row=4,column=2, padx=20,ipady=5,ipadx=20)

    # Save_Data es la funcion que va a guardar los datos en excel
    def save_data(self):
        
        clc = self.clcbx.get()
        
        try:
            nm = str(self.nmcE.get())
            if nm.isdigit() :
                raise ValueError()
            elif len(nm) == 0:
                raise ValueError()
        except ValueError:
            (showerror("Error","Debe Escribir el Nombre del/a competidor/a Correctamente"))
        try: 
            sex = self.optSex.get()
            if sex == '':
                raise ValueError()
        except ValueError:
            showerror("Error","Debe Seleccionar el sexo del/a competidor/a")
        try:
            edad = int(self.edadE.get())
        except ValueError:
            showerror("Error","Debe Escribir la edad del/a competidor/a")
        try:
            peso = float(self.pscE.get())
        except ValueError:
            showerror("Error","Debe Escribir el peso del/a competidor/a")
        else:
            
            datos = (
                ['Nombre','Edad','Peso','Sexo','Color Cinturon'],
                [nm, edad, peso, sex, clc]
            )
            print(datos)
            #print(f"El nombre es: {self.nmcE.get()},'Edad :{self.edadE.get()}','Peso: {self.pscE.get()}','Sexo: {self.optSex.get()}','Cinturon: {self.clcbx.get()}' ")

            self.nmcE.delete(0,END),
            self.edadE.delete(0,END),
            self.pscE.delete(0,END),
            self.optSex.set(None),
            self.clcbx.current(0)


        # datos = (
        #         ['Nombre','Edad','Peso','Sexo','Color Cinturon'],
        #         [nm, edad, peso, sex, clc]
        #      )
        

        # wsCompet = wbcompt.add_worksheet('Categoria_1')
        # row_number = 0
        #     col_number = 0

        #     for nombre, edad, peso, sexo, colc in datos:
        #         wsCompet.write(row_number, col_number, nombre),
        #         wsCompet.write(row_number, col_number+1, edad),
        #         wsCompet.write(row_number, col_number+2, peso),
        #         wsCompet.write(row_number, col_number+3, sexo),
        #         wsCompet.write(row_number,col_number+4, colc)

        #         row_number += 1

        #     wbcompt.close()
        
   
        


                


        
class ListCompet(Frame):
    def __init__(self,master):
        Frame.__init__(self,master)

        # Frame buttons left
        btn_frame = Frame(self, bg='#868B8E', width=200, height=800)
        btn_frame.pack(padx=0, pady=0, side='left')
        btn_frame.grid_propagate(0)

        # Frame Labels and buttons
        label_frame = Frame(self, bg='#A0E7E5', width=800, height=800)
        label_frame.pack(padx=0, pady=0, side='left')
        label_frame.grid_propagate(0)   

        #Buttons Opntions
        btn_agr = Button(btn_frame, text='Agregar Competidor/a', bd=5, command=lambda: master.switch_frame(Competidor))
        btn_agr.grid(row=0, column=0, padx=25, pady=80, ipadx=5, ipady= 5)

        btn_list = Button(btn_frame, text='Listar Competidores/as', bd=5, command=lambda: master.switch_frame(ListCompet))
        btn_list.grid(row=1, column=0, padx=25, pady=(0,40), ipadx=5, ipady= 5)

        btn_word = Button(btn_frame, text='Crear Llaves', bd=5, command=lambda: master.switch_frame(CreateKeys))
        btn_word.grid(row=2, column=0, padx=25, pady=30, ipadx=5, ipady= 5)


class CreateKeys(Frame):
    def __init__(self,master):
        Frame.__init__(self,master)
    
        # Frame buttons left
        btn_frame = Frame(self, bg='#868B8E', width=200, height=800)
        btn_frame.pack(padx=0, pady=0, side='left')
        btn_frame.grid_propagate(0)

        # Frame Labels and buttons
        label_frame = Frame(self, bg='#FFA384', width=800, height=800)
        label_frame.pack(padx=0, pady=0, side='left')
        label_frame.grid_propagate(0)   

        #Buttons Opntions
        btn_agr = Button(btn_frame, text='Agregar Competidor/a', bd=5, command=lambda: master.switch_frame(Competidor))
        btn_agr.grid(row=0, column=0, padx=25, pady=80, ipadx=5, ipady= 5)

        btn_list = Button(btn_frame, text='Listar Competidores/as', bd=5, command=lambda: master.switch_frame(ListCompet))
        btn_list.grid(row=1, column=0, padx=25, pady=(0,40), ipadx=5, ipady= 5)

        btn_word = Button(btn_frame, text='Crear Llaves', bd=5, command=lambda: master.switch_frame(CreateKeys))
        btn_word.grid(row=2, column=0, padx=25, pady=30, ipadx=5, ipady= 5)


#Cierre Pantalla
if __name__ == '__main__':
    app = App()
    app.mainloop()


