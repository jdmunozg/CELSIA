import tkinter as tk
from tkinter import PhotoImage
from tkinter import ttk
from tkinter import messagebox

class integracionAutomatizacion():

    root= None
    num = 0
    estado = True
    aux = 0
    tabla = None
    def __init__(self):
        self.root = tk.Tk()
        self.create_frame()
        

    def create_frame(self):
        
        self.root.title("Generador de Diagnosticos de ONT - Logistica Inversa")
        self.root.geometry("500x800")
        self.root.iconbitmap("logo.ico")
        # Cargar la imagen de fondo
        imagen_fondo = PhotoImage(file="fondo.png")
        # Crear un Label con la imagen de fondo
        label_fondo = tk.Label(self.root, image=imagen_fondo)
        label_fondo.place(relwidth=1, relheight=1)
        # Botón para ejecutar el Diagnostico Huawei
        btn_ejecutar = tk.Button(self.root, text="Generar informe", command=self.creacion_tabla, font=("Arial", 10, "bold"))

        btn_ejecutar.pack(pady=0, side="top")
        # Botón para salir
        btn_salir = tk.Button(self.root, text="Salir del programa", command=self.salir_frame,  font=("Arial", 10, "bold"))
        btn_salir.pack(pady=10, side="top")
        # Iniciar el bucle principal de la interfaz gráfica
        self.root.mainloop()

    def creacion_tabla(self):
        try:
            if self.num != 0:
                for item in self.tabla.get_children():
                    self.tabla.delete(item)
                self.num = 0
            frame_tabla = tk.Frame(self.root)
            frame_tabla.pack()
            # Create a table and insert data
            

            if self.estado == True :
                self.tabla = ttk.Treeview(self.root)

            

            self.tabla["columns"] = ("Atributo", "Valor")
            self.tabla.column("#0", stretch=tk.NO, width=0)  # Oculta la primera columna

            # Estilo para personalizar la tabla
            estilo = ttk.Style()
            estilo.configure("Treeview",
                     background="#FFFFFF",  # Color de fondo de la tabla
                     foreground="black",   # Color del texto
                     rowheight=25)          # Altura de la fila
            estilo.configure("Treeview.Heading", font=("Arial", 10, "bold"))
            self.tabla.tag_configure("oddrow", background="white")  # Color de fondo de filas impares
            self.tabla.tag_configure("evenrow", background="#F0F0F0")  # Color de fondo de filas pares
            for columna in self.tabla["columns"]:
                    self.tabla.column("Atributo", anchor=tk.W, width=120)
                    self.tabla.column("Valor", anchor=tk.W, width=200)
                    self.tabla.heading("Atributo", text="Atributo", anchor=tk.W)
                    self.tabla.heading("Valor", text="Valor", anchor=tk.W)
            self.tabla.insert("", tk.END, values=("Estado TX",self.aux))
            self.tabla.insert("", tk.END, values=("Estado RX", '''estado_rx'''))
            self.tabla.insert("", tk.END, values=("Estado Firmware", '''estado_firmware'''))
            self.tabla.insert("", tk.END, values=("Serial", '''serial_extraido'''))
            self.tabla.insert("", tk.END, values=("Modelo", '''modelo_extraido'''))
            self.tabla.insert("", tk.END, values=("Firmware", '''firmware_extraido'''))
            self.tabla.insert("", tk.END, values=("CPU", '''cpu_extraido'''))
            self.tabla.insert("", tk.END, values=("Memoria", '''mem_extraido'''))
            self.tabla.insert("", tk.END, values=("Potencia TX", '''tx_extraido'''))
            self.tabla.insert("", tk.END, values=("Potencia RX", '''rx_extraido'''))

            if(self.estado == True):
                self.tabla.pack(pady=0)
                self.estado = False
            print(self.tabla)


            self.num = self.num + 1
            self.aux = self.aux + 1

        
        except Exception as e:
            messagebox.showerror("Error", f"Se produjo un error: {str(e)}")

    def salir_frame(self):
        respuesta = messagebox.askyesno("Salir", "¿Estás seguro que quieres salir?")
        if respuesta:
            self.root.destroy()

x = integracionAutomatizacion()