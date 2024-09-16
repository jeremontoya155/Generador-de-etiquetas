import pandas as pd
from tkinter import HORIZONTAL, Tk, filedialog, Canvas, Button, Scale, Listbox, messagebox, Toplevel, Scrollbar, Label, Entry
from tkinter import Frame
from PIL import Image, ImageTk, ImageDraw
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os

class EtiquetaApp:
    def __init__(self, ventana):
        self.ventana = ventana
        self.ventana.geometry("900x800")  # Tamaño de ventana ajustado
        self.imagen_tk = None
        self.df = None
        self.posiciones_texto = {}
        self.etiquetas = {}
        self.tamaños_texto = {}
        self.columnas_seleccionadas = []
        self.tamaño_imagen = 1.0  # Factor de escala de la imagen
        self.imagen_fondo = None
        self.ancho_etiqueta_custom = 0
        self.alto_etiqueta_custom = 0
        self.margen_x_custom = 0
        self.margen_y_custom = 0

        # Frame principal para contener la imagen y los controles a la derecha
        self.main_frame = Frame(ventana)
        self.main_frame.pack(fill="both", expand=True)

        # Frame para la imagen (lado izquierdo)
        self.canvas_frame = Frame(self.main_frame, width=600, height=600)  # Ajustar el espacio para la imagen
        self.canvas_frame.pack(side="left", fill="both", expand=True)

        self.canvas = Canvas(self.canvas_frame, width=580, height=580)  # Canvas ajustado a la ventana
        self.canvas.pack(side="left", fill="both", expand=True)

        # Añadir scrollbars al canvas para que puedas desplazarte si la imagen es más grande
        self.scroll_x = Scrollbar(self.canvas_frame, orient="horizontal", command=self.canvas.xview)
        self.scroll_x.pack(side="bottom", fill="x")
        self.scroll_y = Scrollbar(self.canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scroll_y.pack(side="right", fill="y")
        self.canvas.configure(xscrollcommand=self.scroll_x.set, yscrollcommand=self.scroll_y.set)

        # Frame para los controles (lado derecho)
        self.control_frame = Frame(self.main_frame, width=300)  # Ajustar el espacio para los controles
        self.control_frame.pack(side="right", fill="y")

        # Scroll para los controles si no caben en la pantalla
        self.control_scroll_y = Scrollbar(self.control_frame, orient="vertical")
        self.control_scroll_y.pack(side="right", fill="y")

        self.control_canvas = Canvas(self.control_frame, yscrollcommand=self.control_scroll_y.set)
        self.control_canvas.pack(side="left", fill="both", expand=True)
        self.control_scroll_y.config(command=self.control_canvas.yview)

        self.control_inner_frame = Frame(self.control_canvas)
        self.control_canvas.create_window((0, 0), window=self.control_inner_frame, anchor="nw")
        self.control_inner_frame.bind("<Configure>", lambda e: self.control_canvas.configure(scrollregion=self.control_canvas.bbox("all")))

    # Ajustar el tamaño del canvas cuando el contenido cambie
    def on_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def cargar_excel(self):
        ruta_excel = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if ruta_excel:
            try:
                self.df = pd.read_excel(ruta_excel)
                return self.df
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar el archivo Excel: {str(e)}")
        return None

    def seleccionar_columnas(self):
        def seleccionar():
            seleccionadas = [self.df.columns[i] for i in listbox.curselection()]
            if len(seleccionadas) < 3:
                messagebox.showwarning("Advertencia", "Debes seleccionar al menos 3 columnas.")
            else:
                top.destroy()
                self.columnas_seleccionadas = seleccionadas
                self.editar_etiquetas()

        top = Toplevel(self.ventana)
        top.title("Selecciona columnas")
        listbox = Listbox(top, selectmode="multiple")
        for col in self.df.columns:
            listbox.insert("end", col)
        listbox.pack()

        boton_seleccionar = Button(top, text="Seleccionar", command=seleccionar)
        boton_seleccionar.pack()
        top.mainloop()

    def cargar_imagen(self):
        ruta_imagen = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.png")])
        if ruta_imagen:
            try:
                self.imagen_fondo = Image.open(ruta_imagen)
                return self.imagen_fondo
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar la imagen: {str(e)}")
        return None

    def mover_texto(self, event, etiqueta):
        self.posiciones_texto[etiqueta]['x'] = event.x
        self.posiciones_texto[etiqueta]['y'] = event.y
        self.canvas.coords(self.etiquetas[etiqueta], event.x, event.y)

    def editar_etiquetas(self):
        self.cargar_imagen()
        if not self.imagen_fondo:
            return

        # Escalar la imagen para que no sea más grande que el área disponible
        ancho_max = 580
        alto_max = 580
        self.imagen_fondo.thumbnail((ancho_max, alto_max), Image.Resampling.LANCZOS)

        # Redimensionar imagen de fondo
        self.imagen_redimensionada = self.imagen_fondo
        self.imagen_tk = ImageTk.PhotoImage(self.imagen_redimensionada)
        
        # Ajustar el canvas al tamaño de la imagen
        self.canvas.config(width=self.imagen_redimensionada.width, height=self.imagen_redimensionada.height)
        self.canvas.create_image(0, 0, anchor="nw", image=self.imagen_tk)

        # Crear etiquetas de texto para las columnas seleccionadas
        for i, columna in enumerate(self.columnas_seleccionadas):
            texto = self.df[columna][0]  # Ejemplo con la primera fila
            self.posiciones_texto[columna] = {'x': 10, 'y': 30 * (i + 1)}
            self.tamaños_texto[columna] = 20

            self.etiquetas[columna] = self.canvas.create_text(
                self.posiciones_texto[columna]['x'], self.posiciones_texto[columna]['y'],
                text=texto, font=("Arial", self.tamaños_texto[columna]), anchor="nw"
            )

            # Habilitar drag & drop para mover los textos
            self.canvas.tag_bind(self.etiquetas[columna], '<B1-Motion>',
                                 lambda event, col=columna: self.mover_texto(event, col))

        # Colocar controles en el lado derecho
        self.colocar_controles()

    def colocar_controles(self):
        # Limpiar el frame de controles
        for widget in self.control_inner_frame.winfo_children():
            widget.destroy()

        # Slider para cambiar el tamaño del texto
        for columna in self.columnas_seleccionadas:
            slider_tamaño = Scale(self.control_inner_frame, from_=10, to=50, orient=HORIZONTAL, label="Tamaño de {}".format(columna))

            slider_tamaño.set(self.tamaños_texto[columna])
            slider_tamaño.pack(fill="x")

            # Actualización del tamaño en tiempo real
            def actualizar_tamaño(valor, col=columna):
                self.tamaños_texto[col] = int(valor)
                self.canvas.itemconfig(self.etiquetas[col], font=("Arial", self.tamaños_texto[col]))

            slider_tamaño.config(command=actualizar_tamaño)

        # Opción para definir dimensiones personalizadas de la etiqueta (en milímetros)
        Label(self.control_inner_frame, text="Dimensiones personalizadas de la etiqueta (mm)").pack()

        Label(self.control_inner_frame, text="Ancho (mm)").pack()
        ancho_entry = Entry(self.control_inner_frame)
        ancho_entry.pack()

        Label(self.control_inner_frame, text="Alto (mm)").pack()
        alto_entry = Entry(self.control_inner_frame)
        alto_entry.pack()

        # Opción para definir márgenes (en milímetros)
        Label(self.control_inner_frame, text="Margen entre etiquetas (mm)").pack()

        Label(self.control_inner_frame, text="Margen horizontal (mm)").pack()
        margen_x_entry = Entry(self.control_inner_frame)
        margen_x_entry.pack()

        Label(self.control_inner_frame, text="Margen vertical (mm)").pack()
        margen_y_entry = Entry(self.control_inner_frame)
        margen_y_entry.pack()

        def establecer_dimensiones_y_margenes():
            try:
                self.ancho_etiqueta_custom = int(ancho_entry.get()) * 2.83465  # Convertir mm a puntos
                self.alto_etiqueta_custom = int(alto_entry.get()) * 2.83465  # Convertir mm a puntos
                self.margen_x_custom = int(margen_x_entry.get()) * 2.83465  # Convertir mm a puntos
                self.margen_y_custom = int(margen_y_entry.get()) * 2.83465  # Convertir mm a puntos
                messagebox.showinfo("Configuración", "Dimensiones y márgenes establecidos.")
            except ValueError:
                messagebox.showerror("Error", "Por favor ingresa valores numéricos válidos para las dimensiones y márgenes.")

        boton_dimensiones = Button(self.control_inner_frame, text="Establecer dimensiones y márgenes", command=establecer_dimensiones_y_margenes)
        boton_dimensiones.pack()

        # Botón para exportar a PDF
        boton_exportar = Button(self.control_inner_frame, text="Exportar a PDF", command=self.exportar_pdf)
        boton_exportar.pack(fill="x")

    def exportar_pdf(self):
        c = canvas.Canvas("etiquetas.pdf", pagesize=A4)
        ancho_hoja, alto_hoja = A4  # Tamaño de una hoja A4 en puntos

        # Usar las dimensiones personalizadas para cada etiqueta
        ancho_etiqueta = self.ancho_etiqueta_custom
        alto_etiqueta = self.alto_etiqueta_custom

        # Usar los márgenes personalizados
        margen_x = self.margen_x_custom
        margen_y = self.margen_y_custom

        x_offset = margen_x
        y_offset = alto_hoja - alto_etiqueta - margen_y

        for i, row in self.df.iterrows():
            # Dibujar cada etiqueta
            imagen_pdf = self.generar_imagen_pdf(row)
            c.drawImage(imagen_pdf, x_offset, y_offset, width=ancho_etiqueta, height=alto_etiqueta)

            # Ajustar la posición para la siguiente etiqueta
            x_offset += ancho_etiqueta + margen_x
            if x_offset + ancho_etiqueta > ancho_hoja:  # Si ya no cabe horizontalmente
                x_offset = margen_x
                y_offset -= alto_etiqueta + margen_y

            if y_offset < margen_y:  # Si no cabe más en la página, agregar una nueva página
                c.showPage()
                x_offset = margen_x
                y_offset = alto_hoja - alto_etiqueta - margen_y

        c.save()
        messagebox.showinfo("Éxito", "El archivo PDF se ha generado correctamente.")

    def generar_imagen_pdf(self, fila):
        imagen_etiqueta = self.imagen_redimensionada.copy()
        draw = ImageDraw.Draw(imagen_etiqueta)

        # Dibujar los textos en las posiciones definidas
        for columna in self.columnas_seleccionadas:
            texto = str(fila[columna])
            posicion = (self.posiciones_texto[columna]['x'], self.posiciones_texto[columna]['y'])
            tamaño = self.tamaños_texto[columna]
            draw.text(posicion, texto, fill="black")

        ruta_imagen_pdf = f"etiqueta_{fila.name}.png"
        imagen_etiqueta.save(ruta_imagen_pdf)
        return ruta_imagen_pdf

def iniciar_programa():
    ventana = Tk()
    ventana.geometry("900x800")  # Tamaño fijo de la ventana
    app = EtiquetaApp(ventana)

    # Iniciar el flujo del programa
    df = app.cargar_excel()
    if df is not None:
        app.seleccionar_columnas()

    ventana.mainloop()


if __name__ == "__main__":
    iniciar_programa()
