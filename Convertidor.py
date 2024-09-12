import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
from tkinter import ttk
import pandas as pd
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import tkinter.font as tkFont  # Para obtener las fuentes disponibles

class ApliPrintApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Apli Print Online - Distribución Avanzada")
        self.root.geometry("1400x900")

        # Variables para almacenar datos
        self.df = None
        self.background_image = None
        self.background_photo = None
        self.dragging_label = None
        self.labels_positions = {}
        self.selected_columns = {}
        self.label_width = None
        self.label_height = None
        self.horizontal_margin = None
        self.vertical_margin = None
        self.text_size = 10  # Tamaño de texto predeterminado
        self.font_sizes = {}  # Tamaño de fuente para cada campo
        self.text_colors = {}  # Color del texto
        self.font_families = {}  # Familia de fuentes para cada campo
        self.label_distribution = []  # Distribución de etiquetas

        # Obtener las fuentes disponibles en el sistema
        self.available_fonts = list(tkFont.families())

        # Creación de la interfaz
        self.create_widgets()

    def create_widgets(self):
        # Crear un marco principal
        main_frame = ttk.Frame(self.root, padding="10 10 10 10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Crear un marco para los botones de carga y exportación
        button_frame = ttk.LabelFrame(main_frame, text="Acciones", padding="10 10 10 10")
        button_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        # Botones para seleccionar archivo y fondo
        self.load_file_btn = ttk.Button(button_frame, text="Cargar Archivo", command=self.load_file)
        self.load_file_btn.pack(side=tk.LEFT, padx=5)

        self.load_image_btn = ttk.Button(button_frame, text="Cargar Imagen de Fondo", command=self.load_background_image)
        self.load_image_btn.pack(side=tk.LEFT, padx=5)

        self.preview_btn = ttk.Button(button_frame, text="Vista Previa Completa", command=self.show_preview)
        self.preview_btn.pack(side=tk.LEFT, padx=5)

        self.export_pdf_btn = ttk.Button(button_frame, text="Exportar a PDF", command=self.show_export_options)
        self.export_pdf_btn.pack(side=tk.LEFT, padx=5)

        self.config_distribution_btn = ttk.Button(button_frame, text="Configurar Distribución", command=self.show_distribution_window)
        self.config_distribution_btn.pack(side=tk.LEFT, padx=5)

        # Crear un canvas para mostrar la imagen de fondo y los encabezados
        self.canvas = tk.Canvas(main_frame, bg="grey")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Controles para el tamaño, color y fuente del texto
        control_frame = ttk.LabelFrame(main_frame, text="Controles de Texto e Imagen", padding="10 10 10 10")
        control_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)

        # Control deslizante para el tamaño del texto
        self.text_size_slider = tk.Scale(control_frame, from_=8, to=40, orient=tk.HORIZONTAL, label="Tamaño del Texto", command=self.update_text_size)
        self.text_size_slider.pack(side=tk.LEFT, padx=10)

        # Botón para cambiar el color del texto
        self.color_button = ttk.Button(control_frame, text="Cambiar Color de Texto", command=self.change_text_color)
        self.color_button.pack(side=tk.LEFT, padx=10)

        # Control deslizante para el tamaño de la imagen de fondo
        self.image_size_slider = tk.Scale(control_frame, from_=0.5, to=2.0, resolution=0.1, orient=tk.HORIZONTAL, label="Tamaño de la Imagen de Fondo", command=self.update_background_size)
        self.image_size_slider.pack(side=tk.LEFT, padx=10)

        # Combobox para seleccionar la fuente del texto
        ttk.Label(control_frame, text="Fuente del Texto:").pack(side=tk.LEFT, padx=5)
        self.font_selector = ttk.Combobox(control_frame, values=self.available_fonts, state="readonly")
        self.font_selector.set("Arial")  # Fuente por defecto
        self.font_selector.pack(side=tk.LEFT, padx=10)
        self.font_selector.bind("<<ComboboxSelected>>", self.update_font_family)

    def load_file(self):
        # Abrir archivo Excel o txt y convertirlo a DataFrame
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("Text files", "*.txt")])
        if file_path:
            try:
                if file_path.endswith(".txt"):
                    self.df = pd.read_csv(file_path)
                else:
                    self.df = pd.read_excel(file_path)
                messagebox.showinfo("Éxito", "Archivo cargado exitosamente")
                self.show_column_selection()  # Mostrar la selección de columnas
            except Exception as e:
                messagebox.showerror("Error", f"Error cargando el archivo: {e}")

    def show_column_selection(self):
        # Crear una nueva ventana para seleccionar los encabezados que se desean arrastrar
        self.selection_window = tk.Toplevel(self.root)
        self.selection_window.title("Seleccionar Columnas")
        self.selection_window.geometry("400x400")

        self.selected_columns.clear()

        # Crear checkboxes para cada encabezado y campos para seleccionar el tamaño de fuente
        for column in self.df.columns:
            var = tk.BooleanVar(value=True)
            frame = ttk.Frame(self.selection_window)
            frame.pack(fill=tk.X, padx=5, pady=5)

            # Checkbox para incluir el campo
            cb = tk.Checkbutton(frame, text=column, variable=var)
            cb.pack(side=tk.LEFT, padx=5)
            self.selected_columns[column] = var

            # Campo para ajustar el tamaño de fuente
            font_size_var = tk.StringVar(value="12")  # Tamaño de fuente por defecto
            font_size_entry = ttk.Entry(frame, width=5, textvariable=font_size_var)
            font_size_entry.pack(side=tk.RIGHT, padx=5)
            self.font_sizes[column] = font_size_var
            self.text_colors[column] = "black"  # Color de texto por defecto
            self.font_families[column] = "Arial"  # Fuente por defecto

        # Botón para confirmar la selección y cerrar la ventana
        confirm_btn = ttk.Button(self.selection_window, text="Confirmar", command=self.confirm_column_selection)
        confirm_btn.pack(pady=10)

    def confirm_column_selection(self):
        # Cerrar la ventana de selección de columnas y mostrar los encabezados seleccionados en el canvas
        self.selection_window.destroy()
        self.show_headers()

    def show_headers(self):
        # Limpiar el canvas antes de agregar nuevos encabezados
        self.canvas.delete("all")
        
        # Si hay una imagen de fondo, volver a cargarla
        if self.background_photo:
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.background_photo)
        
        # Comprobar si el DataFrame se ha cargado correctamente
        if self.df is not None and not self.df.empty:
            # Mostrar solo los encabezados seleccionados en el canvas
            for idx, column in enumerate(self.df.columns):
                if self.selected_columns[column].get():
                    # Crear el texto del encabezado en una posición inicial arbitraria
                    font_size = int(self.font_sizes[column].get())
                    font_family = self.font_families[column]
                    text_id = self.canvas.create_text(100, 50 + (idx * 30), text=column, fill=self.text_colors[column], font=(font_family, font_size), tags="draggable")

                    # Guardar el ID del texto en la posición inicial
                    self.labels_positions[column] = (100, 50 + (idx * 30))

                    # Asignar eventos de arrastre al texto
                    self.canvas.tag_bind(text_id, "<Button-1>", self.start_drag)
                    self.canvas.tag_bind(text_id, "<B1-Motion>", self.do_drag)
                    self.canvas.tag_bind(text_id, "<ButtonRelease-1>", self.end_drag)
        else:
            messagebox.showwarning("Advertencia", "El DataFrame está vacío o no se cargó correctamente.")

    def update_text_size(self, value):
        # Actualizar el tamaño del texto seleccionado
        for column in self.selected_columns:
            if self.selected_columns[column].get():
                font_size = int(value)
                self.font_sizes[column].set(font_size)
        self.show_headers()

    def change_text_color(self):
        # Cambiar el color del texto seleccionado
        color = colorchooser.askcolor()[1]
        if color:
            for column in self.selected_columns:
                if self.selected_columns[column].get():
                    self.text_colors[column] = color
            self.show_headers()

    def update_font_family(self, event):
        # Actualizar la fuente del texto seleccionado
        selected_font = self.font_selector.get()
        for column in self.selected_columns:
            if self.selected_columns[column].get():
                self.font_families[column] = selected_font
        self.show_headers()

    def update_background_size(self, value):
        # Cambiar el tamaño de la imagen de fondo
        if self.background_image:
            new_width = int(self.background_image.width * float(value))
            new_height = int(self.background_image.height * float(value))
            resized_image = self.background_image.resize((new_width, new_height), Image.LANCZOS)
            self.background_photo = ImageTk.PhotoImage(resized_image)
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.background_photo)
            self.show_headers()

    # Métodos para arrastrar los textos en el canvas
    def start_drag(self, event):
        self.dragging_label = self.canvas.find_withtag("current")
        self.drag_start_x = event.x
        self.drag_start_y = event.y

    def do_drag(self, event):
        if self.dragging_label:
            dx = event.x - self.drag_start_x
            dy = event.y - self.drag_start_y
            self.canvas.move(self.dragging_label, dx, dy)
            self.drag_start_x = event.x
            self.drag_start_y = event.y

    def end_drag(self, event):
        if self.dragging_label:
            x, y = self.canvas.coords(self.dragging_label)
            column_name = self.canvas.itemcget(self.dragging_label, "text")
            self.labels_positions[column_name] = (x, y)
            self.dragging_label = None

    def load_background_image(self):
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg")])
        if file_path:
            try:
                self.background_image = Image.open(file_path)
                self.background_image_path = file_path
                self.background_photo = ImageTk.PhotoImage(self.background_image)
                self.canvas.create_image(0, 0, anchor=tk.NW, image=self.background_photo)
                self.show_headers()
            except Exception as e:
                messagebox.showerror("Error", f"Error cargando la imagen: {e}")

    def show_distribution_window(self):
        # Ventana emergente para seleccionar la distribución de etiquetas por sucursal
        self.distribution_window = tk.Toplevel(self.root)
        self.distribution_window.title("Configurar Distribución de Etiquetas")
        self.distribution_window.geometry("600x500")

        # Grid que simula la hoja A4
        grid_frame = ttk.Frame(self.distribution_window)
        grid_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Variables para marcar las celdas
        self.grid_cells = []
        self.label_distribution.clear()

        num_rows = 10
        num_cols = 7
        for r in range(num_rows):
            row = []
            for c in range(num_cols):
                cell_var = tk.BooleanVar()
                cell_btn = tk.Checkbutton(grid_frame, variable=cell_var, onvalue=True, offvalue=False)
                cell_btn.grid(row=r, column=c, ipadx=10, ipady=5)
                row.append(cell_var)
            self.grid_cells.append(row)

        # Botón para confirmar la distribución
        confirm_btn = ttk.Button(self.distribution_window, text="Confirmar Distribución", command=self.confirm_distribution)
        confirm_btn.pack(pady=10)

    def confirm_distribution(self):
        # Almacenar la distribución de las etiquetas
        self.label_distribution = [[cell.get() for cell in row] for row in self.grid_cells]
        self.distribution_window.destroy()
        messagebox.showinfo("Distribución Guardada", "Distribución de etiquetas guardada exitosamente.")

    def show_export_options(self):
        # Crear una ventana emergente para seleccionar las opciones de exportación
        self.export_window = tk.Toplevel(self.root)
        self.export_window.title("Opciones de Exportación")
        self.export_window.geometry("300x400")

        # Crear campos para ingresar dimensiones de las etiquetas
        ttk.Label(self.export_window, text="Ancho de etiqueta (mm):").pack(pady=5)
        self.label_width_entry = ttk.Entry(self.export_window)
        self.label_width_entry.insert(0, "50")  # Valor por defecto
        self.label_width_entry.pack(pady=5)

        ttk.Label(self.export_window, text="Altura de etiqueta (mm):").pack(pady=5)
        self.label_height_entry = ttk.Entry(self.export_window)
        self.label_height_entry.insert(0, "30")  # Valor por defecto
        self.label_height_entry.pack(pady=5)

        # Crear campos para márgenes
        ttk.Label(self.export_window, text="Margen horizontal (mm):").pack(pady=5)
        self.horizontal_margin_entry = ttk.Entry(self.export_window)
        self.horizontal_margin_entry.insert(0, "5")  # Valor por defecto
        self.horizontal_margin_entry.pack(pady=5)

        ttk.Label(self.export_window, text="Margen vertical (mm):").pack(pady=5)
        self.vertical_margin_entry = ttk.Entry(self.export_window)
        self.vertical_margin_entry.insert(0, "5")  # Valor por defecto
        self.vertical_margin_entry.pack(pady=5)

        # Crear campo para el tamaño del texto por defecto
        ttk.Label(self.export_window, text="Tamaño del texto por defecto (puntos):").pack(pady=5)
        self.text_size_entry = ttk.Entry(self.export_window)
        self.text_size_entry.insert(0, "10")  # Valor por defecto
        self.text_size_entry.pack(pady=5)

        # Crear checkbox para duplicar etiquetas
        self.duplicate_labels_var = tk.BooleanVar()
        ttk.Checkbutton(self.export_window, text="Duplicar etiquetas en toda la hoja", variable=self.duplicate_labels_var).pack(pady=10)

        # Botón para confirmar y generar el PDF
        confirm_btn = ttk.Button(self.export_window, text="Generar PDF", command=self.export_to_pdf)
        confirm_btn.pack(pady=10)

    def show_preview(self):
        # Mostrar una vista previa avanzada de cómo se verán las etiquetas
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Vista Previa de Etiquetas")
        preview_window.geometry("800x600")
        preview_canvas = tk.Canvas(preview_window, bg="white")
        preview_canvas.pack(fill=tk.BOTH, expand=True)

        # Simular el dibujo de etiquetas para la vista previa
        if self.background_photo:
            preview_canvas.create_image(0, 0, anchor=tk.NW, image=self.background_photo)

        # Dibujar encabezados con los tamaños seleccionados
        for column, (x, y) in self.labels_positions.items():
            if self.selected_columns[column].get():
                font_size = int(self.font_sizes[column].get())
                font_family = self.font_families[column]
                preview_canvas.create_text(x, y, text=column, font=(font_family, font_size), fill=self.text_colors[column])

    def export_to_pdf(self):
        # Obtener las dimensiones de las etiquetas ingresadas
        try:
            self.label_width = float(self.label_width_entry.get()) * 2.83465  # Convertir de mm a puntos (1 mm = 2.83465 puntos)
            self.label_height = float(self.label_height_entry.get()) * 2.83465
            self.horizontal_margin = float(self.horizontal_margin_entry.get()) * 2.83465
            self.vertical_margin = float(self.vertical_margin_entry.get()) * 2.83465
            self.text_size = int(self.text_size_entry.get())  # Tamaño del texto en puntos
        except ValueError:
            messagebox.showerror("Error", "Por favor, ingresa dimensiones válidas.")
            return

        if self.df is None or self.background_image is None:
            messagebox.showwarning("Advertencia", "Por favor, carga un archivo y una imagen de fondo antes de exportar.")
            return

        # Seleccionar carpeta de salida para guardar el PDF
        output_dir = filedialog.askdirectory()
        if output_dir:
            try:
                pdf_file_path = f"{output_dir}/etiquetas.pdf"
                c = canvas.Canvas(pdf_file_path, pagesize=A4)

                # Tamaño del área de la hoja A4
                pdf_width, pdf_height = A4
                num_columns = int((pdf_width + self.horizontal_margin) // (self.label_width + self.horizontal_margin))
                num_rows = int((pdf_height + self.vertical_margin) // (self.label_height + self.vertical_margin))

                # Generar las etiquetas
                current_row = 0
                current_col = 0

                for index, row in self.df.iterrows():
                    # Dibujar la imagen de fondo para cada etiqueta
                    x_pos = current_col * (self.label_width + self.horizontal_margin)
                    y_pos = pdf_height - ((current_row + 1) * (self.label_height + self.vertical_margin))

                    c.drawImage(self.background_image_path, x_pos, y_pos, width=self.label_width, height=self.label_height)

                    # Dibujar los datos de la fila actual en las posiciones guardadas
                    for column, (x, y) in self.labels_positions.items():
                        if self.selected_columns[column].get():
                            # Calcular las posiciones relativas dentro de la etiqueta
                            rel_x = x_pos + (x / self.canvas.winfo_width()) * self.label_width
                            rel_y = y_pos + (y / self.canvas.winfo_height()) * self.label_height
                            font_size = int(self.font_sizes[column].get())
                            font_family = self.font_families[column]
                            c.setFont(font_family, font_size)  # Establecer el tamaño y la fuente
                            c.drawString(rel_x, rel_y, str(row[column]))

                    current_col += 1
                    if current_col >= num_columns:
                        current_col = 0
                        current_row += 1

                    # Si se llena una hoja, crear una nueva
                    if current_row >= num_rows:
                        c.showPage()  # Crear una nueva página
                        current_row = 0

                    # Si se seleccionó la opción de duplicar etiquetas, llenar toda la hoja con la misma
                    if self.duplicate_labels_var.get():
                        for _ in range(num_rows * num_columns - 1):  # Rellenar con copias
                            x_pos = current_col * (self.label_width + self.horizontal_margin)
                            y_pos = pdf_height - ((current_row + 1) * (self.label_height + self.vertical_margin))
                            c.drawImage(self.background_image_path, x_pos, y_pos, width=self.label_width, height=self.label_height)
                            current_col += 1
                            if current_col >= num_columns:
                                current_col = 0
                                current_row += 1
                            if current_row >= num_rows:
                                c.showPage()
                                current_row = 0

                c.save()
                messagebox.showinfo("Éxito", f"PDF generado correctamente: {pdf_file_path}")
                self.export_window.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Error generando el PDF: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ApliPrintApp(root)
    root.mainloop()
