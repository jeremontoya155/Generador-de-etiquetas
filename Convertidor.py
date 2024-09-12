import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

class ApliPrintApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Apli Print Online - Distribución Avanzada")
        self.root.geometry("1400x900")

        # Configuración del estilo general
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", background="#ccc")
        style.configure("TLabel", padding=5, font=("Helvetica", 10))
        style.configure("TEntry", padding=5)
        style.configure("TFrame", background="#f5f5f5")
        style.configure("TLabelFrame", background="#f5f5f5", font=("Helvetica", 12, "bold"))

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
        self.label_distribution = []  # Distribución de etiquetas

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

        # Control de tamaño de fuente
        self.font_size_label = ttk.Label(button_frame, text="Tamaño de fuente:")
        self.font_size_label.pack(side=tk.LEFT, padx=5)

        self.font_size_entry = ttk.Entry(button_frame, width=5)
        self.font_size_entry.insert(0, "12")  # Tamaño de fuente predeterminado
        self.font_size_entry.pack(side=tk.LEFT, padx=5)

        self.update_font_btn = ttk.Button(button_frame, text="Actualizar Fuente", command=self.update_font_size)
        self.update_font_btn.pack(side=tk.LEFT, padx=5)

        # Crear un canvas para mostrar la imagen de fondo y los encabezados
        self.canvas = tk.Canvas(main_frame, bg="grey")
        self.canvas.pack(fill=tk.BOTH, expand=True)

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
                    text_id = self.canvas.create_text(100, 50 + (idx * 30), text=column, fill="black", font=("Arial", int(self.font_sizes[column].get())), tags="draggable")

                    # Guardar el ID del texto en la posición inicial
                    self.labels_positions[column] = (100, 50 + (idx * 30))

                    # Asignar eventos de arrastre al texto
                    self.canvas.tag_bind(text_id, "<Button-1>", self.start_drag)
                    self.canvas.tag_bind(text_id, "<B1-Motion>", self.do_drag)
                    self.canvas.tag_bind(text_id, "<ButtonRelease-1>", self.end_drag)
        else:
            messagebox.showwarning("Advertencia", "El DataFrame está vacío o no se cargó correctamente.")

    def start_drag(self, event):
        # Iniciar arrastre del encabezado
        self.dragging_label = self.canvas.find_withtag("current")
        self.drag_start_x = event.x
        self.drag_start_y = event.y

    def do_drag(self, event):
        # Mover el encabezado arrastrado en el canvas
        if self.dragging_label:
            # Calcular el desplazamiento
            dx = event.x - self.drag_start_x
            dy = event.y - self.drag_start_y

            # Mover el texto dentro del canvas
            self.canvas.move(self.dragging_label, dx, dy)

            # Actualizar la posición inicial del cursor para el próximo movimiento
            self.drag_start_x = event.x
            self.drag_start_y = event.y

    def end_drag(self, event):
        # Guardar la posición donde se soltó el encabezado
        if self.dragging_label:
            # Obtener las nuevas coordenadas del texto
            x, y = self.canvas.coords(self.dragging_label)
            # Guardar la nueva posición del encabezado
            column_name = self.canvas.itemcget(self.dragging_label, "text")
            self.labels_positions[column_name] = (x, y)
            self.dragging_label = None

    def load_background_image(self):
        # Cargar una imagen de fondo
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg")])
        if file_path:
            try:
                self.background_image = Image.open(file_path)
                self.background_image_path = file_path  # Guardar la ruta para exportar
                self.background_image = self.background_image.resize((self.canvas.winfo_width(), self.canvas.winfo_height()), Image.LANCZOS)
                self.background_photo = ImageTk.PhotoImage(self.background_image)
                self.canvas.create_image(0, 0, anchor=tk.NW, image=self.background_photo)
                self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))  # Ajustar el área de desplazamiento del canvas
                
                # Mostrar nuevamente los encabezados para que se dibujen sobre la imagen
                self.show_headers()
            except Exception as e:
                messagebox.showerror("Error", f"Error cargando la imagen: {e}")

    def update_font_size(self):
        # Actualizar el tamaño de fuente para los encabezados visibles
        try:
            new_font_size = int(self.font_size_entry.get())
            for column in self.selected_columns:
                if self.selected_columns[column].get():
                    self.font_sizes[column].set(str(new_font_size))  # Cambiar el tamaño de fuente
            self.show_headers()  # Redibujar encabezados con el nuevo tamaño
        except ValueError:
            messagebox.showerror("Error", "Introduce un tamaño de fuente válido.")

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
                preview_canvas.create_text(x, y, text=column, font=("Helvetica", font_size), fill="black")

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
                            c.setFont("Helvetica", font_size)  # Establecer el tamaño de la fuente
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
