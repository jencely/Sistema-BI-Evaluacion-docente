import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys
import os
from datetime import datetime
from evaluacion_docente import EvaluacionDocenteSystem


class EvaluacionDocenteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Evaluación Docente - UNIBE")
        self.root.geometry("900x600")
        
        # Configurar el estilo
        self.style = ttk.Style()
        self.style.configure('Header.TLabel', font=('Arial', 24, 'bold'))
        self.style.configure('SubHeader.TLabel', font=('Arial', 12))
        self.style.configure('Success.TLabel', foreground='green')
        self.style.configure('Error.TLabel', foreground='red')
        
        # Inicializar el sistema
        self.sistema = EvaluacionDocenteSystem()
        
        # Crear la interfaz
        self.create_widgets()

    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, pady=20, sticky=(tk.W, tk.E))
        
        header_label = ttk.Label(
            header_frame,
            text="Sistema de Evaluación Docente",
            style='Header.TLabel'
        )
        header_label.grid(row=0, column=0, pady=5)
        
        subheader_label = ttk.Label(
            header_frame,
            text="Universidad Iberoamericana del Ecuador",
            style='SubHeader.TLabel'
        )
        subheader_label.grid(row=1, column=0)
        
        # Frame de acciones
        actions_frame = ttk.LabelFrame(main_frame, text="Acciones", padding="10")
        actions_frame.grid(row=1, column=0, pady=10, sticky=(tk.W, tk.E))
        
        # Botones
        ttk.Button(
            actions_frame,
            text="Seleccionar Archivos",
            command=self.procesar_archivos,
            width=20
        ).grid(row=0, column=0, padx=5, pady=5)
        
        ttk.Button(
            actions_frame,
            text="Descargar Plantilla",
            command=self.descargar_plantilla,
            width=20
        ).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(
            actions_frame,
            text="Ver Manual",
            command=self.mostrar_manual,
            width=20
        ).grid(row=0, column=2, padx=5, pady=5)

        ttk.Button(
            actions_frame,
            text="Ver Categorías",
            command=self.mostrar_categorias,
            width=20
        ).grid(row=0, column=3, padx=5, pady=5)
        
        # Frame de log
        log_frame = ttk.LabelFrame(main_frame, text="Registro de Operaciones", padding="10")
        log_frame.grid(row=2, column=0, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.rowconfigure(2, weight=1)
        
        # Text widget para log
        self.log_text = tk.Text(log_frame, height=15, width=80)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar para log
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text['yscrollcommand'] = scrollbar.set
        
        # Frame de estado
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=3, column=0, pady=5, sticky=(tk.W, tk.E))
        
        self.status_label = ttk.Label(
            status_frame,
            text="Listo para procesar archivos",
            style='Success.TLabel'
        )
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
        # Fecha actual
        fecha_label = ttk.Label(
            status_frame,
            text=datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        )
        fecha_label.grid(row=0, column=1, sticky=tk.E)

    def log_message(self, message: str, level: str = "INFO"):
        """Agregar mensaje al log con timestamp"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {level}: {message}\n")
        self.log_text.see(tk.END)
        self.root.update()

    def procesar_archivos(self):
        """Procesar archivos de evaluación"""
        try:
            archivos = filedialog.askopenfilenames(
                title='Selecciona los archivos de evaluación',
                filetypes=[
                    ('Excel con macros', '*.xlsm'),
                    ('Excel files', '*.xlsx'),
                    ('Todos los archivos', '*.*')
                ]
            )
            
            if not archivos:
                return
                
            self.status_label.config(text="Procesando archivos...", style='')
            self.log_message(f"Iniciando procesamiento de {len(archivos)} archivos")
            
            archivos_procesados = 0
            for archivo in archivos:
                try:
                    self.log_message(f"Procesando: {os.path.basename(archivo)}")
                    if self.sistema.procesar_archivo_excel(archivo):
                        archivos_procesados += 1
                        self.log_message(f"Archivo procesado exitosamente: {os.path.basename(archivo)}")
                    else:
                        self.log_message(f"Error procesando archivo: {os.path.basename(archivo)}", "ERROR")
                except Exception as e:
                    self.log_message(f"Error en archivo {os.path.basename(archivo)}: {str(e)}", "ERROR")
            
            self.status_label.config(
                text=f"Se procesaron {archivos_procesados} de {len(archivos)} archivos",
                style='Success.TLabel' if archivos_procesados == len(archivos) else 'Error.TLabel'
            )
            
            messagebox.showinfo(
                "Proceso Completado",
                f"Se procesaron {archivos_procesados} de {len(archivos)} archivos correctamente"
            )
            
        except Exception as e:
            self.log_message(f"Error: {str(e)}", "ERROR")
            self.status_label.config(text="Error en el procesamiento", style='Error.TLabel')
            messagebox.showerror("Error", f"Error en el procesamiento: {str(e)}")

    def descargar_plantilla(self):
        """Abrir carpeta con la plantilla"""
        try:
            ruta_plantilla = os.path.join(
                os.path.dirname(os.path.dirname(__file__)),
                "resources",
                "templates",
                "plantilla_evaluacion.xlsm"
            )
            
            if os.path.exists(ruta_plantilla):
                os.startfile(os.path.dirname(ruta_plantilla))
                self.log_message("Carpeta de plantillas abierta")
            else:
                raise FileNotFoundError("No se encontró la plantilla")
                
        except Exception as e:
            self.log_message(f"Error accediendo a la plantilla: {str(e)}", "ERROR")
            messagebox.showerror("Error", "No se pudo acceder a la plantilla")

    def mostrar_manual(self):
        """Abrir manual de usuario"""
        try:
            ruta_manual = os.path.join(
                os.path.dirname(os.path.dirname(__file__)),
                "docs",
                "manual_usuario.md"
            )
            
            if os.path.exists(ruta_manual):
                os.startfile(ruta_manual)
                self.log_message("Manual de usuario abierto")
            else:
                raise FileNotFoundError("No se encontró el manual")
                
        except Exception as e:
            self.log_message(f"Error accediendo al manual: {str(e)}", "ERROR")
            messagebox.showerror("Error", "No se pudo acceder al manual")

    def mostrar_categorias(self):
        """Muestra una ventana con las categorías disponibles"""
        try:
            ventana = tk.Toplevel(self.root)
            ventana.title("Categorías de Evaluación")
            ventana.geometry("800x600")
            
            # Frame principal
            main_frame = ttk.Frame(ventana, padding="10")
            main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            
            # Título
            ttk.Label(
                main_frame,
                text="Categorías y sus Ítems de Evaluación",
                style='SubHeader.TLabel'
            ).grid(row=0, column=0, pady=10)
            
            # Crear Treeview
            tree = ttk.Treeview(main_frame, columns=("Items"), show="tree headings")
            tree.heading("#0", text="Categoría")
            tree.heading("Items", text="Ítems de Evaluación")
            tree.column("#0", width=200)
            tree.column("Items", width=550)
            
            # Obtener categorías del sistema
            categorias = self.sistema.obtener_categorias_items()
            
            # Llenar el Treeview
            for categoria, items in categorias.items():
                categoria_id = tree.insert("", "end", text=categoria)
                for item in items:
                    tree.insert(categoria_id, "end", values=(item,))
            
            # Agregar scrollbar
            scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            
            # Grid
            tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
            
            # Configurar expansión
            main_frame.columnconfigure(0, weight=1)
            main_frame.rowconfigure(1, weight=1)
            ventana.columnconfigure(0, weight=1)
            ventana.rowconfigure(0, weight=1)
            
            # Botón cerrar
            ttk.Button(
                main_frame,
                text="Cerrar",
                command=ventana.destroy
            ).grid(row=2, column=0, pady=10)
            
        except Exception as e:
            self.log_message(f"Error mostrando categorías: {str(e)}", "ERROR")
            messagebox.showerror("Error", f"Error mostrando categorías: {str(e)}")