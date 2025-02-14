import tkinter as tk
import sys
import os

# Obtener la ruta absoluta del directorio actual
current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.join(current_dir, 'src')

# Agregar el directorio src al path
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)

from src.app_evaluacion import EvaluacionDocenteApp

def main():
    try:
        root = tk.Tk()
        app = EvaluacionDocenteApp(root)
        root.mainloop()
    except Exception as e:
        tk.messagebox.showerror("Error Fatal", f"Error iniciando la aplicación: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()