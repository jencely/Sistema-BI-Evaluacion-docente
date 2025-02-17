import os
import shutil
import PyInstaller.__main__

def build_executable():
    # Limpiar directorios anteriores
    if os.path.exists("dist"):
        shutil.rmtree("dist")
    if os.path.exists("build"):
        shutil.rmtree("build")

    # Configurar PyInstaller
    PyInstaller.__main__.run([
        'main.py',
        '--name=EvaluacionDocente',
        '--windowed',
        '--onedir',
        '--add-data=resources/templates;resources/templates',
        '--add-data=docs;docs',
        '--hidden-import=pandas',
        '--hidden-import=pyodbc',
        '--hidden-import=openpyxl',
        '--hidden-import=evaluacion_docente',  # Agregado
        '--path=src',  # Agregado - incluye la carpeta src en el path
        '--add-data=src;src'  # Agregado - incluye los archivos de src
    ])

    # Copiar archivos adicionales
    if os.path.exists('resources/templates'):
        os.makedirs('dist/EvaluacionDocente/resources/templates', exist_ok=True)
        shutil.copytree(
            'resources/templates',
            'dist/EvaluacionDocente/resources/templates',
            dirs_exist_ok=True
        )
    
    if os.path.exists('docs'):
        os.makedirs('dist/EvaluacionDocente/docs', exist_ok=True)
        shutil.copytree(
            'docs',
            'dist/EvaluacionDocente/docs',
            dirs_exist_ok=True
        )

if __name__ == "__main__":
    build_executable()