import pandas as pd
import pyodbc
from datetime import datetime
import os
import logging
from typing import Optional, List, Dict, Tuple
from tkinter import messagebox

class EvaluacionDocenteSystem:
    def __init__(self):
        # Configurar logging
        log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'logs')
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            filename=os.path.join(log_dir, 'evaluacion_docente.log')
        )
        self.logger = logging.getLogger(__name__)
        
        # Configuración de la base de datos
        self.conn_str = (
            "DRIVER={ODBC Driver 17 for SQL Server};"
            "SERVER=localhost\\SQLEXPRESS;"
            "DATABASE=EvaluacionDocenteDB;"
            "Trusted_Connection=yes;"
        )
        self.conn = None
        
        # Estados válidos para la evaluación
        self.estados_validos = [
            'Cumplimiento satisfactorio',
            'Cumplimiento parcial',
            'Incumplimiento',
            'No Aplica'
        ]

    def conectar_bd(self) -> bool:
        """Establece conexión con la base de datos SQL Server"""
        try:
            self.conn = pyodbc.connect(self.conn_str)
            self.logger.info("Conexión a base de datos establecida")
            return True
        except Exception as e:
            error_msg = f"Error de conexión a la base de datos: {str(e)}"
            self.logger.error(error_msg)
            messagebox.showerror("Error de Conexión", error_msg)
            raise ConnectionError(error_msg)

    def obtener_facultades(self) -> List[str]:
        """Obtiene la lista de facultades activas"""
        try:
            if not self.conn:
                self.conectar_bd()
            cursor = self.conn.cursor()
            cursor.execute("SELECT Nombre FROM Facultades WHERE Estado = 1")
            return [row[0] for row in cursor.fetchall()]
        except Exception as e:
            self.logger.error(f"Error obteniendo facultades: {str(e)}")
            return []

    def obtener_carreras_por_facultad(self, facultad: str) -> List[str]:
        """Obtiene las carreras activas de una facultad específica"""
        try:
            if not self.conn:
                self.conectar_bd()
            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT c.Nombre 
                FROM Carreras c
                JOIN Facultades f ON c.FacultadID = f.FacultadID
                WHERE f.Nombre = ? AND c.Estado = 1
            """, facultad)
            return [row[0] for row in cursor.fetchall()]
        except Exception as e:
            self.logger.error(f"Error obteniendo carreras: {str(e)}")
            return []

    def obtener_categorias_items(self) -> Dict[str, List[str]]:
        """Obtiene las categorías y sus items desde la base de datos"""
        try:
            if not self.conn:
                self.conectar_bd()
            
            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT 
                    c.Nombre as Categoria,
                    i.Nombre as Item
                FROM CategoriasEvaluacion c
                INNER JOIN ItemsEvaluacion i ON c.CategoriaID = i.CategoriaID
                WHERE c.Estado = 1 AND i.Estado = 1
                ORDER BY c.Orden, i.Orden
            """)
            
            categorias = {}
            for row in cursor.fetchall():
                categoria, item = row
                if categoria not in categorias:
                    categorias[categoria] = []
                categorias[categoria].append(item)
            
            return categorias
        except Exception as e:
            error_msg = f"Error obteniendo categorías: {str(e)}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
            return {}

    def validar_datos_excel(self, df_general: pd.DataFrame, df_eval: pd.DataFrame) -> Tuple[bool, str]:
        """Valida los datos del Excel antes de procesarlos"""
        try:
            # Limpiar nombres de columnas
            df_eval.columns = [col.strip() for col in df_eval.columns]
            
            # Verificar datos generales
            campos_requeridos = {
                2: "Periodo Académico",
                3: "Facultad",
                4: "Carrera",
                5: "Revisado Por",
                6: "Asignatura",
                7: "Nombre del Docente"
            }
            
            # Validar campos requeridos
            for idx, campo in campos_requeridos.items():
                if pd.isnull(df_general.iloc[idx, 1]):
                    return False, f"El campo {campo} es requerido"
                
            # Validar facultad y carrera
            facultad = str(df_general.iloc[3, 1]).strip()
            carrera = str(df_general.iloc[4, 1]).strip()
            
            facultades_validas = self.obtener_facultades()
            if facultad not in facultades_validas:
                return False, f"La facultad '{facultad}' no existe en el sistema"
            
            carreras_validas = self.obtener_carreras_por_facultad(facultad)
            if carrera not in carreras_validas:
                return False, f"La carrera '{carrera}' no existe para la facultad '{facultad}'"
            
            # Validar estructura de evaluación
            columnas_requeridas = ['CATEGORÍA', 'ÍTEM DE EVALUACIÓN', 'ESTADO', 'FECHA']
            for columna in columnas_requeridas:
                if columna not in df_eval.columns:
                    return False, f"La columna '{columna}' es requerida en la hoja de evaluación"
            
            # Validar categorías e items
            categorias_items = self.obtener_categorias_items()
            for idx, row in df_eval.iterrows():
                if pd.notna(row['CATEGORÍA']) and pd.notna(row['ÍTEM DE EVALUACIÓN']):
                    categoria = str(row['CATEGORÍA']).strip()
                    item = str(row['ÍTEM DE EVALUACIÓN']).strip()
                    
                    if categoria not in categorias_items:
                        return False, f"Categoría inválida en fila {idx + 2}: {categoria}"
                    
                    if item not in categorias_items[categoria]:
                        return False, f"Ítem inválido para la categoría '{categoria}' en fila {idx + 2}: {item}"
            
            # Validar estados
            estados_invalidos = []
            for idx, row in df_eval.iterrows():
                if pd.notna(row['ESTADO']):
                    estado = str(row['ESTADO']).strip()
                    if estado not in self.estados_validos:
                        estados_invalidos.append(f"Fila {idx + 2}: {estado}")
            
            if estados_invalidos:
                return False, f"Estados inválidos encontrados:\n" + "\n".join(estados_invalidos)
            
            return True, ""
            
        except Exception as e:
            error_msg = f"Error en validación: {str(e)}"
            self.logger.error(error_msg)
            return False, error_msg

    def procesar_archivo_excel(self, ruta_archivo: str) -> bool:
        """Procesa un archivo Excel de evaluación docente"""
        try:
            self.logger.info(f"Procesando archivo: {os.path.basename(ruta_archivo)}")
            
            # Verificar conexión
            if not self.conn:
                self.conectar_bd()
            
            # Leer datos del Excel
            df_general = pd.read_excel(
                ruta_archivo,
                sheet_name='DATOS_GENERALES',
                header=None,
                engine='openpyxl'
            )
            
            df_eval = pd.read_excel(
                ruta_archivo,
                sheet_name='EVALUACION',
                engine='openpyxl'
            )
            
            # Validación de datos
            es_valido, mensaje_error = self.validar_datos_excel(df_general, df_eval)
            if not es_valido:
                raise ValueError(mensaje_error)
            
            cursor = self.conn.cursor()
            
            # Extraer datos generales
            periodo_academico = str(df_general.iloc[2, 1]).strip()
            facultad = str(df_general.iloc[3, 1]).strip()
            carrera = str(df_general.iloc[4, 1]).strip()
            revisado_por = str(df_general.iloc[5, 1]).strip()
            asignatura = str(df_general.iloc[6, 1]).strip()
            nombre_docente = str(df_general.iloc[7, 1]).strip()
            
            # Registrar evaluación
            cursor.execute("""
                DECLARE @EvaluacionID INT;
                EXEC sp_RegistrarEvaluacion 
                    @PeriodoAcademico = ?, 
                    @NombreDocente = ?,
                    @Asignatura = ?,
                    @Carrera = ?,
                    @Facultad = ?,
                    @RevisadoPor = ?,
                    @FechaEvaluacion = ?,
                    @EvaluacionID = @EvaluacionID OUTPUT;
                SELECT @EvaluacionID;
            """, (
                periodo_academico, nombre_docente, asignatura,
                carrera, facultad, revisado_por,
                df_eval['FECHA'].max()
            ))
            
            evaluacion_id = cursor.fetchval()
            if not evaluacion_id:
                raise ValueError("No se pudo obtener el ID de la evaluación")
            
            # Procesar resultados de evaluación
            for _, row in df_eval.iterrows():
                if pd.notna(row['ÍTEM DE EVALUACIÓN']):
                    cursor.execute("""
                        EXEC sp_RegistrarResultadosEvaluacion
                            @EvaluacionID = ?,
                            @ItemNombre = ?,
                            @Estado = ?,
                            @FechaRevision = ?,
                            @Observaciones = ?
                    """, (
                        evaluacion_id,
                        str(row['ÍTEM DE EVALUACIÓN']).strip(),
                        str(row['ESTADO']).strip() if pd.notna(row['ESTADO']) else 'No Aplica',
                        row['FECHA'] if pd.notna(row['FECHA']) else datetime.now().date(),
                        str(row['OBSERVACIONES']).strip() if pd.notna(row['OBSERVACIONES']) else None
                    ))
            
            # Calcular porcentaje de cumplimiento
            cursor.execute("""
                EXEC sp_CalcularPorcentajeCumplimiento @EvaluacionID = ?
            """, evaluacion_id)
            
            self.conn.commit()
            self.logger.info(f"Archivo procesado correctamente. EvaluacionID: {evaluacion_id}")
            return True
            
        except Exception as e:
            error_msg = f"Error procesando archivo: {str(e)}"
            self.logger.error(error_msg)
            if self.conn:
                self.conn.rollback()
            messagebox.showerror("Error", error_msg)
            raise
        finally:
            if self.conn:
                self.conn.close()
                self.conn = None