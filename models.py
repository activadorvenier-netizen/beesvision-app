# models.py
import sqlite3
from datetime import datetime
from typing import Optional, List, Dict

class ReviewDatabase:
    def __init__(self, db_path="reviews.db"):
        self.db_path = db_path
        self._init_db()
    
    def _init_db(self):
        """Crear tablas necesarias si no existen"""
        with sqlite3.connect(self.db_path) as conn:
            conn.executescript("""
                -- Tabla de revisiones
                CREATE TABLE IF NOT EXISTS reviews (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_id TEXT NOT NULL,
                    row_id INTEGER NOT NULL,
                    supervisor_id INTEGER NOT NULL,
                    supervisor_name TEXT NOT NULL,
                    status TEXT NOT NULL,
                    observaciones TEXT,
                    fecha_revision DATETIME DEFAULT CURRENT_TIMESTAMP,
                    excel_file TEXT,
                    mes_revision TEXT,
                    UNIQUE(task_id, supervisor_id)
                );
                
                -- Índices para búsquedas rápidas
                CREATE INDEX IF NOT EXISTS idx_task_id ON reviews(task_id);
                CREATE INDEX IF NOT EXISTS idx_supervisor ON reviews(supervisor_id);
                CREATE INDEX IF NOT EXISTS idx_status ON reviews(status);
                CREATE INDEX IF NOT EXISTS idx_mes ON reviews(mes_revision);
                
                -- Tabla para control de archivos cargados
                CREATE TABLE IF NOT EXISTS uploaded_files (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    filename TEXT NOT NULL,
                    upload_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                    row_count INTEGER,
                    supervisor_id INTEGER,
                    mes_revision TEXT
                );
            """)
    
    def save_review(self, task_id: str, row_id: int, supervisor_id: int, 
                    supervisor_name: str, status: str, observaciones: str = "",
                    excel_file: str = "", mes_revision: str = "") -> bool:
        """Guardar o actualizar una revisión"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # Verificar si ya existe
                cursor.execute("""
                    SELECT id FROM reviews 
                    WHERE task_id = ? AND supervisor_id = ?
                """, (task_id, supervisor_id))
                
                existing = cursor.fetchone()
                
                if existing:
                    # Actualizar
                    cursor.execute("""
                        UPDATE reviews 
                        SET status = ?, observaciones = ?, fecha_revision = CURRENT_TIMESTAMP
                        WHERE task_id = ? AND supervisor_id = ?
                    """, (status, observaciones, task_id, supervisor_id))
                else:
                    # Insertar nuevo
                    cursor.execute("""
                        INSERT INTO reviews 
                        (task_id, row_id, supervisor_id, supervisor_name, status, 
                         observaciones, excel_file, mes_revision)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (task_id, row_id, supervisor_id, supervisor_name, status, 
                          observaciones, excel_file, mes_revision))
                
                conn.commit()
                return True
        except Exception as e:
            print(f"Error al guardar revisión: {e}")
            return False
    
    def get_review_status(self, task_id: str, supervisor_id: int) -> Optional[Dict]:
        """Obtener el estado de una revisión"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT status, observaciones, fecha_revision
                    FROM reviews
                    WHERE task_id = ? AND supervisor_id = ?
                """, (task_id, supervisor_id))
                
                row = cursor.fetchone()
                return dict(row) if row else None
        except Exception as e:
            print(f"Error al obtener revisión: {e}")
            return None
    
    def get_supervisor_stats(self, supervisor_id: int, mes: str = None) -> Dict:
        """Estadísticas de un supervisor (opcionalmente por mes)"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                if mes:
                    cursor.execute("""
                        SELECT COUNT(*) FROM reviews 
                        WHERE supervisor_id = ? AND mes_revision = ?
                    """, (supervisor_id, mes))
                    total = cursor.fetchone()[0]
                    
                    cursor.execute("""
                        SELECT status, COUNT(*) 
                        FROM reviews 
                        WHERE supervisor_id = ? AND mes_revision = ?
                        GROUP BY status
                    """, (supervisor_id, mes))
                else:
                    cursor.execute("""
                        SELECT COUNT(*) FROM reviews WHERE supervisor_id = ?
                    """, (supervisor_id,))
                    total = cursor.fetchone()[0]
                    
                    cursor.execute("""
                        SELECT status, COUNT(*) 
                        FROM reviews 
                        WHERE supervisor_id = ?
                        GROUP BY status
                    """, (supervisor_id,))
                
                by_status = {row[0]: row[1] for row in cursor.fetchall()}
                
                return {
                    "total": total,
                    "by_status": by_status
                }
        except Exception as e:
            print(f"Error al obtener estadísticas: {e}")
            return {"total": 0, "by_status": {}}
    
    def get_all_reviews(self, supervisor_id: Optional[int] = None, mes: str = None) -> List[Dict]:
        """Obtener todas las revisiones (opcionalmente filtradas)"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                
                query = "SELECT * FROM reviews"
                params = []
                conditions = []
                
                if supervisor_id:
                    conditions.append("supervisor_id = ?")
                    params.append(supervisor_id)
                
                if mes:
                    conditions.append("mes_revision = ?")
                    params.append(mes)
                
                if conditions:
                    query += " WHERE " + " AND ".join(conditions)
                
                query += " ORDER BY fecha_revision DESC"
                
                cursor.execute(query, params)
                return [dict(row) for row in cursor.fetchall()]
        except Exception as e:
            print(f"Error al obtener revisiones: {e}")
            return []
    
    def get_pending_tasks(self, supervisor_id: int, task_ids: List[str]) -> List[str]:
        """Obtener tareas pendientes de revisión"""
        if not task_ids:
            return []
        
        try:
            placeholders = ','.join(['?'] * len(task_ids))
            
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                query = f"""
                    SELECT task_id FROM reviews 
                    WHERE task_id IN ({placeholders}) 
                    AND supervisor_id = ?
                """
                cursor.execute(query, task_ids + [supervisor_id])
                reviewed = {row[0] for row in cursor.fetchall()}
                
                return [tid for tid in task_ids if tid not in reviewed]
        except Exception as e:
            print(f"Error al obtener tareas pendientes: {e}")
            return task_ids
    
    def clear_reviews(self, supervisor_id: int, mes: str = None) -> bool:
        """Eliminar revisiones de un supervisor (opcionalmente por mes)"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                if mes:
                    cursor.execute("""
                        DELETE FROM reviews 
                        WHERE supervisor_id = ? AND mes_revision = ?
                    """, (supervisor_id, mes))
                else:
                    cursor.execute("""
                        DELETE FROM reviews WHERE supervisor_id = ?
                    """, (supervisor_id,))
                
                conn.commit()
                return True
        except Exception as e:
            print(f"Error al limpiar revisiones: {e}")
            return False
    
    def delete_review(self, task_id: str, supervisor_id: int) -> bool:
        """Eliminar una revisión específica"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    DELETE FROM reviews 
                    WHERE task_id = ? AND supervisor_id = ?
                """, (task_id, supervisor_id))
                conn.commit()
                return True
        except Exception as e:
            print(f"Error al eliminar revisión: {e}")
            return False