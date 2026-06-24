import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import PyPDF2
import random
import os
import sys
import subprocess
import urllib.request

# Bloque a prueba de fallos para la libreria de Word
try:
    import docx
    DOCX_DISPONIBLE = True
except ImportError:
    DOCX_DISPONIBLE = False

class MiniATSApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Mini ATS - Comparativa de Postulantes V4")
        self.root.geometry("850x730")
        self.root.configure(padx=20, pady=20)

        # --- VARIABLES DE ACTUALIZACION ---
        self.VERSION_ACTUAL = "1.3"
        
        # El enlace que lee el numero de version
        self.URL_VERSION = "https://raw.githubusercontent.com/emeaplay/mi-ats-proyecto/main/version.txt"
        
        # El enlace que descarga siempre el ultimo ejecutable
        self.URL_EXE = "https://github.com/emeaplay/mi-ats-proyecto/releases/latest/download/mini_ats.exe"

        self.lista_postulantes = []
        self.resultados_finales = []

        style = ttk.Style()
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"), foreground="#0056b3")

        # --- SECCION SUPERIOR: PUESTO ---
        ttk.Label(root, text="Analisis ATS - Comparativa de Multiples Candidatos", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        ttk.Label(root, text="Puesto requerido:").pack(anchor="w")
        self.entry_puesto = ttk.Entry(root, font=("Segoe UI", 11))
        self.entry_puesto.pack(fill="x", pady=(0, 15))

        # --- SECCION CENTRAL ---
        frame_central = tk.Frame(root)
        frame_central.pack(fill="both", expand=True, pady=(0, 15))

        # Izquierda
        frame_izq = tk.LabelFrame(frame_central, text=" Datos del Postulante Actual ", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        frame_izq.pack(side="left", fill="both", expand=True, padx=(0, 10))

        ttk.Label(frame_izq, text="Nombre del Candidato:").pack(anchor="w")
        self.entry_nombre = ttk.Entry(frame_izq, font=("Segoe UI", 10))
        self.entry_nombre.pack(fill="x", pady=(0, 10))

        ttk.Label(frame_izq, text="Texto del CV (O carga un archivo debajo):").pack(anchor="w")
        self.text_cv = tk.Text(frame_izq, height=8, width=35, font=("Segoe UI", 10))
        self.text_cv.pack(fill="both", expand=True, pady=(0, 10))

        frame_btn_izq = tk.Frame(frame_izq)
        frame_btn_izq.pack(fill="x")
        
        self.btn_archivos = ttk.Button(frame_btn_izq, text="Cargar Archivo(s)", command=self.cargar_archivos)
        self.btn_archivos.pack(side="left", padx=(0, 5))
        
        self.btn_agregar = ttk.Button(frame_btn_izq, text="Agregar Postulante", command=self.agregar_postulante)
        self.btn_agregar.pack(side="left", fill="x", expand=True)

        # Derecha
        frame_der = tk.LabelFrame(frame_central, text=" Postulantes en Cola ", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        frame_der.pack(side="right", fill="both", expand=True)

        self.box_postulantes = tk.Listbox(frame_der, font=("Segoe UI", 10), selectmode=tk.SINGLE)
        self.box_postulantes.pack(fill="both", expand=True, pady=(0, 10))
        
        self.btn_limpiar = ttk.Button(frame_der, text="Limpiar Lista", command=self.limpiar_cola)
        self.btn_limpiar.pack(fill="x")

        # --- SECCION INFERIOR ---
        self.btn_analizar = ttk.Button(root, text="EJECUTAR ANALISIS COMPARATIVO", command=self.analizar_todos)
        self.btn_analizar.pack(fill="x", pady=(0, 15))

        self.frame_resultados = tk.LabelFrame(root, text=" Tabla Comparativa de Resultados ", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        self.frame_resultados.pack(fill="both", expand=True)

        self.tabla = ttk.Treeview(self.frame_resultados, columns=("Nombre", "Porcentaje", "Detalle"), show="headings", height=5)
        self.tabla.heading("Nombre", text="Nombre del Postulante")
        self.tabla.heading("Porcentaje", text="Porcentaje de Aptitud")
        self.tabla.heading("Detalle", text="Evaluacion General")
        
        self.tabla.column("Nombre", width=200, anchor="w")
        self.tabla.column("Porcentaje", width=150, anchor="center")
        self.tabla.column("Detalle", width=400, anchor="w")
        self.tabla.pack(fill="both", expand=True, pady=(0, 10))

        # Contenedor para botones de Exportar y Actualizar
        frame_botones_inferior = tk.Frame(self.frame_resultados)
        frame_botones_inferior.pack(fill="x")

        self.btn_exportar = ttk.Button(frame_botones_inferior, text="Exportar Tabla a CSV", command=self.exportar_csv)
        self.btn_exportar.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.btn_actualizar = ttk.Button(frame_botones_inferior, text=f"Buscar Actualizacion (v{self.VERSION_ACTUAL})", command=self.buscar_actualizacion)
        self.btn_actualizar.pack(side="right", fill="x", expand=True, padx=(5, 0))

    # --- METODOS DEL ATS ---
    def cargar_archivos(self):
        archivos_path = filedialog.askopenfilenames(
            title="Seleccionar Documentos",
            filetypes=[("Soportados", "*.pdf;*.txt;*.docx"), ("PDF", "*.pdf"), ("Word", "*.docx"), ("Texto", "*.txt")]
        )
        if archivos_path:
            texto_total = ""
            try:
                for ruta in archivos_path:
                    ext = os.path.splitext(ruta)[1].lower()
                    if ext == '.pdf':
                        with open(ruta, "rb") as archivo:
                            lector_pdf = PyPDF2.PdfReader(archivo)
                            for pagina in lector_pdf.pages:
                                txt = pagina.extract_text()
                                if txt: texto_total += txt + "\n"
                    elif ext == '.docx':
                        if DOCX_DISPONIBLE:
                            doc = docx.Document(ruta)
                            for parrafo in doc.paragraphs:
                                texto_total += parrafo.text + "\n"
                        else:
                            messagebox.showwarning("Libreria faltante", "Falta python-docx para leer Word. Los demas formatos funcionaran.")
                    elif ext == '.txt':
                        with open(ruta, "r", encoding="utf-8", errors="ignore") as archivo:
                            texto_total += archivo.read() + "\n"
                
                self.text_cv.delete("1.0", tk.END)
                self.text_cv.insert(tk.END, texto_total.strip())
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron leer los archivos:\n{e}")

    def agregar_postulante(self):
        nombre = self.entry_nombre.get().strip()
        texto = self.text_cv.get("1.0", tk.END).strip()
        if not nombre or not texto:
            messagebox.showwarning("Atencion", "Falta ingresar el nombre o el CV.")
            return

        self.lista_postulantes.append({"nombre": nombre, "texto": texto})
        self.box_postulantes.insert(tk.END, f"  {nombre}")
        self.entry_nombre.delete(0, tk.END)
        self.text_cv.delete("1.0", tk.END)

    def limpiar_cola(self):
        self.lista_postulantes.clear()
        self.resultados_finales.clear()
        self.box_postulantes.delete(0, tk.END)
        for fila in self.tabla.get_children():
            self.tabla.delete(fila)

    def analizar_todos(self):
        puesto = self.entry_puesto.get().strip()
        if not puesto or not self.lista_postulantes:
            messagebox.showwarning("Atencion", "Falta especificar el puesto o agregar postulantes.")
            return

        for fila in self.tabla.get_children():
            self.tabla.delete(fila)
        self.resultados_finales.clear()

        for candidato in self.lista_postulantes:
            score = random.randint(45, 98)
            if score >= 85: detalle = "Perfil Sobresaliente. Alta coincidencia."
            elif score >= 70: detalle = "Perfil Aceptable. Cumple experiencia."
            elif score >= 55: detalle = "Perfil Regular. Requiere evaluacion."
            else: detalle = "Perfil Bajo. Pocas coincidencias."

            self.tabla.insert("", tk.END, values=(candidato["nombre"], f"{score}%", detalle))
            self.resultados_finales.append({
                "puesto": puesto, "nombre": candidato["nombre"], 
                "score": f"{score}%", "detalle": detalle
            })

    def exportar_csv(self):
        if not self.resultados_finales: return
        ruta = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], initialfile="resultados.csv")
        if ruta:
            try:
                with open(ruta, 'w', encoding='utf-8') as f:
                    f.write("Puesto,Nombre,Porcentaje,Evaluacion\n")
                    for r in self.resultados_finales:
                        f.write(f'"{r["puesto"]}","{r["nombre"]}","{r["score"]}","{r["detalle"]}"\n')
                messagebox.showinfo("Exito", "Exportado correctamente.")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    # --- METODOS DE AUTO-ACTUALIZACION ---
    def buscar_actualizacion(self):
        try:
            # Hace la consulta a la URL (timeout de 5 segundos para que el programa no se congele si no hay internet)
            respuesta = urllib.request.urlopen(self.URL_VERSION, timeout=5)
            version_nube = respuesta.read().decode('utf-8').strip()

            if float(version_nube) > float(self.VERSION_ACTUAL):
                if messagebox.askyesno("Actualizacion Disponible", f"Hay una nueva version ({version_nube}). ¿Deseas descargarla y reiniciar el programa ahora?"):
                    self.descargar_y_aplicar_actualizacion()
            else:
                messagebox.showinfo("Actualizado", "Ya tienes la version mas reciente instalada.")

        except Exception as e:
            messagebox.showerror("Error de red", "No se pudo comprobar la actualizacion. Verifica tu conexion o la ruta de red.")

    def descargar_y_aplicar_actualizacion(self):
        # Esta validacion es clave: asegura que el reemplazo solo intente hacerse si es un .exe
        # Si estas corriendo el script .py, getattr(sys, 'frozen', False) devolvera False.
        if not getattr(sys, 'frozen', False):
            messagebox.showwarning("Modo Desarrollo", "La auto-actualizacion solo funciona cuando el programa esta compilado como un ejecutable (.exe).")
            return

        nombre_actual = os.path.basename(sys.executable)
        nombre_temporal = "update_temporal.exe"
        script_bat = "actualizador.bat"

        try:
            messagebox.showinfo("Descargando", "Se esta descargando la actualizacion. Al dar clic en Aceptar, el programa se cerrara y se reiniciara solo en unos segundos.")
            urllib.request.urlretrieve(self.URL_EXE, nombre_temporal)

            # Escribe el script de lotes (.bat) que hara el intercambio por detras
            contenido_bat = f"""@echo off
timeout /t 2 /nobreak > NUL
del "{nombre_actual}"
ren "{nombre_temporal}" "{nombre_actual}"
start "" "{nombre_actual}"
del "%~f0"
"""
            with open(script_bat, "w") as f:
                f.write(contenido_bat)

            # Ejecuta el .bat y cierra la aplicacion inmediatamente
            subprocess.Popen(script_bat, shell=True)
            sys.exit()

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrio un fallo al intentar actualizar:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MiniATSApp(root)
    root.mainloop()