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
        
        # --- 1. PRIMERO DEFINIMOS LAS VARIABLES CLAVE ---
        self.VERSION_ACTUAL = "1.4.5"
        self.URL_VERSION = "https://raw.githubusercontent.com/emeaplay/mi-ats-proyecto/main/version.txt"
        self.URL_EXE = "https://github.com/emeaplay/mi-ats-proyecto/releases/latest/download/mini_ats.exe"

        # --- 2. LUEGO LAS USAMOS EN LA INTERFAZ ---
        # Ahora Python ya sabe qué es VERSION_ACTUAL y lo pondrá en el título
        self.root.title(f"Mini ATS - MRB Seleccion (EMEsoft) v{self.VERSION_ACTUAL}")
        
        # Altura ajustada a 700 para evitar que se corte en laptops
        self.root.geometry("900x700") 
        
        # --- PALETA DE COLORES REFINADA (BOOTSTRAP 5) ---
        self.COLOR_BG = "#f8f9fa"          
        self.COLOR_PRIMARY = "#0d6efd"     
        self.COLOR_SUCCESS = "#198754"     
        self.COLOR_DANGER = "#dc3545"      
        self.COLOR_SECONDARY = "#6c757d"   
        self.COLOR_TEXT = "#212529"        
        
        self.root.configure(bg=self.COLOR_BG, padx=20, pady=15)

        self.lista_postulantes = []
        self.resultados_finales = []


        # --- CONFIGURACION DE ESTILOS TTK ---
        style = ttk.Style()
        style.theme_use("clam")  
        
        style.configure("TLabel", font=("Segoe UI", 10), background=self.COLOR_BG, foreground=self.COLOR_TEXT)
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), background=self.COLOR_BG, foreground=self.COLOR_PRIMARY)
        style.configure("Sub.TLabel", font=("Segoe UI", 9), background=self.COLOR_BG, foreground=self.COLOR_SECONDARY)
        
        style.configure("TEntry", fieldbackground="#ffffff", bordercolor="#dee2e6", lightcolor="#dee2e6", darkcolor="#dee2e6")
        
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=28, background="#ffffff", fieldbackground="#ffffff", foreground=self.COLOR_TEXT)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#e9ecef", foreground=self.COLOR_TEXT, relief="flat")
        style.map("Treeview", background=[('selected', '#e7f1ff')], foreground=[('selected', self.COLOR_PRIMARY)])

        # --- SECCION SUPERIOR: ENCABEZADO Y BOTON ACTUALIZAR ---
        frame_header = tk.Frame(root, bg=self.COLOR_BG)
        frame_header.pack(fill="x", pady=(0, 15))

        # Contenedor para los titulos (Izquierda)
        frame_titulos = tk.Frame(frame_header, bg=self.COLOR_BG)
        frame_titulos.pack(side="left")
        ttk.Label(frame_titulos, text="Analisis ATS - Comparativa de Multiples Candidatos", style="Header.TLabel").pack(anchor="w")
        ttk.Label(frame_titulos, text="MRB Seleccion Peru - Gestion de Talento Humano", style="Sub.TLabel").pack(anchor="w")

        # Boton pequeño de actualizacion (Derecha)
        self.btn_actualizar = tk.Button(
            frame_header, text=f"↻ Buscar Actualizacion (v{self.VERSION_ACTUAL})", 
            command=self.buscar_actualizacion, font=("Segoe UI", 8, "bold"), 
            bg=self.COLOR_PRIMARY, fg="#ffffff", relief="flat", bd=0, 
            activebackground=self.COLOR_PRIMARY, activeforeground="#ffffff", cursor="hand2", padx=10, pady=4
        )
        self.btn_actualizar.pack(side="right", anchor="n")

        # --- PUESTO REQUERIDO ---
        ttk.Label(root, text="Puesto requerido / Vacante analizada:").pack(anchor="w")
        self.entry_puesto = ttk.Entry(root, font=("Segoe UI", 11))
        self.entry_puesto.pack(fill="x", pady=(0, 15))

        # --- SECCION CENTRAL: INTERFAZ DE DOS COLUMNAS ---
        frame_central = tk.Frame(root, bg=self.COLOR_BG)
        frame_central.pack(fill="both", expand=True, pady=(0, 15))

        # Columna Izquierda: Captura de Datos
        frame_izq = tk.LabelFrame(frame_central, text=" Datos del Postulante Actual ", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg=self.COLOR_PRIMARY, padx=15, pady=10, relief="solid", bd=1)
        frame_izq.pack(side="left", fill="both", expand=True, padx=(0, 15))

        ttk.Label(frame_izq, text="Nombre del Candidato:", background="#ffffff").pack(anchor="w")
        self.entry_nombre = ttk.Entry(frame_izq, font=("Segoe UI", 10))
        self.entry_nombre.pack(fill="x", pady=(0, 8))

        ttk.Label(frame_izq, text="Texto del CV (O carga un archivo debajo):", background="#ffffff").pack(anchor="w")
        self.text_cv = tk.Text(frame_izq, height=6, width=35, font=("Segoe UI", 10), bg="#ffffff", fg=self.COLOR_TEXT, bd=1, relief="solid", highlightthickness=0, insertbackground=self.COLOR_TEXT)
        self.text_cv.pack(fill="both", expand=True, pady=(0, 10))

        frame_btn_izq = tk.Frame(frame_izq, bg="#ffffff")
        frame_btn_izq.pack(fill="x")
        self.btn_archivos = self.crear_boton_estilo(frame_btn_izq, "Cargar Archivo(s)", self.COLOR_SECONDARY, self.cargar_archivos)
        self.btn_archivos.pack(side="left", padx=(0, 5))
        self.btn_agregar = self.crear_boton_estilo(frame_btn_izq, "Agregar a la Cola", self.COLOR_PRIMARY, self.agregar_postulante)
        self.btn_agregar.pack(side="left", fill="x", expand=True)

        # Columna Derecha: Lista en espera
        frame_der = tk.LabelFrame(frame_central, text=" Postulantes en Cola ", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg=self.COLOR_PRIMARY, padx=15, pady=10, relief="solid", bd=1)
        frame_der.pack(side="right", fill="both", expand=True)

        self.box_postulantes = tk.Listbox(frame_der, font=("Segoe UI", 10), selectmode=tk.SINGLE, bg="#ffffff", fg=self.COLOR_TEXT, bd=1, relief="solid", highlightthickness=0)
        self.box_postulantes.pack(fill="both", expand=True, pady=(0, 10))
        
        self.btn_limpiar = self.crear_boton_estilo(frame_der, "Limpiar Lista", self.COLOR_DANGER, self.limpiar_cola)
        self.btn_limpiar.pack(fill="x")

        # --- SECCION INFERIOR: ACCION PRINCIPAL Y RESULTADOS ---
        self.btn_analizar = self.crear_boton_estilo(root, "⚡ COMPARAR", self.COLOR_SUCCESS, self.analizar_todos)
        self.btn_analizar.pack(fill="x", pady=(0, 15))

        self.frame_resultados = tk.LabelFrame(root, text=" Tabla Comparativa ", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg=self.COLOR_PRIMARY, padx=15, pady=10, relief="solid", bd=1)
        self.frame_resultados.pack(fill="both", expand=True)

        # Altura de la tabla reducida a 4 para garantizar que encaje en la pantalla
        self.tabla = ttk.Treeview(self.frame_resultados, columns=("Nombre", "Porcentaje", "Detalle"), show="headings", height=4)
        self.tabla.heading("Nombre", text="Nombre del Postulante")
        self.tabla.heading("Porcentaje", text="Porcentaje de Aptitud")
        self.tabla.heading("Detalle", text="Evaluacion General")
        
        self.tabla.column("Nombre", width=220, anchor="w")
        self.tabla.column("Porcentaje", width=160, anchor="center")
        self.tabla.column("Detalle", width=420, anchor="w")
        self.tabla.pack(fill="both", expand=True, pady=(0, 10))

        # Boton Exportar ahora abarca todo el ancho debajo de la tabla
        self.btn_exportar = self.crear_boton_estilo(self.frame_resultados, "⬇ Exportar Tabla a CSV", self.COLOR_SECONDARY, self.exportar_csv)
        self.btn_exportar.pack(fill="x")

    # --- HELPERS VISUALES: BOTONES FLAT (PLANO) ---
    def crear_boton_estilo(self, contenedor, texto, color_fondo, comando):
        btn = tk.Button(
            contenedor, text=texto, command=comando,
            font=("Segoe UI", 10, "bold"), bg=color_fondo, fg="#ffffff",
            relief="flat", bd=0, highlightthickness=0, 
            activebackground=color_fondo, activeforeground="#ffffff",
            cursor="hand2", padx=15, pady=7
        )
        return btn

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
                            messagebox.showwarning("Libreria faltante", "Falta python-docx para leer Word.")
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
            respuesta = urllib.request.urlopen(self.URL_VERSION, timeout=5)
            version_nube = respuesta.read().decode('utf-8').strip()

            # MAGIA AQUI: Convierte "1.4.2" en listas matemáticas [1, 4, 2] para comparar sin errores
            v_nube = [int(n) for n in version_nube.split(".")]
            v_actual = [int(n) for n in self.VERSION_ACTUAL.split(".")]

            if v_nube > v_actual:
                if messagebox.askyesno("Actualizacion Disponible", f"Hay una nueva version ({version_nube}). ¿Deseas descargarla y reiniciar el programa ahora?"):
                    self.descargar_y_aplicar_actualizacion()
            else:
                messagebox.showinfo("Actualizado", "Ya tienes la version mas reciente instalada.")
        
        except ValueError:
            messagebox.showerror("Error Interno", "Fallo al leer el formato de la versión (Revisa que solo tenga números y puntos).")
        except Exception as e:
            # Ahora si te mostrara el error real si es que de verdad falla el internet
            messagebox.showerror("Error de red", f"No se pudo comprobar la actualizacion.\nDetalle tecnico: {e}")

    def descargar_y_aplicar_actualizacion(self):
        if not getattr(sys, 'frozen', False):
            messagebox.showwarning("Modo Desarrollo", "La auto-actualizacion solo funciona en el ejecutable (.exe).")
            return

        nombre_actual = os.path.basename(sys.executable)
        nombre_temporal = "update_temporal.exe"
        script_bat = "actualizador.bat"

        try:
            messagebox.showinfo("Descargando", "Descargando actualizacion... El programa se reiniciara automaticamente.")
            urllib.request.urlretrieve(self.URL_EXE, nombre_temporal)

            contenido_bat = f"""@echo off
timeout /t 2 /nobreak > NUL
del "{nombre_actual}"
ren "{nombre_temporal}" "{nombre_actual}"
start "" "{nombre_actual}"
del "%~f0"
"""
            with open(script_bat, "w") as f:
                f.write(contenido_bat)

            subprocess.Popen(script_bat, shell=True)
            sys.exit()
        except Exception as e:
            import traceback
            detalle = traceback.format_exc()
            messagebox.showerror("Error de Diagnostico", f"Fallo interno detectado:\n{e}\n\nDetalle:\n{detalle}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MiniATSApp(root)
    root.mainloop()