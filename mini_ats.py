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
        self.root.title("Mini ATS - Comparativa de Postulantes V4 (Bootstrap Style)")
        self.root.geometry("900x780")
        
        # --- PALETA DE COLORES REFINADA (BOOTSTRAP 5) ---
        self.COLOR_BG = "#f8f9fa"          # Fondo general gris muy claro amigable
        self.COLOR_PRIMARY = "#0d6efd"     # Azul primario Bootstrap
        self.COLOR_SUCCESS = "#198754"     # Verde Success Bootstrap
        self.COLOR_DANGER = "#dc3545"      # Rojo Danger Bootstrap
        self.COLOR_SECONDARY = "#6c757d"   # Gris secundario
        self.COLOR_TEXT = "#212529"        # Texto casi negro moderno
        
        self.root.configure(bg=self.COLOR_BG, padx=25, pady=25)

        # --- VARIABLES DE ACTUALIZACION ---
        self.VERSION_ACTUAL = "1.42"
        self.URL_VERSION = "https://raw.githubusercontent.com/emeaplay/mi-ats-proyecto/main/version.txt"
        self.URL_EXE = "https://github.com/emeaplay/mi-ats-proyecto/releases/latest/download/mini_ats.exe"

        self.lista_postulantes = []
        self.resultados_finales = []

        # --- CONFIGURACION DE ESTILOS TTK (TABLA Y CAMPOS) ---
        style = ttk.Style()
        style.theme_use("clam")  # Clam permite personalizar bordes y fondos limpiamente
        
        style.configure("TLabel", font=("Segoe UI", 10), background=self.COLOR_BG, foreground=self.COLOR_TEXT)
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), background=self.COLOR_BG, foreground=self.COLOR_PRIMARY)
        style.configure("Sub.TLabel", font=("Segoe UI", 10, "bold"), background=self.COLOR_BG, foreground=self.COLOR_SECONDARY)
        
        # Estilo para cajas de texto de entrada (Inputs)
        style.configure("TEntry", fieldbackground="#ffffff", bordercolor="#dee2e6", lightcolor="#dee2e6", darkcolor="#dee2e6")
        
        # Estilo moderno para la Tabla de Resultados (Treeview)
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=28, background="#ffffff", fieldbackground="#ffffff", foreground=self.COLOR_TEXT)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#e9ecef", foreground=self.COLOR_TEXT, relief="flat")
        style.map("Treeview", background=[('selected', '#e7f1ff')], foreground=[('selected', self.COLOR_PRIMARY)])

        # --- SIMULADOR DE BOTONES REDONDEADOS BOOTSTRAP ---
        # Creamos imagenes en memoria de 2x2 pixeles transparentes con bordes y esquinas suaves para simular el radio de Bootstrap
        self.img_btn_primary = self.crear_forma_boton(self.COLOR_PRIMARY)
        self.img_btn_success = self.crear_forma_boton(self.COLOR_SUCCESS)
        self.img_btn_danger = self.crear_forma_boton(self.COLOR_DANGER)
        self.img_btn_secondary = self.crear_forma_boton(self.COLOR_SECONDARY)

        # --- SECCION SUPERIOR: ENCABEZADO Y PUESTO ---
        ttk.Label(root, text="Analisis ATS - Comparativa de Multiples Candidatos", style="Header.TLabel").pack(anchor="w", pady=(0, 2))
        ttk.Label(root, text="Panel de gestion y control de perfiles para procesos de seleccion", style="Sub.TLabel").pack(anchor="w", pady=(0, 15))
        
        ttk.Label(root, text="Puesto requerido / Vacante analizada:").pack(anchor="w")
        self.entry_puesto = ttk.Entry(root, font=("Segoe UI", 11))
        self.entry_puesto.pack(fill="x", pady=(0, 15))

        # --- SECCION CENTRAL: INTERFAZ DE DOS COLUMNAS ---
        frame_central = tk.Frame(root, bg=self.COLOR_BG)
        frame_central.pack(fill="both", expand=True, pady=(0, 15))

        # Columna Izquierda: Captura de Datos
        frame_izq = tk.LabelFrame(frame_central, text=" Datos del Postulante Actual ", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg=self.COLOR_PRIMARY, padx=15, pady=15, relief="solid", bd=1)
        frame_izq.pack(side="left", fill="both", expand=True, padx=(0, 15))

        ttk.Label(frame_izq, text="Nombre del Candidato:", background="#ffffff").pack(anchor="w")
        self.entry_nombre = ttk.Entry(frame_izq, font=("Segoe UI", 10))
        self.entry_nombre.pack(fill="x", pady=(0, 12))

        ttk.Label(frame_izq, text="Texto del CV (O carga un archivo debajo):", background="#ffffff").pack(anchor="w")
        self.text_cv = tk.Text(frame_izq, height=8, width=35, font=("Segoe UI", 10), bg="#ffffff", fg=self.COLOR_TEXT, bd=1, relief="solid", highlightthickness=0, insertbackground=self.COLOR_TEXT)
        self.text_cv.pack(fill="both", expand=True, pady=(0, 12))

        # Contenedor de Botones Izquierda
        frame_btn_izq = tk.Frame(frame_izq, bg="#ffffff")
        frame_btn_izq.pack(fill="x")
        
        self.btn_archivos = self.crear_boton_redondeado(frame_btn_izq, "Cargar Archivo(s)", self.img_btn_secondary, self.cargar_archivos)
        self.btn_archivos.pack(side="left", padx=(0, 5))
        
        self.btn_agregar = self.crear_boton_redondeado(frame_btn_izq, "Agregar Postulante a la Cola", self.img_btn_primary, self.agregar_postulante)
        self.btn_agregar.pack(side="left", fill="x", expand=True)

        # Columna Derecha: Lista en espera
        frame_der = tk.LabelFrame(frame_central, text=" Postulantes en Cola ", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg=self.COLOR_PRIMARY, padx=15, pady=15, relief="solid", bd=1)
        frame_der.pack(side="right", fill="both", expand=True)

        self.box_postulantes = tk.Listbox(frame_der, font=("Segoe UI", 10), selectmode=tk.SINGLE, bg="#ffffff", fg=self.COLOR_TEXT, bd=1, relief="solid", highlightthickness=0)
        self.box_postulantes.pack(fill="both", expand=True, pady=(0, 12))
        
        self.btn_limpiar = self.crear_boton_redondeado(frame_der, "Limpiar Lista de Espera", self.img_btn_danger, self.limpiar_cola)
        self.btn_limpiar.pack(fill="x")

        # --- SECCION INFERIOR: ACCION PRINCIPAL Y RESULTADOS ---
        self.btn_analizar = self.crear_boton_redondeado(root, "⚡ EJECUTAR ANALISIS COMPARATIVO ATS", self.img_btn_success, self.analizar_todos)
        self.btn_analizar.pack(fill="x", pady=(0, 15))

        self.frame_resultados = tk.LabelFrame(root, text=" Tabla Comparativa de Resultados ", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg=self.COLOR_PRIMARY, padx=15, pady=15, relief="solid", bd=1)
        self.frame_resultados.pack(fill="both", expand=True)

        # Construccion e integracion de la Tabla
        self.tabla = ttk.Treeview(self.frame_resultados, columns=("Nombre", "Porcentaje", "Detalle"), show="headings", height=5)
        self.tabla.heading("Nombre", text="Nombre del Postulante")
        self.tabla.heading("Porcentaje", text="Porcentaje de Aptitud")
        self.tabla.heading("Detalle", text="Evaluacion General")
        
        self.tabla.column("Nombre", width=220, anchor="w")
        self.tabla.column("Porcentaje", width=160, anchor="center")
        self.tabla.column("Detalle", width=420, anchor="w")
        self.tabla.pack(fill="both", expand=True, pady=(0, 12))

        # Contenedor inferior para botones secundarios
        frame_botones_inferior = tk.Frame(self.frame_resultados, bg="#ffffff")
        frame_botones_inferior.pack(fill="x")

        self.btn_exportar = self.crear_boton_redondeado(frame_botones_inferior, "Exportar Tabla a CSV", self.img_btn_secondary, self.exportar_csv)
        self.btn_exportar.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.btn_actualizar = self.crear_boton_redondeado(frame_botones_inferior, f"Buscar Actualizacion (v{self.VERSION_ACTUAL})", self.img_btn_primary, self.buscar_actualizacion)
        self.btn_actualizar.pack(side="right", fill="x", expand=True, padx=(5, 0))

    # --- HELPERS VISUALES: BOTONES ESTILO BOOTSTRAP ---
    def crear_forma_boton(self, color):
        """Genera una imagen solida en memoria con transparencia nativa."""
        img = tk.PhotoImage(width=12, height=12)
        img.blank() # Esto hace que TODA la imagen sea transparente desde el inicio
        
        # Pintamos solo el cuerpo del boton, dejando las 4 esquinas sin tocar (transparentes)
        for x in range(12):
            for y in range(12):
                if not ((x==0 and y==0) or (x==11 and y==0) or (x==0 and y==11) or (x==11 and y==11)):
                    img.put(color, (x, y))
        return img

    def crear_boton_redondeado(self, contenedor, texto, imagen_fondo, comando):
        """Construye un boton amigable utilizando labels compuestos."""
        btn = tk.Button(
            contenedor, text=texto, image=imagen_fondo, compound="center",
            command=comando, font=("Segoe UI", 10, "bold"), fg="#ffffff",
            relief="flat", bd=0, highlightthickness=0, activebackground=self.COLOR_BG,
            cursor="hand2", padx=15, pady=8
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
            respuesta = urllib.request.urlopen(self.URL_VERSION, timeout=5)
            version_nube = respuesta.read().decode('utf-8').strip()

            if float(version_nube) > float(self.VERSION_ACTUAL):
                if messagebox.askyesno("Actualizacion Disponible", f"Hay una nueva version ({version_nube}). ¿Deseas descargarla y reiniciar el programa ahora?"):
                    self.descargar_y_aplicar_actualizacion()
            else:
                messagebox.showinfo("Actualizado", "Ya tienes la version mas reciente instalada.")
        except Exception as e:
            messagebox.showerror("Error de red", "No se pudo comprobar la actualizacion. Verifica tu conexion.")

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
            messagebox.showerror("Error", f"Ocurrio un fallo al intentar actualizar:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MiniATSApp(root)
    root.mainloop()