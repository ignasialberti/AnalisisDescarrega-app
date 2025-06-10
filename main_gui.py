# -*- coding: utf-8 -*-
import time
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk  # Per utilitzar estils moderns
import threading
import subprocess
import os
import sys
import tkinter.scrolledtext as scrolledtext
from utils_config import carregar_clau_deepseek
from utils_config import guardar_clau_deepseek # Importem el mòdul de configuració
from tqdm import tqdm
import sys, os
def resource_path(rel_path):
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.abspath(".")
    return os.path.join(base, rel_path)


# Importem els mòduls dels scripts (pressuposem que es troben al mateix directori)
try:
    import AnalisisDeepseek
except ImportError:
    AnalisisDeepseek = None

try:
    import AnalisisLocalOllama
except ImportError:
    AnalisisLocalOllama = None

try:
    import DescarregaPlaywright
except ImportError:
    DescarregaPlaywright = None

try:
    import DescarregaSelenium
except ImportError:
    DescarregaSelenium = None

# Referències globals als processos de descàrrega
proc_desc_playwright = None
proc_desc_selenium = None

# Variables globals per als fils d'anàlisi automàtica
thread_deepseek = None
thread_ollama = None

def run_analisis_deepseek():
    global thread_deepseek
    if AnalisisDeepseek is None:
        messagebox.showerror("Error", "No s'ha pogut importar AnalisisDeepseek.py")
        return
    # Ja no preguntem, només cridem la funció principal
    thread_deepseek = threading.Thread(target=AnalisisDeepseek.executar_analisi, daemon=True)
    thread_deepseek.start()

def stop_analisi_deepseek():
    global thread_deepseek
    # Aquí pots implementar un flag d'aturada si vols aturar el fil de veritat
    messagebox.showinfo("Aturat", "Procés d'anàlisi Deepseek aturat (implementa la lògica d'aturada real si cal).")

def run_analisis_local_ollama():
    global thread_ollama
    if AnalisisLocalOllama is None:
        messagebox.showerror("Error", "No s'ha pogut importar AnalisisLocalOllama.py")
        return
    # Ja no preguntem, només cridem la funció principal
    thread_ollama = threading.Thread(target=AnalisisLocalOllama.executar_analisi, daemon=True)
    thread_ollama.start()

def stop_analisi_ollama():
    global thread_ollama
    # Aquí pots implementar un flag d'aturada si vols aturar el fil de veritat
    messagebox.showinfo("Aturat", "Procés d'anàlisi Ollama aturat (implementa la lògica d'aturada real si cal).")

def run_descarrega_playwright():
    global proc_desc_playwright
    try:
        script_path = os.path.join(os.path.dirname(__file__), "DescarregaPlaywright.py")
        print(f"[INFO] Executant: {script_path}")
        proc_desc_playwright = subprocess.Popen(
            ["python", script_path],
            stdout=subprocess.PIPE,  # Captura stdout
            stderr=subprocess.PIPE,  # Captura stderr
            text=True,  # Decodifica la sortida com a text
            encoding="utf-8"  # Força la codificació UTF-8
        )

        import time

        # Llegeix la sortida del subprocess i imprimeix-la al terminal de la GUI
        def llegir_sortida():
            print("[INFO] Capturant sortida del subprocess...")
            while proc_desc_playwright.poll() is None:  # Comprova si el subprocess encara s'executa
                # Llegeix línies de stdout
                for line in iter(proc_desc_playwright.stdout.readline, ''):
                    try:
                        print(line.strip())
                    except Exception as e:
                        print(f"[ERROR] Error processant stdout: {e}")
                    time.sleep(0.1)  # Evita bloquejar el bucle

                # Llegeix línies de stderr
                for line in iter(proc_desc_playwright.stderr.readline, ''):
                    try:
                        print(f"[ERROR] {line.strip()}")
                    except Exception as e:
                        print(f"[ERROR] Error processant stderr: {e}")
                    time.sleep(0.1)  # Evita bloquejar el bucle

        threading.Thread(target=llegir_sortida, daemon=True).start()

    except Exception as e:
        print(f"[ERROR] No s'ha pogut executar DescarregaPlaywright.py: {e}")

def stop_descarrega_playwright():
    global proc_desc_playwright
    if proc_desc_playwright and proc_desc_playwright.poll() is None:
        proc_desc_playwright.terminate()
        messagebox.showinfo("Aturat", "Descàrrega Playwright aturada.")
    else:
        messagebox.showinfo("Info", "No hi ha cap descàrrega Playwright activa.")

def run_descarrega_selenium():
    global proc_desc_selenium
    try:
        script_path = os.path.join(os.path.dirname(__file__), "DescarregaSelenium.py")
        proc_desc_selenium = subprocess.Popen(
            ["python", script_path]
        )
    except Exception as e:
        messagebox.showerror("Error", f"No s'ha pogut executar DescarregaSelenium.py:\n{e}")

def stop_descarrega_selenium():
    global proc_desc_selenium
    if proc_desc_selenium and proc_desc_selenium.poll() is None:
        proc_desc_selenium.terminate()
        messagebox.showinfo("Aturat", "Descàrrega Selenium aturada.")
    else:
        messagebox.showinfo("Info", "No hi ha cap descàrrega Selenium activa.")

# Afegiu aquest codi a la part on definiu la GUI principal (on hi ha les altres pestanyes)
def executar_analisi_local_rag():
    ruta = os.path.join(os.path.dirname(__file__), "RAG", "AnalisisRAGLocalOllama.py")
    subprocess.Popen([sys.executable, ruta])

def executar_creacio_embeddings():
    ruta = os.path.join(os.path.dirname(__file__), "RAG", "embeddings.py")
    if not os.path.exists(ruta):
        print(f"[ERROR] El fitxer embeddings.py no existeix a la ruta: {ruta}")
        return
    print(f"[INFO] Executant embeddings.py a la ruta: {ruta}")
    try:
        proc_embeddings = subprocess.Popen(
            [sys.executable, ruta],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8"
        )

        # Monitoritza el procés fins que acabi
        def monitoritzar_proces():
            print("[INFO] Monitoritzant el procés embeddings.py...")
            proc_embeddings.wait()  # Espera que el procés acabi
            print("[INFO] El procés embeddings.py ha finalitzat.")
            print("✅ Fitxer index.faiss generat amb èxit (comprova la carpeta de sortida).")

        threading.Thread(target=monitoritzar_proces, daemon=True).start()

    except Exception as e:
        print(f"[ERROR] No s'ha pogut executar embeddings.py: {e}")

# Creació de la finestra principal de l'aplicació
root = tk.Tk()
root.title("Aplicació d'Anàlisi")
root.resizable(False, False)
root.configure(bg="#D6EAF8")  # Blau clar com a color de fons

notebook = ttk.Notebook(root, style="TNotebook")
notebook.pack(padx=10, pady=10, expand=True, fill="both")

# Configuració d'estil amb ttk.Style
style = ttk.Style()
style.theme_use("clam")

# Colors personalitzats
color_fons = "#D6EAF8"      # Blau clar per al fons general i pestanyes
color_botons = "#0078D7"    # Blau fort per als botons
color_botons_actiu = "#005A9E"

# Estil per als botons
style.configure("TButton", background=color_botons, foreground="white", font=("Helvetica", 12, "bold"), padding=10)
style.map("TButton", background=[("active", color_botons_actiu)])

# Estil per a les pestanyes del notebook
style.configure("TNotebook", background=color_fons, borderwidth=0)
style.configure("TNotebook.Tab", background=color_fons, foreground="#154360", font=("Helvetica", 11, "bold"), padding=[10, 5])
style.map("TNotebook.Tab",
          background=[("selected", "#85C1E9")],   # Blau una mica més intens per la pestanya seleccionada
          foreground=[("selected", "#154360")])

# Paràmetres d'estil
button_height = 2  # Alçada més petita per als botons
button_width = 20  # Amplada més raonable per als botons blaus
stop_width = 7     # Amplada per als botons STOP

# --- Pestanya DESCÀRREGA ---
frame_desc = tk.Frame(notebook, bg=color_fons)

for idx, (text, run_cmd, stop_cmd) in enumerate([
    ("Descarrega Playwright", run_descarrega_playwright, stop_descarrega_playwright),
    ("Descarrega Selenium", run_descarrega_selenium, stop_descarrega_selenium)
]):
    row = tk.Frame(frame_desc, bg=color_fons)
    row.grid_columnconfigure(0, weight=1)  # Espaiador esquerra
    row.grid_columnconfigure(1, weight=0)  # Botó principal
    row.grid_columnconfigure(2, weight=0)  # Botó STOP
    row.grid_columnconfigure(3, weight=1)  # Espaiador dreta

    # Espaiador esquerra (col 0)
    tk.Label(row, bg=color_fons).grid(row=0, column=0, sticky="ew")

    # Botó principal (col 1)
    btn = tk.Button(
        row, text=text, command=run_cmd,
        bg=color_botons, fg="white", font=("Helvetica", 12, "bold"),
        width=button_width, height=button_height, relief="flat", activebackground=color_botons_actiu
    )
    btn.grid(row=0, column=1, padx=(0, 5), pady=8, sticky="ew")

    # Botó STOP (col 2)
    btn_stop = tk.Button(
        row, text="STOP", command=stop_cmd,
        bg="#E74C3C", fg="white", font=("Helvetica", 12, "bold"),
        width=stop_width, height=button_height, relief="flat", activebackground="#922B21"
    )
    btn_stop.grid(row=0, column=2, padx=(5, 0), pady=8, sticky="ew")

    # Espaiador dreta (col 3)
    tk.Label(row, bg=color_fons).grid(row=0, column=3, sticky="ew")

    row.pack(fill="x", padx=120, pady=8)  # Centra la fila i redueix amplada

notebook.add(frame_desc, text="Descàrrega")

# --- Pestanya ANÀLISI ---
frame_ana = tk.Frame(notebook, bg=color_fons)

for idx, (text, run_cmd, stop_cmd) in enumerate([
    ("Anàlisi Deepseek", run_analisis_deepseek, stop_analisi_deepseek),
    ("Anàlisi Ollama (local)", run_analisis_local_ollama, stop_analisi_ollama)
]):
    row = tk.Frame(frame_ana, bg=color_fons)
    row.grid_columnconfigure(0, weight=1)
    row.grid_columnconfigure(1, weight=0)
    row.grid_columnconfigure(2, weight=0)
    row.grid_columnconfigure(3, weight=1)

    tk.Label(row, bg=color_fons).grid(row=0, column=0, sticky="ew")
    btn = tk.Button(
        row, text=text, command=run_cmd,
        bg=color_botons, fg="white", font=("Helvetica", 12, "bold"),
        width=button_width, height=button_height, relief="flat", activebackground=color_botons_actiu
    )
    btn.grid(row=0, column=1, padx=(0, 5), pady=8, sticky="ew")
    btn_stop = tk.Button(
        row, text="STOP", command=stop_cmd,
        bg="#E74C3C", fg="white", font=("Helvetica", 12, "bold"),
        width=stop_width, height=button_height, relief="flat", activebackground="#922B21"
    )
    btn_stop.grid(row=0, column=2, padx=(5, 0), pady=8, sticky="ew")
    tk.Label(row, bg=color_fons).grid(row=0, column=3, sticky="ew")
    row.pack(fill="x", padx=120, pady=8)

notebook.add(frame_ana, text="Anàlisi")

# --- Pestanya RAG LOCAL ---
frame_rag = tk.Frame(notebook, bg=color_fons)

for idx, (text, run_cmd, stop_cmd) in enumerate([
    ("Anàlisi Local amb RAG", executar_analisi_local_rag, lambda: None),
    ("Creació de embeddings", executar_creacio_embeddings, lambda: None)
]):
    row = tk.Frame(frame_rag, bg=color_fons)
    row.grid_columnconfigure(0, weight=1)  # Espaciador izquierda
    row.grid_columnconfigure(1, weight=3)  # Botón principal ocupa més espai
    row.grid_columnconfigure(2, weight=1)  # Botón STOP ocupa menys
    row.grid_columnconfigure(3, weight=1)  # Espaciador derecha

    tk.Label(row, bg=color_fons).grid(row=0, column=0, sticky="ew")

    btn = tk.Button(
        row, text=text, command=run_cmd,
        bg=color_botons, fg="white", font=("Helvetica", 12, "bold"),
        width=button_width, height=button_height, relief="flat", activebackground=color_botons_actiu
    )
    btn.grid(row=0, column=1, padx=(0, 5), pady=8, sticky="ew")

    btn_stop = tk.Button(
        row, text="STOP", command=stop_cmd,
        bg="#E74C3C", fg="white", font=("Helvetica", 12, "bold"),
        width=stop_width, height=button_height, relief="flat", activebackground="#922B21"
    )
    btn_stop.grid(row=0, column=2, padx=(5, 0), pady=8, sticky="ew")

    tk.Label(row, bg=color_fons).grid(row=0, column=3, sticky="ew")

    row.pack(fill="x", padx=40, pady=8)  # Redueix el marge lateral

notebook.add(frame_rag, text="RAG Local")

# Pestanya AJUDA
frame_help = tk.Frame(notebook, bg=color_fons)
text_ajuda = tk.Text(frame_help, wrap=tk.WORD, height=20, width=80, bg=color_fons, font=("Helvetica", 11))
ajuda = """
🛈 INSTRUCCIONS D’ÚS

📥 DESCÀRREGA:
 - Playwright i Selenium descarreguen documents a partir dels enllaços especificats en un Excel.
 - Fes clic a "Descarrega" per iniciar. STOP per aturar.

🧠 ANÀLISI:
 - Anàlisi Deepseek: envia els documents a una API remota. Definir la clau API a la pestanya Configuració.
 - Anàlisi Ollama: utilitza models locals.

🔍 RAG LOCAL:
 - Crea embeddings i realitza anàlisi contextual amb base local.
 - Útil per detectar clàusules específiques en plecs llargs.

📁 Els resultats es guarden automàticament en carpetes amb el nom de l’expedient.
📄 El resum de l’execució es genera en format Excel.
"""
text_ajuda.insert(tk.END, ajuda)
text_ajuda.config(state='disabled')
text_ajuda.pack(padx=10, pady=10, fill="both")
notebook.add(frame_help, text="Ajuda")

# === PESTANYA CONFIGURACIÓ ===
frame_config = tk.Frame(notebook, bg=color_fons)

label_clau = tk.Label(frame_config, text="🔑 Introdueix la clau DeepSeek API:", bg=color_fons, font=("Helvetica", 11))
label_clau.pack(pady=10)
label_clau = tk.Label(frame_config, text="Registre i obtenció a https://api-docs.deepseek.com/:", bg=color_fons, font=("Helvetica", 11))
label_clau.pack(pady=10)

entry_clau = tk.Entry(frame_config, font=("Helvetica", 11), width=60)
entry_clau.pack(pady=5)
entry_clau.insert(0, carregar_clau_deepseek())  # Carregar la clau existent

def guardar_clau():
    clau = entry_clau.get().strip()
    guardar_clau_deepseek(clau)  # Guardar la clau
    messagebox.showinfo("Desat", "Clau API guardada correctament.")

btn_guardar = tk.Button(frame_config, text="Desar clau", command=guardar_clau,
          bg=color_botons, fg="white", font=("Helvetica", 12, "bold"),
          width=20, height=2, relief="flat", activebackground=color_botons_actiu)
btn_guardar.pack(pady=10)

notebook.add(frame_config, text="🔧 Configuració")

# Afegim un terminal de sortida (ScrollText)
frame_terminal = tk.Frame(root, bg=color_fons)
text_terminal = scrolledtext.ScrolledText(frame_terminal, wrap=tk.WORD, height=15, width=90, state='disabled', font=("Courier", 10))
text_terminal.pack(padx=10, pady=10)
frame_terminal.pack(padx=10, pady=(0, 2), fill="both")  # Abans: pady=(0, 10)

# Redirigir stdout i stderr al terminal de la GUI
class TerminalOutput:
    def write(self, message):
        text_terminal.configure(state='normal')
        text_terminal.insert(tk.END, message)
        text_terminal.see(tk.END)
        text_terminal.configure(state='disabled')

    def flush(self):  # Per compatibilitat
        pass

sys.stdout = TerminalOutput()
sys.stderr = TerminalOutput()

# Centrar la finestra i definir la mida
window_width = 700  # Reduir l'amplada de la finestra
window_height = 600  # Reduir l'alçada de la finestra
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_top = int(screen_height / 2 - window_height / 2)
position_right = int(screen_width / 2 - window_width / 2)
root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

# Ajustar la mida dels botons
button_height = 3  # Augmentar l'alçada dels botons
button_width = 25  # Augmentar l'amplada dels botons
stop_width = 10    # Augmentar l'amplada dels botons STOP

# Actualitzar els botons amb les noves dimensions
for idx, (text, run_cmd, stop_cmd) in enumerate([
    ("Descarrega Playwright", run_descarrega_playwright, stop_descarrega_playwright),
    ("Descarrega Selenium", run_descarrega_selenium, stop_descarrega_selenium)
]):
    btn = tk.Button(
        row, text=text, command=run_cmd,
        bg=color_botons, fg="white", font=("Helvetica", 12, "bold"),
        width=button_width, height=button_height, relief="flat", activebackground=color_botons_actiu
    )
    btn_stop = tk.Button(
        row, text="STOP", command=stop_cmd,
        bg="#E74C3C", fg="white", font=("Helvetica", 12, "bold"),
        width=stop_width, height=button_height, relief="flat", activebackground="#922B21"
    )

# Icona de la finestra (si tens el fitxer logo.ico)
try:
    root.iconbitmap("logo.ico")
except Exception:
    pass

# Inici del bucle principal de Tkinter
root.mainloop()
