import os
import re
from urllib.parse import urljoin
from playwright.sync_api import sync_playwright
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import sys
import io

# Protegim l'accés a buffer en entorns sense terminal (PyInstaller --windowed)
if sys.stdout and hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
if sys.stderr and hasattr(sys.stderr, "buffer"):
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

#ULTIMA ACTUALTZACIÓ: 2025-06-02
INVALID_CHARS = r"\\/:*?\"<>|"
PDF_KEYWORDS = [
    "pcap", "pcp", "ppt-", "prescripcions tecniques", "prescripcions tècniques",
    "pca", "plec clàusules", "clausulesadministratives", "plec", "pct", "ppt",
    "pliego administrativo", "pliego tecnico", "plec administratiu", "plec tècnic",
    "tècnic", "plec condicions", "normes reguladores"
]

def sanitize_folder_name(name: str) -> str:
    for ch in INVALID_CHARS:
        name = name.replace(ch, "_")
    return name

def is_relevant_pdf(text: str) -> bool:
    txt = text.lower()
    return any(k in txt for k in PDF_KEYWORDS)

def save_pdf_from_response(response, folder: str, url: str):
    ctype = response.headers.get('content-type', '')
    body = response.body()

    if response.ok and 'application/pdf' in ctype and body.startswith(b'%PDF'):
        content_disposition = response.headers.get('content-disposition', '')
        fname = None
        if 'filename=' in content_disposition:
            fname_match = re.search(r'filename="?([^"]+)"?', content_disposition)
            if fname_match:
                fname = fname_match.group(1)

        if not fname:
            fname = os.path.basename(url.split('?')[0]) or 'document.pdf'
        if not fname.lower().endswith('.pdf'):
            fname += '.pdf'

        path = os.path.join(folder, fname)
        with open(path, 'wb') as f:
            f.write(body)
        print(f"  Descarregat: {path}")
    else:
        print(f"  No és un PDF vàlid a {url} (ctype={ctype})")

def save_pdf_from_direct_url(page, url: str, folder: str):
    try:
        resp = page.request.get(url)
        save_pdf_from_response(resp, folder, url)
    except Exception as e:
        print(f"  Error descarregant directe {url}: {e}")

def descarrega_per_titol_estructura(page, folder_path):
    rellevants = []
    descarregats = set()

    rows = page.locator("app-documents-publicacio >> .row")
    for i in range(rows.count()):
        row = rows.nth(i)
        try:
            titol = row.locator(".col-md-4").inner_text().strip().lower()
            titol_normalitzat = titol.replace(":", "").strip()
            if titol_normalitzat in [
                "plec de clàusules administratives",
                "plec de prescripcions tècniques"
            ]:
                boto = row.locator(".col-md-8 >> button")
                if boto.count() == 0:
                    boto = row.locator(".col-md-8 >> a")

                if boto.count() > 0:
                    try:
                        with page.expect_download(timeout=30000) as download_info:  # Timeout augmentat a 30 segons
                            boto.first.click()
                        download = download_info.value
                        filename = sanitize_folder_name(download.suggested_filename)

                        if filename not in descarregats:
                            save_path = os.path.join(folder_path, filename)
                            download.save_as(save_path)
                            descarregats.add(filename)
                            print(f"[PDF] Descarregat via títol associat: {filename}")
                            rellevants.append((save_path, save_path))
                        else:
                            print(f"[INFO] Ja descarregat anteriorment: {filename}")
                    except Exception as e:
                        print(f"[ERROR] Error descarregant fitxer vinculat a: {titol} → {e}")
        except Exception as e:
            print(f"[ERROR] Error processant fila {i}: {e}")

    return rellevants

def process_annunci(page, base_url: str, folder_name: str) -> list[tuple[str, str]]:
    page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
    page.wait_for_timeout(2000)

    rellevants = []
    folder_path = os.path.join('Documents_Descarregats_Playwright', folder_name)
    os.makedirs(folder_path, exist_ok=True)

    anchors = page.query_selector_all('a[href]')
    print(f"[INFO] Trobats {len(anchors)} links <a> renderitzats.")

    for a in anchors:
        href = a.get_attribute('href')
        text = a.inner_text().strip()
        if not href:
            continue

        full_url = urljoin(base_url, href)
        if is_relevant_pdf(text):
            print(f"[PDF] PDF rellevant detectat via <a>: {full_url}")
            save_pdf_from_direct_url(page, full_url, folder_path)
            rellevants.append((full_url, href))

    buttons = page.query_selector_all('button, input[type="button"], input[type="submit"]')
    print(f"[INFO] Trobats {len(buttons)} botons <button> o <input>")

    for btn in buttons:
        text = btn.inner_text().strip()
        value = btn.get_attribute("value") or ""

        if (".pdf" in text.lower() or ".pdf" in value.lower()) and is_relevant_pdf(text + " " + value):
            print(f"[INFO] Intentant clic amb expect_download: {text}")
            try:
                with page.expect_download(timeout=10000) as download_info:
                    btn.click()
                download = download_info.value
                filename = sanitize_folder_name(download.suggested_filename)
                save_path = os.path.join(folder_path, filename)
                download.save_as(save_path)
                print(f"[PDF] PDF descarregat via download: {filename}")
                rellevants.append((save_path, save_path))
            except Exception as e1:
                print(f"[WARN] Fallada a expect_download per {text}: {e1}")
                print(f"[INFO] Intentant clic amb expect_response de fallback per: {text}")
                try:
                    with page.expect_response(lambda r: ".pdf" in r.url.lower(), timeout=10000) as resp_info:
                        btn.click()
                    response = resp_info.value
                    full_url = response.url
                    save_pdf_from_response(response, folder_path, full_url)
                    rellevants.append((full_url, full_url))
                except Exception as e2:
                    print(f"[WARN] Cap PDF en clicar botó: {text}. Errors: {e1} // {e2}")

    rellevants += descarrega_per_titol_estructura(page, folder_path)
    return rellevants



def main(excel_path: str):
    errors_desc = []
    df = pd.read_excel(excel_path)
    required = {'CODI_EXPEDIENT', 'ENLLAC_PUBLICACIO'}
    if not required.issubset(df.columns):
        print(f" Falten columnes a l'Excel ({required - set(df.columns)}).")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(user_agent="Mozilla/5.0")

        for _, row in df.iterrows():
            code = str(row['CODI_EXPEDIENT'])
            url = row['ENLLAC_PUBLICACIO']

            folder_name = sanitize_folder_name(code)
            page = context.new_page()
            page.goto(url, wait_until='networkidle')

            try:
                page.locator("button:has-text('Accepta')").click(timeout=3000)
            except:
                pass

            # Navegació per pestanyes
            try:
                page.click("text=Anunci de licitació", timeout=5000)
                page.wait_for_load_state('networkidle')
            except:
                try:
                    page.click("text=Anuncio de licitación", timeout=5000)
                    page.wait_for_load_state('networkidle')
                except:
                    try:
                        page.click("text=Adjudicació", timeout=5000)
                        page.wait_for_load_state('networkidle')
                        print(f" No trobat 'Anunci de licitació' ni 'Anuncio de licitación', s'intenta 'Adjudicació' per expedient {code}.")
                    except:
                        print(f" No trobat cap pestanya vàlida per expedient {code}.")

            pdfs = process_annunci(page, page.url, folder_name)

            if pdfs:
                print(f"  {len(pdfs)} PDF(s) rellevant(s) descarregats per {code}")
            else:
                print(f" Cap PDF rellevant per a {code} → no es crea cap carpeta.")
                errors_desc.append(code)

            page.close()

        context.close()
        browser.close()

    # Generació del resum Excel
    base_path = "Documents_Descarregats_Playwright"
    resultats = []

    # Regexs més flexibles per identificar PPT i PCAP
    ppt_keywords = (
        r'ppt[\s_\-\.]*|ptt[\s_\-\.]*|pptx[\s_\-\.]*|'
        r'plec[\s_\-\.]*t[eè]cnic|plec[\s_\-\.]*prescripcions|prescripcions|'
        r'pliego[\s_\-\.]*t[eé]cnico|mem[oò]ria|memoria|'
        r'pleg[\s_\-\.]*t[eè]cnic|tecnic|tecnica|tecnicas|'
        r'prescripcions[\s_\-\.]*tecnic|prescripcions[\s_\-\.]*tecnica|'
        r'pt[\s_\-\.]*t[eè]cnic|pt[\s_\-\.]*tecnica'
    )
    pcap_keywords = (
        r'pca[pb]?[\s_\-\.]*|pcp[\s_\-\.]*|pcap[\s_\-\.]*|'
        r'plec[\s_\-\.]*administratiu|plec[\s_\-\.]*condicions|'
        r'condicions[\s_\-\.]*administratives|condicions[\s_\-\.]*particulars|'
        r'plec[\s_\-\.]*cl[aà]usules|cl[aà]usules|'
        r'pliego[\s_\-\.]*administrativo|pliego[\s_\-\.]*condiciones|'
        r'condiciones[\s_\-\.]*administrativas|condiciones[\s_\-\.]*particulares|'
        r'normes[\s_\-\.]*reguladores|administratiu|administrativo|'
        r'particulars|particulares|pc-adm|pc[\s_\-\.]*adm'
    )

    for carpeta in os.listdir(base_path):
        ruta_carpeta = os.path.join(base_path, carpeta)
        if not os.path.isdir(ruta_carpeta):
            continue

        ppt_trobat = False
        pcap_trobat = False
        altres_trobat = False
        n_docs_identificats = 0  # Nova columna: nombre de documents identificats (PPT, PCAP o ALTRES)
        n_docs_total = 0         # Nova columna: nombre total de documents a la carpeta

        for fitxer in os.listdir(ruta_carpeta):
            n_docs_total += 1
            nom = fitxer.lower()
            es_ppt = bool(re.search(ppt_keywords, nom, re.IGNORECASE))
            es_pcap = bool(re.search(pcap_keywords, nom, re.IGNORECASE))
            if es_ppt:
                ppt_trobat = True
                n_docs_identificats += 1
            if es_pcap:
                pcap_trobat = True
                n_docs_identificats += 1
            if not es_ppt and not es_pcap:
                altres_trobat = True

        # Nova columna: tots identificats?
        tots_identificats = n_docs_identificats == n_docs_total

        # Nova columna: més de 2 documents?
        dos_docs_o_mes = n_docs_total >= 2

        resultats.append({
            "CODI_EXPEDIENT": carpeta,
            "TROBAT_PPT": ppt_trobat,
            "TROBAT_PCAP": pcap_trobat,
            "TROBAT_ALTRES": altres_trobat,
            "N_DOCS_IDENTIFICATS": n_docs_identificats,
            "N_DOCS_TOTAL": n_docs_total,
            "TOTS_IDENTIFICATS": tots_identificats,  # Nova columna booleana
            "DOS_DOCS_O_MES": dos_docs_o_mes         # Nova columna booleana amb el nou títol
        })

    df_resultats = pd.DataFrame(resultats)
    with pd.ExcelWriter("Resum_descarregues.xlsx") as writer:
        df_resultats.to_excel(writer, index=False, sheet_name="Resum")
        if errors_desc:
            df_errors = pd.DataFrame({"CODI_EXPEDIENT": errors_desc})
            df_errors.to_excel(writer, index=False, sheet_name="Errors")
    print(" Resum generat: Resum_descarregues.xlsx")

if __name__ == '__main__':
    # Colors iguals que la teva app principal
    color_fons = "#D6EAF8"
    color_botons = "#0078D7"
    color_botons_actiu = "#005A9E"

    def descarrega_predefinit():
        root.destroy()
        main('Contractes 2024_només amb publicitat.xlsx')

    def descarrega_personalitzada():
        file_path = filedialog.askopenfilename(
            title="Selecciona un fitxer Excel",
            filetypes=[("Fitxers Excel", "*.xlsx *.xls")]
        )
        if file_path:
            root.destroy()
            main(file_path)

    root = tk.Tk()
    root.title("Descarrega PDFs")
    root.configure(bg=color_fons)
    root.geometry("420x200")

    label = tk.Label(
        root,
        text="Selecciona una opció per descarregar:",
        bg=color_fons,
        fg="#222",
        font=("Helvetica", 13, "bold")
    )
    label.pack(pady=(25, 15))

    btn1 = tk.Button(
        root, text="Descarrega amb fitxer predefinit",
        command=descarrega_predefinit,
        bg=color_botons, fg="white", font=("Helvetica", 12, "bold"),
        width=30, height=2, relief="flat", activebackground=color_botons_actiu
    )
    btn1.pack(pady=5)

    btn2 = tk.Button(
        root, text="Descarrega a partir d'un altre document",
        command=descarrega_personalitzada,
        bg=color_botons, fg="white", font=("Helvetica", 12, "bold"),
        width=30, height=2, relief="flat", activebackground=color_botons_actiu
    )
    btn2.pack(pady=5)

    root.mainloop()


