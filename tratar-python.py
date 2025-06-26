import os
import re
import time
import subprocess
import pythoncom
import gc
from win32com.client import gencache, GetActiveObject
from pywintypes import com_error

# Configurações
ROOT_DIRS = [
    r"Z:\Pedidos\Sign - Lona",
    r"Z:\Pedidos\Sign - Adesivos",
    r"Z:\Pedidos\Digital Colorido",
    r"Z:\Pedidos\Digital PB",
]
ALLOWED_EXT = {".jpg", ".jpeg", ".png", ".tif", ".cdr"}
MAX_FILES_PER_SUBFOLDER = 5  # se mais que isto, ignora a subpasta
GMS_PROJECT = "Graficonauta"
GMS_MODULE = "Dump"
GMS_PROCEDURE = "TratamentoAutomatico"
GMS_REFUGO_POR_FONTE = "RefugarPorFonteFaltando"
POLL_TIMEOUT = 60  # segundos
AUTO_CLOSE_MULTIPLE = True  # fecha todos os docs abertos se mais de um estiver aberto
IGNORED_SYSTEM_FONTS = {"arial", "calibri"}
ONLY_DIGITAL_CDR = True  # se True, só processa arquivos .cdr com 'impressao-digital' no nome

# --- Lista de palavras-chave para CDRs ---
CDR_KEYWORDS = ["impressao-digital", "banner", "lona", "vinil-adesivo", "papel-adesivo", "adesivo"]  # adicione o que precisar

# Inicializa COM e CorelDRAW tipado
pythoncom.CoInitialize()
app = gencache.EnsureDispatch("CorelDRAW.Application")
app.Visible = True
# Suprime diálogos de alerta (incluindo fontes faltando)
try:
    app.Preferences.Application.EnableAlerts = False
except AttributeError:
    pass


def RefugarPorFonteFaltando(caminho_arquivo: str, missing_fonts: list):
    """
    Executa macro de refugo e loga fontes realmente faltantes.
    """
    fontes_str = ",".join(missing_fonts)
    run_macro(GMS_PROJECT, GMS_MODULE, GMS_REFUGO_POR_FONTE, POLL_TIMEOUT, caminho_arquivo, fontes_str)
    print(f"Rejeitado {caminho_arquivo}: fontes faltando -> {missing_fonts}")


import pythoncom

def run_macro(project: str, module: str, procedure: str, timeout: int = POLL_TIMEOUT, *args, doc=None):
    full = f"{module}.{procedure}"
    app.GMSManager.RunMacro(project, full, *args)

    start = time.time()
    target = doc or app.ActiveDocument

    while True:
        # 1) se existir flag de busy, check nela
        busy = None
        try:
            busy = getattr(app, "Busy", None)
            # ou, se sua versão oferece: busy = app.GMSManager.IsBusy
        except Exception:
            busy = None

        if busy is not None:
            if not busy:
                break
        else:
            # 2) fallback tradicional de polling Pages.Count
            try:
                _ = target.Pages.Count
                break
            except pythoncom.com_error:
                pass

        # 3) pump COM messages para evitar “cansaço” do loop
        pythoncom.PumpWaitingMessages()

        # 4) timeout check
        if time.time() - start > timeout:
            raise TimeoutError(f"Macro {full} não finalizou em {timeout}s")

        # 5) pequena pausa para não saturar CPU/Corel
        time.sleep(0.2)



def get_pdf_page_count(pdf_path: str) -> int:
    try:
        output = subprocess.check_output(["mutool", "info", pdf_path], text=True)
        for line in output.splitlines():
            if line.lower().startswith("pages:"):
                return int(line.split()[1])
    except Exception:
        pass
    return 0


def processar_arquivo(caminho_arquivo: str):
    # Fecha múltiplos docs se habilitado
    if AUTO_CLOSE_MULTIPLE:
        try:
            docs = list(app.Documents)
            if len(docs) > 1:
                for d in docs:
                    d.Close()
        except Exception:
            pass

    basename = os.path.basename(caminho_arquivo)
    ext = os.path.splitext(basename)[1].lower()
    args = []

    # Lógica específica para .cdr
    if ext == ".cdr":
        # Se only_digital habilitado e NÃO encontrar nenhuma das keywords, ignora
        if ONLY_DIGITAL_CDR:
            lower_name = basename.lower()
            if not any(keyword in lower_name for keyword in CDR_KEYWORDS):
                print(f"Ignorando CDR sem keywords: {basename}")
                return

        # Abre o documento UMA ÚNICA VEZ
        doc = app.OpenDocument(caminho_arquivo)
        try:
            # Checa fontes faltantes
            missing_fonts = []
            try:
                count = doc.MissingFontListCount
                for i in range(count):
                    missing_fonts.append(doc.MissingFontList(i).Name)
            except AttributeError:
                try:
                    missing_fonts = [
                        f.Name for f in doc.Fonts if not getattr(f, 'IsInstalled', True)
                    ]
                except Exception:
                    missing_fonts = []

            real_missing = [
                f for f in missing_fonts
                if f.lower() not in IGNORED_SYSTEM_FONTS
            ]
            if real_missing:
                RefugarPorFonteFaltando(caminho_arquivo, real_missing)
                return

            # Se passou na checagem, dispara a macro diretamente no mesmo doc
            run_macro(
                GMS_PROJECT,
                GMS_MODULE,
                GMS_PROCEDURE,
                POLL_TIMEOUT,
                doc=doc
            )
            print(f"{caminho_arquivo} tratado")
        finally:
            doc.Close()
        return

    # Parâmetros para TIF Impressao-Digital
    if ext == ".tif" and "impressao-digital" in basename.lower():
        name_without_ext = os.path.splitext(basename)[0]
        # Remove sufixo numérico (ex: -01, 01) no final do nome
        m = re.match(r"^(.*?)(?:-)?\d+$", name_without_ext, re.IGNORECASE)
        if m:
            base_name_without_numbers = m.group(1)
            pdf_name = f"{base_name_without_numbers}.pdf"
        else:
            pdf_name = f"{name_without_ext}.pdf"

        dirpath = os.path.dirname(caminho_arquivo)
        pdf_path = os.path.join(dirpath, pdf_name)

        if os.path.isfile(pdf_path):
            page_count = get_pdf_page_count(pdf_path)
            print(f"PDF associado encontrado: {pdf_path}")
            print(f"Número de páginas do PDF: {page_count}")
            args = [True, page_count]
        else:
            print(f"ATENÇÃO: PDF não encontrado - {pdf_path}")
            args = [False, 0]

    # Abre e processa outros arquivos
    try:
        doc = app.OpenDocument(caminho_arquivo)
        # Espera doc carregar
        time.sleep(3)
        run_macro(GMS_PROJECT, GMS_MODULE, GMS_PROCEDURE, POLL_TIMEOUT, *args, doc=doc)
        print(f"{caminho_arquivo} tratado com argumentos: {args}")
    except Exception as e:
        print(f"Erro ao processar {basename}: {str(e)}")
    finally:
        return


def um_arquivo_por_subpasta(root_dirs):
    for raiz in root_dirs:
        for dirpath, dirnames, filenames in os.walk(raiz):
            abs_dirpath = os.path.abspath(dirpath)
            if abs_dirpath in [os.path.abspath(p) for p in root_dirs]:
                continue
            if any(name.lower() == "observacoes.txt" for name in filenames):
                continue
            valid_files = [f for f in filenames if os.path.splitext(f)[1].lower() in ALLOWED_EXT]
            if MAX_FILES_PER_SUBFOLDER and len(valid_files) > MAX_FILES_PER_SUBFOLDER:
                continue
            if not valid_files:
                continue
            yield dirpath, os.path.join(dirpath, valid_files[0])


def ensure_corel_app():
    """
    Garante que exista uma instância viva do CorelDRAW.
    Se não encontrar via ROT, cria uma nova.
    Também (re)inicializa COM e suprime alertas.
    """
    try:
        # Tenta pegar instância já aberta
        app = GetActiveObject("CorelDRAW.Application")
    except com_error:
        # Não há instância: inicializa COM e cria uma
        pythoncom.CoInitialize()
        app = gencache.EnsureDispatch("CorelDRAW.Application")
        app.Visible = True
    # Suprime alerts (tenta ambas as propriedades)
    try:
        app.DisplayAlerts = False
    except AttributeError:
        try:
            app.Preferences.Application.EnableAlerts = False
        except Exception:
            pass
    return app

def main_loop():
    global app
    while True:
        # 1) Garante Corel aberto
        try:
            # um simples acesso para forçar exception se estiver fechado
            _ = app.Version
        except Exception:
            print("CorelDRAW não está respondendo. Reiniciando instância...")
            app = ensure_corel_app()

        # 2) Varre e processa arquivos
        for subpasta, arquivo in um_arquivo_por_subpasta(ROOT_DIRS):
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Processando: {arquivo}")
            try:
                processar_arquivo(arquivo)
            except Exception as e:
                print(f"Erro no arquivo {arquivo}: {e}")
                # Se der erro COM grave, reinicia o Corel para a próxima iteração
                if isinstance(e, com_error):
                    app = ensure_corel_app()
        # 3) Liberar memória e evitar vazamentos
        gc.collect()
        # 4) Pausa antes da próxima iteração
        time.sleep(3)

if __name__ == "__main__":
    # 0) Cria/garante a instância de Corel antes de tudo
    app = ensure_corel_app()
    main_loop()