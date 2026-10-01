#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word -> PDF em lote
===================
Escolhe uma pasta e converte todos os ficheiros .doc e .docx para PDF.

Motores de conversão (usa o primeiro disponível):
  1. Microsoft Word (só Windows, requer:  pip install pywin32)
  2. LibreOffice   (Windows, macOS, Linux - gratuito)

Utilização:
  python word_para_pdf.py                 -> abre a janela para escolher a pasta
  python word_para_pdf.py "C:\\pasta"      -> converte diretamente (sem janela)

Os PDF são guardados numa subpasta "PDF" dentro da pasta escolhida.
"""

import os
import sys
import shutil
import platform
import threading
import subprocess
from pathlib import Path

EXTENSOES = {".doc", ".docx"}
PASTA_SAIDA = "PDF"
FILTRO_PDF_SEM_PERDAS = (
    'pdf:writer_pdf_Export:{'
    '"UseLosslessCompression":{"type":"boolean","value":"true"},'
    '"ReduceImageResolution":{"type":"boolean","value":"false"}}'
)


# --------------------------------------------------------------------------
# Deteção de motores
# --------------------------------------------------------------------------
def encontrar_libreoffice():
    for nome in ("soffice", "libreoffice"):
        caminho = shutil.which(nome)
        if caminho:
            return caminho
    candidatos = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidatos:
        if os.path.exists(c):
            return c
    return None


def word_disponivel():
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
        return True
    except ImportError:
        return False


# --------------------------------------------------------------------------
# Recolha de ficheiros
# --------------------------------------------------------------------------
def listar_ficheiros(pasta: Path, subpastas: bool):
    if subpastas:
        candidatos = pasta.rglob("*")
    else:
        candidatos = pasta.glob("*")
    ficheiros = []
    for f in candidatos:
        if not f.is_file():
            continue
        if f.suffix.lower() not in EXTENSOES:
            continue
        if f.name.startswith("~$"):          # ficheiros temporários do Word
            continue
        if PASTA_SAIDA in f.relative_to(pasta).parts[:-1]:
            continue                          # ignora a própria pasta de saída
        ficheiros.append(f)
    return sorted(ficheiros)


def destino_pdf(ficheiro: Path, raiz: Path) -> Path:
    rel = ficheiro.relative_to(raiz).with_suffix(".pdf")
    destino = raiz / PASTA_SAIDA / rel
    destino.parent.mkdir(parents=True, exist_ok=True)
    return destino


# --------------------------------------------------------------------------
# Conversão
# --------------------------------------------------------------------------
def converter_com_word(ficheiros, raiz, log, progresso):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    ok = falhas = 0
    try:
        for i, f in enumerate(ficheiros, 1):
            destino = destino_pdf(f, raiz)
            try:
                doc = word.Documents.Open(
                    str(f.resolve()), ReadOnly=True,
                    AddToRecentFiles=False, ConfirmConversions=False)
                # ExportAsFixedFormat preserva melhor imagens com transparência
                # e símbolos do que SaveAs2 (evita fundos escuros/pretos).
                doc.ExportAsFixedFormat(
                    OutputFileName=str(destino.resolve()),
                    ExportFormat=17,          # 17 = PDF
                    OpenAfterExport=False,
                    OptimizeFor=0,            # 0 = qualidade de impressão
                    Range=0,                  # documento completo
                    Item=0,
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=1,        # marcadores a partir dos títulos
                    DocStructureTags=True,
                    BitmapMissingFonts=False,
                    UseISO19005_1=False)      # PDF/A pode causar fundos pretos
                doc.Close(False)
                log(f"OK      {f.relative_to(raiz)}")
                ok += 1
            except Exception as e:
                log(f"ERRO    {f.relative_to(raiz)}  ({e})")
                falhas += 1
            progresso(i, len(ficheiros))
    finally:
        word.Quit()
        pythoncom.CoUninitialize()
    return ok, falhas


def converter_com_libreoffice(soffice, ficheiros, raiz, log, progresso):
    ok = falhas = 0
    for i, f in enumerate(ficheiros, 1):
        destino = destino_pdf(f, raiz)
        try:
            r = None
            # Compressão sem perdas: JPEG não suporta transparência (fundo preto)
            for filtro in (FILTRO_PDF_SEM_PERDAS, "pdf"):
                r = subprocess.run(
                    [soffice, "--headless", "--convert-to", filtro,
                     "--outdir", str(destino.parent), str(f)],
                    capture_output=True, text=True, timeout=300)
                if r.returncode == 0 and destino.exists():
                    break
            if r.returncode == 0 and destino.exists():
                log(f"OK      {f.relative_to(raiz)}")
                ok += 1
            else:
                log(f"ERRO    {f.relative_to(raiz)}  ({r.stderr.strip() or 'sem saída'})")
                falhas += 1
        except Exception as e:
            log(f"ERRO    {f.relative_to(raiz)}  ({e})")
            falhas += 1
        progresso(i, len(ficheiros))
    return ok, falhas


def converter_pasta(pasta, subpastas, log, progresso, motor="auto"):
    pasta = Path(pasta)
    ficheiros = listar_ficheiros(pasta, subpastas)
    if not ficheiros:
        log("Não foram encontrados ficheiros .doc/.docx nesta pasta.")
        return

    log(f"{len(ficheiros)} ficheiro(s) encontrado(s).")

    usar_word = motor in ("auto", "word") and word_disponivel()
    if motor == "word" and not usar_word:
        log("Microsoft Word (pywin32) não está disponível; a tentar LibreOffice.")
    if usar_word:
        log("Motor: Microsoft Word\n")
        ok, falhas = converter_com_word(ficheiros, pasta, log, progresso)
    else:
        soffice = encontrar_libreoffice()
        if not soffice:
            log("Não encontrei o Microsoft Word (pywin32) nem o LibreOffice.\n"
                "Instala o LibreOffice (https://www.libreoffice.org) ou, no "
                "Windows com Word, executa:  pip install pywin32")
            return
        log("Motor: LibreOffice\n")
        ok, falhas = converter_com_libreoffice(soffice, ficheiros, pasta, log, progresso)

    log(f"\nConcluído: {ok} convertido(s), {falhas} com erro.")
    log(f"PDF guardados em: {pasta / PASTA_SAIDA}")


# --------------------------------------------------------------------------
# Interface gráfica
# --------------------------------------------------------------------------
def abrir_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Word → PDF")
    root.geometry("640x420")

    subpastas = tk.BooleanVar(value=False)
    motor_sel = tk.StringVar(value="Automático")
    pasta_sel = tk.StringVar(value="")

    topo = ttk.Frame(root, padding=10)
    topo.pack(fill="x")

    ttk.Entry(topo, textvariable=pasta_sel, state="readonly").pack(
        side="left", fill="x", expand=True, padx=(0, 8))

    caixa = tk.Text(root, height=15, wrap="word", state="disabled")
    barra = ttk.Progressbar(root, mode="determinate")

    def log(msg):
        def _f():
            caixa.config(state="normal")
            caixa.insert("end", msg + "\n")
            caixa.see("end")
            caixa.config(state="disabled")
        root.after(0, _f)

    def progresso(i, total):
        root.after(0, lambda: barra.config(maximum=total, value=i))

    def escolher():
        p = filedialog.askdirectory(title="Escolhe a pasta com os ficheiros Word")
        if p:
            pasta_sel.set(p)
            botao_conv.config(state="normal")

    def converter():
        if not pasta_sel.get():
            return
        botao_conv.config(state="disabled")
        botao_esc.config(state="disabled")
        caixa.config(state="normal")
        caixa.delete("1.0", "end")
        caixa.config(state="disabled")
        barra.config(value=0)

        def tarefa():
            try:
                motor = {"Automático": "auto", "Microsoft Word": "word",
                         "LibreOffice": "libreoffice"}[motor_sel.get()]
                converter_pasta(pasta_sel.get(), subpastas.get(), log, progresso, motor)
            except Exception as e:
                log(f"Erro inesperado: {e}")
            finally:
                root.after(0, lambda: (botao_conv.config(state="normal"),
                                       botao_esc.config(state="normal")))

        threading.Thread(target=tarefa, daemon=True).start()

    botao_esc = ttk.Button(topo, text="Escolher pasta…", command=escolher)
    botao_esc.pack(side="left")

    opcoes = ttk.Frame(root, padding=(10, 0))
    opcoes.pack(fill="x")
    ttk.Checkbutton(opcoes, text="Incluir subpastas", variable=subpastas).pack(side="left")
    ttk.Label(opcoes, text="   Motor:").pack(side="left")
    ttk.Combobox(opcoes, textvariable=motor_sel, state="readonly", width=16,
                 values=["Automático", "Microsoft Word", "LibreOffice"]).pack(side="left")
    botao_conv = ttk.Button(opcoes, text="Converter para PDF", command=converter, state="disabled")
    botao_conv.pack(side="right")

    barra.pack(fill="x", padx=10, pady=8)
    caixa.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    root.mainloop()


# --------------------------------------------------------------------------
def autoteste():
    """Verifica se as dependências ficaram incluídas no executável."""
    erros = []
    try:
        import tkinter  # noqa: F401
    except Exception as e:
        erros.append(f"tkinter: {e}")
    if platform.system() == "Windows":
        try:
            import pythoncom  # noqa: F401
            import win32com.client  # noqa: F401
        except Exception as e:
            erros.append(f"pywin32: {e}")
    return erros


if __name__ == "__main__":
    if "--autoteste" in sys.argv:
        i = sys.argv.index("--autoteste")
        destino = sys.argv[i + 1] if len(sys.argv) > i + 1 else "autoteste.txt"
        erros = autoteste()
        with open(destino, "w", encoding="utf-8") as fh:
            fh.write("OK\n" if not erros else "\n".join(erros) + "\n")
        sys.exit(1 if erros else 0)
    if len(sys.argv) > 1:
        motor = next((a.split("=", 1)[1] for a in sys.argv if a.startswith("--motor=")), "auto")
        converter_pasta(sys.argv[1], "--subpastas" in sys.argv, print,
                        lambda i, t: None, motor)
    else:
        abrir_gui()
