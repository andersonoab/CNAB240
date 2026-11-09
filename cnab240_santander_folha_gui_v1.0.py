# -*- coding: utf-8 -*-
"""
Gerador CNAB 240 Santander - Pagamento de Salarios (Folha) - Interface CustomTkinter
Versao 1.0 - Anderson Souza

- Motor CNAB identico ao gerar_cnab240_santander_folha_audiowave.py (validado byte a byte
  contra o arquivo-modelo Santander/ADP: Header Arquivo 060, Header Lote C-30-01 031,
  Segmentos A/B, Trailers).
- Leitura universal (xlsx/xlsm/xls/xlsb/csv/txt) com cascata de engines.
- Data de pagamento, NSA, valor de teste e dados da empresa definidos na interface.
- Valor por favorecido: coluna "Valor" da planilha; se vazia, usa o "Valor teste".
  Duplo clique na linha abre o detalhe e permite ajustar o valor daquele favorecido.

Dependencias: customtkinter, pandas, openpyxl (xlrd p/ .xls e pyxlsb p/ .xlsb sao opcionais)
"""
import csv
import datetime as dt
import io
import os
import sys
import unicodedata
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import customtkinter as ctk
import pandas as pd

APP_NOME = "Gerador CNAB 240 Santander - Folha"
APP_VERSAO = "1.0"

# ============================================================== PALETA SONOVA OFC
AZUL = "#0083CA"
AZUL_ESC = "#003C64"
AZUL_CLARO = "#6EB4DC"
VERDE = "#2E8B57"
VERDE_HOVER = "#236B43"
VINHO = "#7D0041"
VERMELHO = "#8C321E"
CINZA = "#8A8A8A"
TXT = "#333333"
TXT2 = "#646464"
BORDA = "#CCCCCC"
HEAD_TAB = "#E8F3FA"
ZEBRA = "#F7F9FB"
FONTE_TIT = "Rufina"
FONTE = "Roboto Light"

# ============================================================== CONFIGURACAO PADRAO (chamado 52761)
EMPRESA_PADRAO = {
    "cnpj": "34006534000147",
    "nome": "AUDIO WAVE APARELHOS AUDITIVOS LTDA",
    "agencia": "3689",
    "conta": "13007129",
    "conta_dv": "1",
    "cod_transmissao": "00333689008305861830",
    "logradouro": "RUA ODINEIR DE FREITAS CASTRO",
    "numero": "121",
    "complemento": "LOJA 06",
    "cidade": "VITORIA",
    "cep": "29090440",
    "uf": "ES",
}
EMPRESA = dict(EMPRESA_PADRAO)
BANCO = "033"
NOME_BANCO = "BANCO SANTANDER"
LAYOUT_ARQUIVO = "060"
LAYOUT_LOTE = "031"
SERVICO = "30"   # Pagamento de Salarios
FORMA = "01"     # Credito em conta corrente


# ============================================================== MOTOR CNAB (identico ao validado)
def so_digitos(v):
    return "".join(c for c in str(v if v is not None else "") if c.isdigit())


def alfa(v, n):
    s = unicodedata.normalize("NFKD", str(v or "")).encode("ascii", "ignore").decode("ascii")
    s = " ".join(s.upper().split())
    return s[:n].ljust(n)


def num(v, n):
    s = so_digitos(v)
    if len(s) > n:
        raise ValueError(f"Valor numerico '{v}' excede {n} posicoes")
    return s.zfill(n)


def brancos(n):
    return " " * n


def cpf_valido(c):
    c = so_digitos(c).zfill(11)
    if len(c) != 11 or len(set(c)) == 1:
        return False
    for n in (9, 10):
        s = sum(int(c[i]) * (n + 1 - i) for i in range(n))
        if (s * 10) % 11 % 10 != int(c[n]):
            return False
    return True


def dv_conta_santander(ag, ct):
    base = so_digitos(ag).zfill(4) + "00" + so_digitos(ct).zfill(8)
    pesos = [9, 7, 3, 1, 0, 0, 9, 7, 1, 3, 1, 9, 7, 3]
    t = sum((int(a) * p) % 10 for a, p in zip(base, pesos))
    return str((10 - t % 10) % 10)


def valor_centavos(v):
    if v is None or str(v).strip() == "" or str(v).strip().lower() == "nan":
        return None
    if isinstance(v, (int, float)):
        return int(round(float(v) * 100))
    s = str(v).strip().replace("R$", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    return int(round(float(s) * 100))


def checar(reg, nome):
    if len(reg) != 240:
        raise RuntimeError(f"{nome} com {len(reg)} posicoes (esperado 240)")
    return reg


def header_arquivo(agora, nsa):
    e = EMPRESA
    r = (BANCO + "0000" + "0" + brancos(9) + "2" + num(e["cnpj"], 14) + e["cod_transmissao"]
         + num(e["agencia"], 5) + " " + num(e["conta"], 12) + e["conta_dv"] + " "
         + alfa(e["nome"], 30) + alfa(NOME_BANCO, 30) + brancos(10) + "1"
         + agora.strftime("%d%m%Y") + agora.strftime("%H%M%S") + num(nsa, 6)
         + LAYOUT_ARQUIVO + "00000" + brancos(20) + brancos(20) + brancos(29))
    return checar(r, "Header arquivo")


def header_lote():
    e = EMPRESA
    r = (BANCO + "0001" + "1" + "C" + SERVICO + FORMA + LAYOUT_LOTE + " " + "2"
         + num(e["cnpj"], 14) + e["cod_transmissao"] + num(e["agencia"], 5) + " "
         + num(e["conta"], 12) + e["conta_dv"] + " " + alfa(e["nome"], 30) + brancos(40)
         + alfa(e["logradouro"], 30) + num(e["numero"], 5) + alfa(e["complemento"], 15)
         + alfa(e["cidade"], 20) + num(e["cep"], 8) + alfa(e["uf"], 2) + brancos(8) + brancos(10))
    return checar(r, "Header lote")


def segmento_a(seq, f, data, cent):
    r = (BANCO + "0001" + "3" + num(seq, 5) + "A" + "0" + "00" + "000" + BANCO
         + num(f["agencia"], 5) + " " + num(f["conta"], 12) + f["dv"] + " "
         + alfa(f["nome"], 30) + num(f["cpf"], 20) + data + "BRL" + num(0, 15)
         + num(cent, 15) + brancos(20) + data + num(cent, 15) + brancos(40)
         + "01" + brancos(5) + brancos(2) + brancos(3) + "0" + brancos(10))
    return checar(r, "Segmento A")


def segmento_b(seq, f, data, cent):
    r = (BANCO + "0001" + "3" + num(seq, 5) + "B" + brancos(3) + "1" + num(f["cpf"], 14)
         + brancos(30) + num(0, 5) + brancos(15) + brancos(15) + brancos(20) + num(0, 8)
         + brancos(2) + data + num(cent, 15) + num(0, 15) * 4 + "0000" + brancos(11)
         + "0" + "2007  " + brancos(8))
    return checar(r, "Segmento B")


def trailer_lote(qtd_reg, total):
    r = (BANCO + "0001" + "5" + brancos(9) + num(qtd_reg, 6) + num(total, 18)
         + num(0, 18) + brancos(6) + brancos(165) + brancos(10))
    return checar(r, "Trailer lote")


def trailer_arquivo(qtd_reg):
    r = BANCO + "9999" + "9" + brancos(9) + num(1, 6) + num(qtd_reg, 6) + brancos(6) + brancos(205)
    return checar(r, "Trailer arquivo")


def validar_favorecido(f, cent):
    erros = []
    if not cpf_valido(f["cpf"]):
        erros.append("CPF invalido")
    if f["banco"] != BANCO:
        erros.append(f"Banco {f['banco']} <> 033")
    if not f["agencia"] or not f["conta"]:
        erros.append("Agencia/conta ausente")
    if len(f["dv"]) != 1:
        erros.append("DV da conta ausente/invalido")
    elif f["agencia"] and f["conta"] and dv_conta_santander(f["agencia"], f["conta"]) != f["dv"]:
        erros.append(f"DV conta nao confere p/ ag {f['agencia']} (calculado "
                     f"{dv_conta_santander(f['agencia'], f['conta'])}, informado {f['dv']})")
    if cent is None or cent <= 0:
        erros.append("Valor ausente ou <= 0")
    return erros


def montar_arquivo(validos, data_ddmmaaaa, nsa, agora=None):
    """validos: lista de (favorecido, centavos). Retorna (registros, total_centavos)."""
    agora = agora or dt.datetime.now()
    regs = [header_arquivo(agora, nsa), header_lote()]
    seq, total = 0, 0
    for f, cent in validos:
        seq += 1
        regs.append(segmento_a(seq, f, data_ddmmaaaa, cent))
        seq += 1
        regs.append(segmento_b(seq, f, data_ddmmaaaa, cent))
        total += cent
    regs.append(trailer_lote(seq + 2, total))
    regs.append(trailer_arquivo(len(regs) + 1))
    return regs, total


# ============================================================== LEITURA UNIVERSAL
ALIASES = {
    "cpf": ("CPF", "CPF FAVORECIDO", "DOCUMENTO"),
    "nome": ("NOME", "NOME FAVORECIDO", "FAVORECIDO", "COLABORADOR", "FUNCIONARIO"),
    "banco": ("BANCO", "COD BANCO", "CODIGO BANCO"),
    "agencia": ("AGENCIA", "AG"),
    "conta": ("CONTA", "CONTA CORRENTE", "C/C"),
    "dv": ("DIGITO", "DV", "DIG", "DIGITO CONTA", "DV CONTA"),
    "valor": ("VALOR", "LIQUIDO", "VALOR LIQUIDO", "VLR LIQUIDO", "VALOR A PAGAR"),
}
OBRIGATORIAS = ("cpf", "nome", "banco", "agencia", "conta", "dv")


def norm_cab(v):
    return alfa(v, 60).strip()


def listar_abas(caminho):
    ext = os.path.splitext(caminho)[1].lower()
    if ext in (".csv", ".txt"):
        return ["(arquivo texto)"]
    engines = {".xlsx": ["openpyxl"], ".xlsm": ["openpyxl"], ".xls": ["xlrd", "openpyxl"],
               ".xlsb": ["pyxlsb"]}.get(ext, ["openpyxl", "xlrd", "pyxlsb"])
    ultimo = None
    for eng in engines:
        try:
            return list(pd.ExcelFile(caminho, engine=eng).sheet_names)
        except Exception as e:  # tenta a proxima engine
            ultimo = e
    raise RuntimeError(f"Nao foi possivel abrir o arquivo ({ultimo})")


def ler_tabela(caminho, aba):
    ext = os.path.splitext(caminho)[1].lower()
    if ext in (".csv", ".txt"):
        bruto = open(caminho, "rb").read()
        for enc in ("utf-8-sig", "latin1", "cp1252"):
            try:
                texto = bruto.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        try:
            sep = csv.Sniffer().sniff(texto[:5000], delimiters=";,\t|").delimiter
        except csv.Error:
            sep = ";"
        return pd.read_csv(io.StringIO(texto), sep=sep, dtype=str, keep_default_na=False)
    engines = {".xlsx": ["openpyxl"], ".xlsm": ["openpyxl"], ".xls": ["xlrd", "openpyxl"],
               ".xlsb": ["pyxlsb"]}.get(ext, ["openpyxl", "xlrd", "pyxlsb"])
    ultimo = None
    for eng in engines:
        try:
            return pd.read_excel(caminho, sheet_name=aba, dtype=str, engine=eng, keep_default_na=False)
        except Exception as e:
            ultimo = e
    raise RuntimeError(f"Nao foi possivel ler a aba '{aba}' ({ultimo})")


def mapear_colunas(df):
    cab = {norm_cab(c): c for c in df.columns}
    mapa = {}
    for campo, nomes in ALIASES.items():
        for n in nomes:
            if n in cab:
                mapa[campo] = cab[n]
                break
    faltam = [c for c in OBRIGATORIAS if c not in mapa]
    return mapa, faltam


def extrair_favorecidos(df, mapa):
    out = []
    for i, r in df.iterrows():
        vals = [str(r[c]).strip() for c in df.columns]
        if not any(v and v.lower() != "nan" for v in vals):
            continue
        ag_raw = so_digitos(r[mapa["agencia"]])
        out.append({
            "linha": int(i) + 2,
            "cpf": so_digitos(r[mapa["cpf"]]).zfill(11),
            "nome": str(r[mapa["nome"]]).strip(),
            "banco": so_digitos(r[mapa["banco"]]).zfill(3),
            "agencia": ag_raw,
            "conta": so_digitos(r[mapa["conta"]]),
            "dv": so_digitos(r[mapa["dv"]]),
            "valor_planilha": (str(r[mapa["valor"]]).strip() if "valor" in mapa else ""),
            "valor_manual": None,
        })
    return out


def fmt_brl(cent):
    if cent is None:
        return "-"
    return f"R$ {cent/100:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def proximo_dia_util(d):
    d = d + dt.timedelta(days=1)
    while d.weekday() >= 5:
        d += dt.timedelta(days=1)
    return d


# ============================================================== INTERFACE
class App(ctk.CTk):
    COLS = [("st", "Status", 80, False), ("linha", "Linha", 55, False), ("cpf", "CPF", 110, False),
            ("nome", "Nome", 260, True), ("ag", "Agência", 70, False), ("conta", "Conta", 95, False),
            ("dv", "DV", 40, False), ("valor", "Valor", 110, False), ("obs", "Ocorrência", 300, True)]

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        self.title(f"{APP_NOME} v{APP_VERSAO}")
        self.geometry("1280x800")
        self.minsize(1000, 650)
        self.configure(fg_color="white")

        self.caminho = None
        self.favorecidos = []
        self.ordem = {"col": None, "desc": False}
        self.ultimo_arquivo = None

        self._estilos()
        self._header()
        self.tabs = ctk.CTkTabview(self, fg_color="white", segmented_button_selected_color=AZUL,
                                   segmented_button_selected_hover_color=AZUL_ESC,
                                   segmented_button_unselected_color=AZUL_CLARO, text_color="white")
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(6, 0))
        self.t_ger = self.tabs.add("Gerar remessa")
        self.t_emp = self.tabs.add("Empresa / Convênio")
        self.t_prev = self.tabs.add("Prévia do arquivo")
        self._aba_gerar()
        self._aba_empresa()
        self._aba_previa()
        self._footer()

    # ---------------------------------------------------------- estilos ttk
    def _estilos(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure("Ofc.Treeview", background="white", fieldbackground="white", foreground=TXT,
                     rowheight=24, font=(FONTE, 9), bordercolor=BORDA, borderwidth=1)
        st.configure("Ofc.Treeview.Heading", background=HEAD_TAB, foreground=AZUL_ESC,
                     font=(FONTE, 9, "bold"), relief="flat", bordercolor=BORDA)
        st.map("Ofc.Treeview.Heading", background=[("active", HEAD_TAB)])
        st.map("Ofc.Treeview", background=[("selected", AZUL)], foreground=[("selected", "white")])
        for o in ("Vertical", "Horizontal"):
            st.layout(f"Ofc.{o}.TScrollbar", [(f"{o}.Scrollbar.trough", {"sticky": "nswe", "children": [
                (f"{o}.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
            st.configure(f"Ofc.{o}.TScrollbar", troughcolor="#EEF3F7", background="#B9CBD8",
                         bordercolor="#EEF3F7", lightcolor="#B9CBD8", darkcolor="#B9CBD8",
                         arrowsize=11, width=11, relief="flat")
            st.map(f"Ofc.{o}.TScrollbar", background=[("active", AZUL), ("pressed", AZUL)])

    def _scroll_auto(self, sb):
        def setter(lo, hi):
            if float(lo) <= 0.0 and float(hi) >= 1.0:
                sb.grid_remove()
            else:
                sb.grid()
            sb.set(lo, hi)
        return setter

    # ---------------------------------------------------------- header / footer
    def _header(self):
        h = ctk.CTkFrame(self, fg_color=AZUL, corner_radius=0, height=58)
        h.pack(fill="x")
        h.pack_propagate(False)
        ctk.CTkLabel(h, text="CNAB 240 Santander · Pagamento de Salários",
                     font=(FONTE_TIT, 20, "bold"), text_color="white").pack(side="left", padx=18)
        ctk.CTkLabel(h, text=f"v{APP_VERSAO}  |  Sonova", font=(FONTE, 12),
                     text_color="white").pack(side="right", padx=18)

    def _footer(self):
        f = ctk.CTkFrame(self, fg_color=AZUL, corner_radius=0, height=28)
        f.pack(fill="x", side="bottom")
        f.pack_propagate(False)
        self.lbl_status = ctk.CTkLabel(f, text="Pronto. Selecione a planilha de contas.",
                                       font=(FONTE, 11), text_color="white")
        self.lbl_status.pack(side="left", padx=12)
        ctk.CTkLabel(f, text="Anderson Souza", font=(FONTE, 11), text_color="white").pack(side="right", padx=12)

    def status(self, msg):
        self.lbl_status.configure(text=msg)
        self.update_idletasks()

    def _btn(self, pai, texto, cmd, tipo="primario", w=150):
        cores = {"avancar": (VERDE, VERDE_HOVER), "primario": (AZUL, AZUL_ESC),
                 "secundario": (AZUL_CLARO, AZUL), "aux": (CINZA, "#6E6E6E"),
                 "destrutivo": (VERMELHO, "#6E2717")}[tipo]
        return ctk.CTkButton(pai, text=texto, command=cmd, width=w, height=32, corner_radius=6,
                             fg_color=cores[0], hover_color=cores[1], text_color="white",
                             font=(FONTE, 12, "bold"))

    # ---------------------------------------------------------- aba gerar
    def _aba_gerar(self):
        p = self.t_ger
        # linha 1 - arquivo
        l1 = ctk.CTkFrame(p, fg_color="white", border_color=BORDA, border_width=1)
        l1.pack(fill="x", pady=(4, 6))
        self._btn(l1, "Selecionar planilha", self.escolher_arquivo, "primario", 170).grid(
            row=0, column=0, padx=10, pady=10)
        self.lbl_arq = ctk.CTkLabel(l1, text="Nenhum arquivo selecionado", text_color=TXT2,
                                    font=(FONTE, 12), anchor="w")
        self.lbl_arq.grid(row=0, column=1, sticky="w", padx=6)
        ctk.CTkLabel(l1, text="Aba:", font=(FONTE, 12), text_color=TXT).grid(row=0, column=2, padx=(20, 4))
        self.cb_aba = ctk.CTkComboBox(l1, values=["-"], width=200, command=lambda _: self.carregar(),
                                      border_color=BORDA, button_color=AZUL)
        self.cb_aba.grid(row=0, column=3, padx=4)
        self._btn(l1, "Recarregar", self.carregar, "secundario", 110).grid(row=0, column=4, padx=10)
        l1.grid_columnconfigure(1, weight=1)

        # linha 2 - parametros
        l2 = ctk.CTkFrame(p, fg_color="white", border_color=BORDA, border_width=1)
        l2.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(l2, text="Data de pagamento (DD/MM/AAAA):", font=(FONTE, 12), text_color=TXT).grid(
            row=0, column=0, padx=(10, 4), pady=10)
        self.ent_data = ctk.CTkEntry(l2, width=110, border_color=BORDA)
        self.ent_data.insert(0, proximo_dia_util(dt.date.today()).strftime("%d/%m/%Y"))
        self.ent_data.grid(row=0, column=1, padx=4)
        self._btn(l2, "Hoje", lambda: self._set_data(dt.date.today()), "aux", 60).grid(row=0, column=2, padx=2)
        self._btn(l2, "+1 dia útil", self._data_mais_um, "aux", 90).grid(row=0, column=3, padx=2)

        ctk.CTkLabel(l2, text="NSA:", font=(FONTE, 12), text_color=TXT).grid(row=0, column=4, padx=(20, 4))
        self.ent_nsa = ctk.CTkEntry(l2, width=70, border_color=BORDA)
        self.ent_nsa.insert(0, "1")
        self.ent_nsa.grid(row=0, column=5)

        ctk.CTkLabel(l2, text="Valor teste (se planilha sem valor):", font=(FONTE, 12),
                     text_color=TXT).grid(row=0, column=6, padx=(20, 4))
        self.ent_vteste = ctk.CTkEntry(l2, width=90, border_color=BORDA)
        self.ent_vteste.insert(0, "0,01")
        self.ent_vteste.grid(row=0, column=7)
        self.var_forcar = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(l2, text="Usar valor teste p/ todos", variable=self.var_forcar,
                        command=self.revalidar, fg_color=AZUL, font=(FONTE, 12),
                        text_color=TXT).grid(row=0, column=8, padx=12)
        self.ent_vteste.bind("<FocusOut>", lambda e: self.revalidar())

        # KPIs
        k = ctk.CTkFrame(p, fg_color="white")
        k.pack(fill="x", pady=(0, 6))
        self.kpis = {}
        for i, (key, tit, cor) in enumerate([("tot", "Linhas lidas", AZUL_ESC), ("ok", "Aptos", VERDE),
                                              ("bloq", "Bloqueados", VINHO), ("vlr", "Total a pagar", AZUL)]):
            c = ctk.CTkFrame(k, fg_color="white", border_color=BORDA, border_width=1, corner_radius=8)
            c.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 6, 0))
            k.grid_columnconfigure(i, weight=1)
            ctk.CTkLabel(c, text=tit, font=(FONTE, 11), text_color=TXT2).pack(anchor="w", padx=12, pady=(8, 0))
            lb = ctk.CTkLabel(c, text="-", font=(FONTE_TIT, 22, "bold"), text_color=cor)
            lb.pack(anchor="w", padx=12, pady=(0, 8))
            self.kpis[key] = lb

        # tabela
        moldura = tk.Frame(p, bg="white", highlightbackground=BORDA, highlightthickness=1)
        moldura.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(moldura, columns=[c[0] for c in self.COLS], show="headings",
                                 style="Ofc.Treeview", height=10)
        for cid, tit, w, elast in self.COLS:
            self.tree.heading(cid, text=tit, command=lambda c=cid: self.ordenar(c))
            self.tree.column(cid, width=w, minwidth=40, stretch=elast,
                             anchor="w" if elast or cid == "cpf" else "center")
        sy = ttk.Scrollbar(moldura, orient="vertical", command=self.tree.yview, style="Ofc.Vertical.TScrollbar")
        sx = ttk.Scrollbar(moldura, orient="horizontal", command=self.tree.xview, style="Ofc.Horizontal.TScrollbar")
        self.tree.configure(yscrollcommand=self._scroll_auto(sy), xscrollcommand=self._scroll_auto(sx))
        self.tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        moldura.grid_rowconfigure(0, weight=1)
        moldura.grid_columnconfigure(0, weight=1)
        self.tree.tag_configure("par", background="white")
        self.tree.tag_configure("impar", background=ZEBRA)
        self.tree.tag_configure("bloq", foreground=VINHO)
        self.tree.bind("<Double-1>", self.detalhe)

        # acoes
        a = ctk.CTkFrame(p, fg_color="white")
        a.pack(fill="x", pady=8)
        self._btn(a, "Gerar arquivo .REM", self.gerar, "avancar", 180).pack(side="right")
        self._btn(a, "Exportar ocorrências (CSV)", self.exportar_ocorrencias, "secundario", 200).pack(
            side="right", padx=8)
        self._btn(a, "Abrir pasta do arquivo", self.abrir_pasta, "aux", 170).pack(side="right")
        self._btn(a, "Limpar", self.limpar, "destrutivo", 90).pack(side="left")

    def _set_data(self, d):
        self.ent_data.delete(0, "end")
        self.ent_data.insert(0, d.strftime("%d/%m/%Y"))

    def _data_mais_um(self):
        try:
            d = dt.datetime.strptime(self.ent_data.get().strip(), "%d/%m/%Y").date()
        except ValueError:
            d = dt.date.today()
        self._set_data(proximo_dia_util(d))

    # ---------------------------------------------------------- aba empresa
    def _aba_empresa(self):
        p = ctk.CTkScrollableFrame(self.t_emp, fg_color="white")
        p.pack(fill="both", expand=True)
        campos = [("cnpj", "CNPJ (14 dígitos)"), ("nome", "Razão social"), ("agencia", "Agência"),
                  ("conta", "Conta (sem dígito)"), ("conta_dv", "Dígito da conta"),
                  ("cod_transmissao", "Código de transmissão (20 posições)"), ("logradouro", "Logradouro"),
                  ("numero", "Número"), ("complemento", "Complemento"), ("cidade", "Cidade"),
                  ("cep", "CEP (8 dígitos)"), ("uf", "UF")]
        self.ent_emp = {}
        for i, (k, t) in enumerate(campos):
            ctk.CTkLabel(p, text=t, font=(FONTE, 12), text_color=TXT, anchor="w").grid(
                row=i, column=0, sticky="w", padx=14, pady=5)
            e = ctk.CTkEntry(p, width=380, border_color=BORDA)
            e.insert(0, EMPRESA[k])
            e.grid(row=i, column=1, sticky="w", padx=6, pady=5)
            self.ent_emp[k] = e
        b = ctk.CTkFrame(p, fg_color="white")
        b.grid(row=len(campos), column=0, columnspan=2, sticky="w", padx=14, pady=12)
        self._btn(b, "Aplicar", self.aplicar_empresa, "primario", 120).pack(side="left")
        self._btn(b, "Restaurar padrão (chamado 52761)", self.restaurar_empresa, "aux", 260).pack(
            side="left", padx=8)
        ctk.CTkLabel(p, text="Layout fixo: banco 033 · arquivo 060 · lote 031 · serviço 30 (Salários) · "
                             "forma 01 (crédito em conta corrente).", font=(FONTE, 11),
                     text_color=TXT2).grid(row=len(campos) + 1, column=0, columnspan=2, sticky="w", padx=14)

    def aplicar_empresa(self, silencioso=False):
        novo = {k: e.get().strip() for k, e in self.ent_emp.items()}
        erros = []
        if len(so_digitos(novo["cnpj"])) != 14:
            erros.append("CNPJ deve ter 14 dígitos")
        if len(so_digitos(novo["cod_transmissao"])) != 20:
            erros.append("Código de transmissão deve ter 20 dígitos")
        if len(so_digitos(novo["conta_dv"])) != 1:
            erros.append("Dígito da conta deve ter 1 posição")
        if len(so_digitos(novo["cep"])) != 8:
            erros.append("CEP deve ter 8 dígitos")
        if not so_digitos(novo["agencia"]) or not so_digitos(novo["conta"]):
            erros.append("Agência e conta são obrigatórias")
        elif dv_conta_santander(novo["agencia"], novo["conta"]) != so_digitos(novo["conta_dv"]):
            erros.append(f"Dígito da conta da empresa não confere (calculado "
                         f"{dv_conta_santander(novo['agencia'], novo['conta'])})")
        if erros:
            messagebox.showerror("Empresa", "\n".join(erros))
            return False
        novo["cnpj"] = so_digitos(novo["cnpj"])
        novo["cod_transmissao"] = so_digitos(novo["cod_transmissao"])
        novo["conta_dv"] = so_digitos(novo["conta_dv"])
        novo["cep"] = so_digitos(novo["cep"])
        EMPRESA.update(novo)
        if not silencioso:
            self.status("Dados da empresa aplicados.")
        return True

    def restaurar_empresa(self):
        for k, e in self.ent_emp.items():
            e.delete(0, "end")
            e.insert(0, EMPRESA_PADRAO[k])
        self.aplicar_empresa()

    # ---------------------------------------------------------- aba previa
    def _aba_previa(self):
        p = self.t_prev
        ctk.CTkLabel(p, text="Último arquivo gerado (régua de 240 posições)", font=(FONTE_TIT, 14, "bold"),
                     text_color=AZUL_ESC).pack(anchor="w", pady=(4, 4))
        fr = tk.Frame(p, bg="white", highlightbackground=BORDA, highlightthickness=1)
        fr.pack(fill="both", expand=True)
        self.txt = tk.Text(fr, wrap="none", font=("Consolas", 9), fg=TXT, bg="white", relief="flat")
        sy = ttk.Scrollbar(fr, orient="vertical", command=self.txt.yview, style="Ofc.Vertical.TScrollbar")
        sx = ttk.Scrollbar(fr, orient="horizontal", command=self.txt.xview, style="Ofc.Horizontal.TScrollbar")
        self.txt.configure(yscrollcommand=self._scroll_auto(sy), xscrollcommand=self._scroll_auto(sx))
        self.txt.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        fr.grid_rowconfigure(0, weight=1)
        fr.grid_columnconfigure(0, weight=1)

    # ---------------------------------------------------------- carga / validacao
    def escolher_arquivo(self):
        c = filedialog.askopenfilename(title="Planilha de contas", filetypes=[
            ("Planilhas e textos", "*.xlsx *.xlsm *.xls *.xlsb *.csv *.txt"), ("Todos", "*.*")])
        if not c:
            return
        self.caminho = c
        self.lbl_arq.configure(text=os.path.basename(c), text_color=TXT)
        try:
            abas = listar_abas(c)
        except Exception as e:
            messagebox.showerror("Leitura", str(e))
            return
        self.cb_aba.configure(values=abas)
        self.cb_aba.set(abas[0])
        self.carregar()

    def carregar(self):
        if not self.caminho:
            return
        try:
            aba = self.cb_aba.get()
            df = ler_tabela(self.caminho, None if aba == "(arquivo texto)" else aba)
            mapa, faltam = mapear_colunas(df)
            if faltam:
                messagebox.showerror("Colunas", "Colunas obrigatórias não encontradas: " + ", ".join(faltam)
                                     + "\n\nCabeçalhos lidos: " + ", ".join(map(str, df.columns)))
                return
            self.favorecidos = extrair_favorecidos(df, mapa)
            extra = "" if "valor" in mapa else " (sem coluna Valor: será usado o valor teste)"
            self.status(f"{len(self.favorecidos)} linhas lidas da aba '{aba}'{extra}.")
            self.revalidar()
        except Exception as e:
            messagebox.showerror("Leitura", f"Falha ao ler a planilha:\n{e}")

    def _cent_teste(self):
        try:
            return valor_centavos(self.ent_vteste.get())
        except ValueError:
            return None

    def _cent(self, f):
        if f["valor_manual"] is not None:
            return f["valor_manual"]
        if not self.var_forcar.get():
            try:
                c = valor_centavos(f["valor_planilha"])
            except ValueError:
                c = None
            if c is not None:
                return c
        return self._cent_teste()

    def revalidar(self):
        for f in self.favorecidos:
            f["cent"] = self._cent(f)
            f["erros"] = validar_favorecido(f, f["cent"])
        self.render()

    def render(self):
        self.tree.delete(*self.tree.get_children())
        lista = list(self.favorecidos)
        col = self.ordem["col"]
        if col:
            chave = {"st": lambda f: bool(f["erros"]), "linha": lambda f: f["linha"],
                     "cpf": lambda f: f["cpf"], "nome": lambda f: f["nome"],
                     "ag": lambda f: f["agencia"].zfill(5), "conta": lambda f: f["conta"].zfill(12),
                     "dv": lambda f: f["dv"], "valor": lambda f: f["cent"] or 0,
                     "obs": lambda f: "; ".join(f["erros"])}[col]
            lista.sort(key=chave, reverse=self.ordem["desc"])
        for i, f in enumerate(lista):
            tags = ["impar" if i % 2 else "par"] + (["bloq"] if f["erros"] else [])
            self.tree.insert("", "end", iid=str(f["linha"]), tags=tags, values=(
                "BLOQUEADO" if f["erros"] else "OK", f["linha"], f["cpf"], f["nome"], f["agencia"],
                f["conta"], f["dv"], fmt_brl(f["cent"]) + (" *" if f["valor_manual"] is not None else ""),
                "; ".join(f["erros"])))
        for cid, tit, _, _ in self.COLS:
            seta = (" v" if self.ordem["desc"] else " ^") if cid == col else ""
            self.tree.heading(cid, text=tit + seta)
        ok = [f for f in self.favorecidos if not f["erros"]]
        self.kpis["tot"].configure(text=str(len(self.favorecidos)))
        self.kpis["ok"].configure(text=str(len(ok)))
        self.kpis["bloq"].configure(text=str(len(self.favorecidos) - len(ok)))
        self.kpis["vlr"].configure(text=fmt_brl(sum(f["cent"] for f in ok)))

    def ordenar(self, col):
        if self.ordem["col"] == col:
            self.ordem["desc"] = not self.ordem["desc"]
        else:
            self.ordem = {"col": col, "desc": False}
        self.render()

    # ---------------------------------------------------------- detalhe
    def detalhe(self, _evt=None):
        sel = self.tree.focus()
        if not sel:
            return
        f = next((x for x in self.favorecidos if str(x["linha"]) == sel), None)
        if not f:
            return
        w = ctk.CTkToplevel(self)
        w.title(f"Linha {f['linha']} - {f['nome']}")
        w.geometry("560x460")
        w.configure(fg_color="white")
        w.transient(self)
        w.grab_set()
        h = ctk.CTkFrame(w, fg_color=AZUL, corner_radius=0, height=40)
        h.pack(fill="x")
        ctk.CTkLabel(h, text=f["nome"], font=(FONTE_TIT, 14, "bold"), text_color="white").pack(
            side="left", padx=12, pady=8)
        corpo = ctk.CTkFrame(w, fg_color="white")
        corpo.pack(fill="both", expand=True, padx=14, pady=10)
        calc = dv_conta_santander(f["agencia"], f["conta"]) if f["agencia"] and f["conta"] else "-"
        dados = [("Linha na planilha", f["linha"]), ("CPF", f["cpf"]), ("Banco", f["banco"]),
                 ("Agência", f["agencia"]), ("Conta", f["conta"]), ("DV informado", f["dv"]),
                 ("DV calculado (Santander)", calc), ("Valor na planilha", f["valor_planilha"] or "(vazio)"),
                 ("Valor usado", fmt_brl(f["cent"])),
                 ("Ocorrências", "; ".join(f["erros"]) or "Nenhuma")]
        for i, (k, v) in enumerate(dados):
            ctk.CTkLabel(corpo, text=k, font=(FONTE, 12), text_color=TXT2, anchor="w").grid(
                row=i, column=0, sticky="w", pady=2)
            ctk.CTkLabel(corpo, text=str(v), font=(FONTE, 12, "bold"), text_color=TXT, anchor="w",
                         wraplength=320, justify="left").grid(row=i, column=1, sticky="w", padx=10, pady=2)
        n = len(dados)
        ctk.CTkLabel(corpo, text="Valor manual (teste):", font=(FONTE, 12), text_color=TXT).grid(
            row=n, column=0, sticky="w", pady=(12, 2))
        ent = ctk.CTkEntry(corpo, width=120, border_color=BORDA)
        if f["valor_manual"] is not None:
            ent.insert(0, f"{f['valor_manual']/100:.2f}".replace(".", ","))
        ent.grid(row=n, column=1, sticky="w", padx=10, pady=(12, 2))

        def salvar():
            t = ent.get().strip()
            if t == "":
                f["valor_manual"] = None
            else:
                try:
                    c = valor_centavos(t)
                except ValueError:
                    messagebox.showerror("Valor", "Valor inválido.", parent=w)
                    return
                if c is None or c <= 0:
                    messagebox.showerror("Valor", "Informe um valor maior que zero.", parent=w)
                    return
                f["valor_manual"] = c
            w.destroy()
            self.revalidar()

        b = ctk.CTkFrame(w, fg_color="white")
        b.pack(fill="x", padx=14, pady=10)
        self._btn(b, "Salvar valor", salvar, "primario", 130).pack(side="right")
        self._btn(b, "Fechar", w.destroy, "aux", 90).pack(side="right", padx=8)

    # ---------------------------------------------------------- gerar
    def gerar(self):
        if not self.favorecidos:
            messagebox.showwarning("Gerar", "Carregue a planilha primeiro.")
            return
        if not self.aplicar_empresa(silencioso=True):
            self.tabs.set("Empresa / Convênio")
            return
        try:
            d = dt.datetime.strptime(self.ent_data.get().strip(), "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Data", "Data de pagamento inválida. Use DD/MM/AAAA.")
            return
        if d.date() < dt.date.today():
            if not messagebox.askyesno("Data", "A data de pagamento é anterior a hoje. Continuar?"):
                return
        if d.weekday() >= 5:
            if not messagebox.askyesno("Data", "A data de pagamento cai em fim de semana. Continuar?"):
                return
        try:
            nsa = int(so_digitos(self.ent_nsa.get()) or "0")
        except ValueError:
            nsa = 0
        if not 1 <= nsa <= 999999:
            messagebox.showerror("NSA", "NSA deve ser entre 1 e 999999.")
            return
        self.revalidar()
        validos = [(f, f["cent"]) for f in self.favorecidos if not f["erros"]]
        bloq = [f for f in self.favorecidos if f["erros"]]
        if not validos:
            messagebox.showerror("Gerar", "Nenhum favorecido apto.")
            return
        if bloq and not messagebox.askyesno(
                "Bloqueados", f"{len(bloq)} favorecido(s) bloqueado(s) ficarão FORA do arquivo.\nContinuar?"):
            return
        ddmm = d.strftime("%d%m%Y")
        nome = f"CNAB240_SANTANDER_FOLHA_{alfa(EMPRESA['nome'], 30).strip().replace(' ', '_')[:20]}" \
               f"_{ddmm}_NSA{nsa:06d}.REM"
        destino = filedialog.asksaveasfilename(title="Salvar remessa", initialfile=nome,
                                               defaultextension=".REM",
                                               filetypes=[("Remessa CNAB", "*.REM *.txt"), ("Todos", "*.*")])
        if not destino:
            return
        try:
            regs, total = montar_arquivo(validos, ddmm, nsa)
            conteudo = "".join(r + "\r\n" for r in regs)
            if os.path.exists(destino):
                base, ext = os.path.splitext(destino)
                destino = f"{base}_{dt.datetime.now():%H%M%S}{ext}"
            with open(destino, "w", encoding="ascii", newline="") as fh:
                fh.write(conteudo)
            # conferencia pos-gravacao
            lido = open(destino, "rb").read().split(b"\r\n")
            lido = [x for x in lido if x]
            assert len(lido) == len(regs) and all(len(x) == 240 for x in lido)
        except Exception as e:
            messagebox.showerror("Gerar", f"Falha ao gerar/gravar o arquivo:\n{e}")
            return
        self.ultimo_arquivo = destino
        self._mostrar_previa(regs)
        self.status(f"Gerado: {os.path.basename(destino)} | {len(validos)} favorecidos | "
                    f"{len(regs)} registros | {fmt_brl(total)}")
        messagebox.showinfo("Remessa gerada",
                            f"Arquivo: {destino}\n\nFavorecidos: {len(validos)}\nRegistros: {len(regs)}\n"
                            f"Total: {fmt_brl(total)}\nBloqueados (fora do arquivo): {len(bloq)}\n\n"
                            "Conferência: todos os registros com 240 posições e CRLF.")

    def _mostrar_previa(self, regs):
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        regua = "".join(str((i // 10) % 10) if i % 10 == 0 else " " for i in range(1, 241))
        regua2 = "".join(str(i % 10) for i in range(1, 241))
        self.txt.insert("end", regua + "\n" + regua2 + "\n" + "-" * 240 + "\n")
        for r in regs:
            self.txt.insert("end", r + "\n")
        self.txt.configure(state="disabled")

    # ---------------------------------------------------------- utilitarios
    def exportar_ocorrencias(self):
        if not self.favorecidos:
            return
        c = filedialog.asksaveasfilename(title="Exportar ocorrências", defaultextension=".csv",
                                         initialfile="ocorrencias_cnab.csv", filetypes=[("CSV", "*.csv")])
        if not c:
            return
        with open(c, "w", encoding="utf-8-sig", newline="") as fh:
            w = csv.writer(fh, delimiter=";")
            w.writerow(["Status", "Linha", "CPF", "Nome", "Banco", "Agencia", "Conta", "DV",
                        "DV calculado", "Valor", "Ocorrencia"])
            for f in self.favorecidos:
                calc = dv_conta_santander(f["agencia"], f["conta"]) if f["agencia"] and f["conta"] else ""
                w.writerow(["BLOQUEADO" if f["erros"] else "OK", f["linha"], f["cpf"], f["nome"], f["banco"],
                            f["agencia"], f["conta"], f["dv"], calc,
                            f"{(f['cent'] or 0)/100:.2f}".replace(".", ","), "; ".join(f["erros"])])
        self.status(f"Ocorrências exportadas: {os.path.basename(c)}")

    def abrir_pasta(self):
        if self.ultimo_arquivo and os.path.exists(self.ultimo_arquivo):
            os.startfile(os.path.dirname(self.ultimo_arquivo)) if sys.platform == "win32" else None
        else:
            messagebox.showinfo("Pasta", "Nenhum arquivo gerado nesta sessão.")

    def limpar(self):
        self.caminho = None
        self.favorecidos = []
        self.lbl_arq.configure(text="Nenhum arquivo selecionado", text_color=TXT2)
        self.cb_aba.configure(values=["-"])
        self.cb_aba.set("-")
        self.render()
        for k in self.kpis.values():
            k.configure(text="-")
        self.status("Pronto. Selecione a planilha de contas.")


if __name__ == "__main__":
    App().mainloop()
