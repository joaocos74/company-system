# main.py (com suporte a logo embutida em base64)

import os
import re
import datetime
import unicodedata
import base64
from io import BytesIO

import customtkinter as ctk
from tkinter import *
from tkinter import messagebox
from PIL import Image  # pip install pillow
import openpyxl  # pip install openpyxl

# Telas existentes
from cadastro import CadastroApp
from pesquisar import PesquisarApp
from Atualizar import AtualizarApp

try:
    from estatistica import EstatisticaApp
except Exception:
    EstatisticaApp = None

# ========== NOVO: Importar UsuarioApp ==========
try:
    from usuario import UsuarioApp
except Exception:
    UsuarioApp = None

# Importa a logo embutida (opcional)
try:
    from logo_b64 import LOGO_B64
except Exception:
    LOGO_B64 = None

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("green")

# -------- utilidades p/ ler a planilha --------
def _norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s

def _extrair_ano(valor):
    if isinstance(valor, (datetime.datetime, datetime.date)):
        return valor.year
    if valor is None:
        return None
    m = re.search(r"(19|20)\d{2}", str(valor))
    return int(m.group(0)) if m else None

def _col_index(sheet, header_name):
    target = _norm(header_name)
    headers = [c.value for c in sheet[1]]
    for idx, h in enumerate(headers):
        if _norm(h) == target:
            return idx
    if "inspe" in target:
        for idx, h in enumerate(headers):
            if "inspe" in _norm(h):
                return idx
    for idx, h in enumerate(headers):
        if _norm(h).startswith(target):
            return idx
    raise ValueError(f"Coluna '{header_name}' não encontrada. Cabeçalhos: {headers}")

def calcular_kpis_home(caminho="cadastros.xlsx"):
    """
    Retorna: total_cadastrados, ano_mais_recente, inspecionados_no_ano, cobertura_percent
    (Cobertura = inspecionados_no_ano / total_cadastrados * 100)
    """
    wb = openpyxl.load_workbook(caminho, data_only=True)
    sh = wb.active
    try:
        idx_insp = _col_index(sh, "ÚLTIMA INSPEÇÃO")
    except Exception:
        idx_insp = _col_index(sh, "INSPECAO")

    total_cadastrados = 0
    anos = set()
    contagem_por_ano = {}

    for r in sh.iter_rows(min_row=2, values_only=True):
        if any(cell not in (None, "", " ") for cell in r):
            total_cadastrados += 1
            val = r[idx_insp] if idx_insp < len(r) else None
            ano = _extrair_ano(val)
            if ano is not None:
                anos.add(ano)
                contagem_por_ano[ano] = contagem_por_ano.get(ano, 0) + 1

    if not anos:
        return total_cadastrados, None, 0, 0.0

    ultimo_ano = max(anos)
    inspec_ultimo = contagem_por_ano.get(ultimo_ano, 0)
    cobertura = (inspec_ultimo / total_cadastrados * 100) if total_cadastrados > 0 else 0.0
    return total_cadastrados, ultimo_ano, inspec_ultimo, cobertura

def _fmt_int(n: int) -> str:
    return f"{n:,}".replace(",", ".")

class MenuPrincipal(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Menu Principal")
        self.geometry("980x640")
        self.minsize(880, 600)

        # refs do logo
        self._logo_pil = None
        self._logo_ctk = None

        # ref do alerta
        self._alert_label = None

        self.create_menu()

    # ==================== LAYOUT ====================
    def create_menu(self):
        # Barra lateral
        menu_frame = ctk.CTkFrame(self, width=200, height=640, corner_radius=0, fg_color="#054721")
        menu_frame.place(x=0, y=0)

        btn_usuario = ctk.CTkButton(menu_frame, text="USUÁRIO", width=150, height=50, command=self.open_user)
        btn_usuario.place(x=25, y=50)

        btn_pesquisar = ctk.CTkButton(menu_frame, text="PESQUISAR", width=150, height=50, command=self.open_pesquisar)
        btn_pesquisar.place(x=25, y=130)

        btn_cadastro = ctk.CTkButton(menu_frame, text="CADASTRO", width=150, height=50, command=self.open_cadastro)
        btn_cadastro.place(x=25, y=210)

        btn_Atualizar = ctk.CTkButton(menu_frame, text="ATUALIZAÇÃO", width=150, height=50, command=self.Atualizar)
        btn_Atualizar.place(x=25, y=290)

        btn_Funcao = ctk.CTkButton(menu_frame, text="ESTATÍSTICAS", width=150, height=50, command=self.Funcao)
        btn_Funcao.place(x=25, y=370)

        # Área principal (direita)
        self.main_area = ctk.CTkFrame(self, width=780, height=640, corner_radius=0, fg_color="#f8f8f8")
        self.main_area.place(x=200, y=0)

        # Container central (texto + logo)
        self.hero = ctk.CTkFrame(self.main_area, fg_color="transparent")
        self.hero.place(relx=0.5, rely=0.40, anchor=CENTER)

        self.title_lbl = ctk.CTkLabel(
            self.hero,
            text="Bem-vindo ao Sistema de Inspeções Visa Taiobeiras!",
            font=("Segoe UI Semibold", 22),
            text_color="black"
        )
        self.title_lbl.pack(pady=(0, 16))

        # Label para a logo
        self.logo_lbl = ctk.CTkLabel(self.hero, text="")
        self.logo_lbl.pack()

        # Carrega/ajusta a logo e alerta inicial
        self._init_logo()
        self.main_area.bind("<Configure>", self._resize_logo)
        self.after(150, self.show_alert_stats)  # mostra alerta ao abrir

    # ==================== ALERTA ====================
    def _show_alert(self, text: str, duration_ms: int = 9000):
        """Mostra um ALERTA textual simples (sem cartão) e some sozinho."""
        # remove alerta anterior, se existir
        if self._alert_label is not None:
            try:
                self._alert_label.destroy()
            except Exception:
                pass
            self._alert_label = None

        self._alert_label = ctk.CTkLabel(
            self.main_area,
            text=text,
            font=("Segoe UI Semibold", 14),
            fg_color="transparent",
            text_color="red",
        )
        # Base central da área principal
        self._alert_label.place(relx=0.5, rely=0.96, anchor="s")

        # Agenda sumir
        self.after(duration_ms, self._hide_alert)

    def _hide_alert(self):
        if self._alert_label is not None:
            try:
                self._alert_label.destroy()
            except Exception:
                pass
            self._alert_label = None

    def show_alert_stats(self):
        """Calcula e mostra o alerta textual com os números mais recentes."""
        try:
            total, ultimo_ano, inspec, cobertura = calcular_kpis_home("cadastros.xlsx")
            if ultimo_ano is not None:
                msg = f"Total: {_fmt_int(total)} • Inspecionados {ultimo_ano}: {_fmt_int(inspec)} • Cobertura: {cobertura:.2f}%"
            else:
                msg = f"Total: {_fmt_int(total)} • Nenhuma inspeção encontrada"
            self._show_alert(msg)
        except FileNotFoundError:
            self._show_alert("Arquivo 'cadastros.xlsx' não encontrado.")
        except Exception as e:
            self._show_alert(f"Erro ao calcular estatísticas: {e}")

    # ==================== LOGO ====================
    def _init_logo(self):
        """
        Carrega a logo:
        1) Tenta logo.png local
        2) Se não encontrar, usa LOGO_B64 embutido
        3) Fallback textual
        """
        self._logo_pil = None

        # 1️⃣ tenta arquivo logo.png local
        for path in ["/mnt/data/logo.png", "logo.png"]:
            if os.path.exists(path):
                try:
                    self._logo_pil = Image.open(path).convert("RGBA")
                    break
                except Exception:
                    self._logo_pil = None

        # 2️⃣ se não houver arquivo, tenta base64 embutido
        if self._logo_pil is None and LOGO_B64:
            try:
                data = base64.b64decode(LOGO_B64)
                self._logo_pil = Image.open(BytesIO(data)).convert("RGBA")
            except Exception:
                self._logo_pil = None

        # 3️⃣ fallback textual
        if self._logo_pil is None:
            self.logo_lbl.configure(
                text="(adicione logo.png ou gere logo_b64.py)",
                font=("Segoe UI", 12),
                text_color="gray"
            )
        else:
            self._apply_logo_with_size(360, 360)

    def _apply_logo_with_size(self, w, h):
        if not self._logo_pil:
            return
        ow, oh = self._logo_pil.size
        if ow == 0 or oh == 0:
            return
        scale = min(w / ow, h / oh)
        tw, th = max(1, int(ow * scale)), max(1, int(oh * scale))
        img = self._logo_pil.resize((tw, th), Image.LANCZOS)
        self._logo_ctk = ctk.CTkImage(light_image=img, dark_image=img, size=(tw, th))
        self.logo_lbl.configure(image=self._logo_ctk, text="")

    def _resize_logo(self, event=None):
        if not self._logo_pil:
            return
        area_w = max(280, self.main_area.winfo_width() - 120)
        area_h = max(220, self.main_area.winfo_height() - 260)
        self._apply_logo_with_size(area_w, area_h)

    # ==================== AÇÕES DO MENU ====================
    def open_user(self):
        """Abre a janela de gerenciamento de usuário."""
        if UsuarioApp is None:
            messagebox.showerror(
                "Erro",
                "Módulo 'usuario.py' não encontrado!\n\n"
                "Certifique-se de que o arquivo 'usuario.py' está na mesma pasta do main.py\n\n"
                "Se não tiver o arquivo, faça download e copie para a pasta do projeto."
            )
            return

        try:
            UsuarioApp(self)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir módulo de usuário:\n{e}")

    def open_pesquisar(self):
        PesquisarApp(self)

    def open_cadastro(self):
        CadastroApp(self)

    def Atualizar(self):
        AtualizarApp(self)
        self.show_alert_stats()

    def Funcao(self):
        if EstatisticaApp is None:
            messagebox.showinfo("Estatísticas", "A aba de Estatísticas não está disponível neste ambiente.")
            return
        EstatisticaApp(self)


if __name__ == "__main__":
    menu = MenuPrincipal()
    menu.mainloop()
