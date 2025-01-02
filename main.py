import customtkinter as ctk
from tkinter import *
from tkinter import messagebox

# Importando as janelas de outras abas
from cadastro import CadastroApp
from pesquisar import PesquisarApp
from Atualizar import AtualizarApp

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

class MenuPrincipal(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Menu Principal")
        self.geometry("800x600")
        
        self.create_menu()

    def create_menu(self):
        menu_frame = ctk.CTkFrame(self, width=200, height=600, corner_radius=0, fg_color="#004aad")
        menu_frame.place(x=0, y=0)

        btn_usuario = ctk.CTkButton(menu_frame, text="USUÁRIO", width=150, height=50, command=self.open_user)
        btn_usuario.place(x=25, y=50)

        btn_pesquisar = ctk.CTkButton(menu_frame, text="PESQUISAR", width=150, height=50, command=self.open_pesquisar)
        btn_pesquisar.place(x=25, y=130)

        btn_cadastro = ctk.CTkButton(menu_frame, text="CADASTRO", width=150, height=50, command=self.open_cadastro)
        btn_cadastro.place(x=25, y=210)

        btn_Atualizar = ctk.CTkButton(menu_frame, text="ATUALIZAÇÃO", width=150, height=50, command=self.Atualizar)
        btn_Atualizar.place(x=25, y=290)

        btn_Funcao = ctk.CTkButton(menu_frame, text="FUNÇÃO", width=150, height=50, command=self.Funcao)
        btn_Funcao.place(x=25, y=370)

        main_area = ctk.CTkFrame(self, width=600, height=600, corner_radius=0, fg_color="lightgray")
        main_area.place(x=200, y=0)
        ctk.CTkLabel(main_area, text="Bem-vindo ao Menu Principal!", font=("Arial", 18)).place(relx=0.5, rely=0.5, anchor=CENTER)

    def open_user(self):
        messagebox.showinfo("Usuário", "A funcionalidade de USUÁRIO será implementada aqui.")

    def open_pesquisar(self):
        # Abre janela de pesquisa
        PesquisarApp(self)

    def open_cadastro(self):
        # Abre janela de cadastro
        CadastroApp(self)

    def Atualizar(self):
        AtualizarApp(self)

    def Funcao(self):
        messagebox.showinfo("FUNÇÃO", "A funcionalidade de FUNÇÃO será implementada aqui.")


if __name__ == "__main__":
    menu = MenuPrincipal()
    menu.mainloop()

