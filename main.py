import tkinter as tk
from tkinter import ttk
import os

# Importa os módulos que criamos na pasta 'modulos'
from modulos import cadastro, busca 

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Controle - Wickbold")
        self.root.geometry("400x350")
        
        # --- CARREGAR TEMA ---
        self.carregar_tema()

        # --- LAYOUT PRINCIPAL ---
        main_frame = ttk.Frame(root)
        main_frame.pack(expand=True, fill='both', padx=20, pady=20)

        # Título
        lbl_titulo = ttk.Label(main_frame, text="Gerenciador de Ativos", font=("Segoe UI", 16, "bold"))
        lbl_titulo.pack(pady=(10, 30))

        # Botão 1: Adicionar
        btn_add = ttk.Button(main_frame, text="📂 Adicionar Novo Registro", command=self.abrir_cadastro)
        btn_add.pack(pady=10, fill='x', ipadx=10, ipady=10)

        # Botão 2: Buscar
        btn_search = ttk.Button(main_frame, text="🔍 Pesquisar Colaborador", command=self.abrir_busca)
        btn_search.pack(pady=10, fill='x', ipadx=10, ipady=10)
        
        # Botão 3: Sair
        btn_sair = ttk.Button(main_frame, text="Sair", command=root.quit, style="Accent.TButton")
        btn_sair.pack(side='bottom', pady=20)

        # # Créditos
        # lbl_footer = ttk.Label(main_frame, text="v1.0 - Dev. Victor Alejandro", font=("Segoe UI", 8))
        # lbl_footer.pack(side='bottom', pady=5)

    def carregar_tema(self):
        """ Tenta carregar o tema Forest para o menu ficar padronizado """
        try:
            diretorio_base = os.path.dirname(os.path.abspath(__file__))
            caminho_tema = os.path.join(diretorio_base, 'theme', 'forest-dark.tcl')
            
            if os.path.exists(caminho_tema):
                self.root.tk.call("source", caminho_tema)
                ttk.Style().theme_use("forest-dark")
        except Exception as e:
            print(f"Aviso: Não foi possível carregar o tema. {e}")

    def abrir_cadastro(self):
        """ Chama a função do arquivo modulos/cadastro.py """
        cadastro.abrir_janela_cadastro(self.root)

    def abrir_busca(self):
        """ Chama a função do arquivo modulos/busca.py """
        busca.abrir_janela_busca(self.root)

if __name__ == "__main__":
    root = tk.Tk()
    
    # Centraliza a janela na tela (opcional)
    largura_janela = 400
    altura_janela = 350
    largura_tela = root.winfo_screenwidth()
    altura_tela = root.winfo_screenheight()
    pos_x = (largura_tela // 2) - (largura_janela // 2)
    pos_y = (altura_tela // 2) - (altura_janela // 2)
    root.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

    app = MainApp(root)
    root.mainloop()