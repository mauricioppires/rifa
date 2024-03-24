#!C:\\python3\\Python312\\python.exe
# -*- coding: utf-8 -*-

"""
Mauricio P Pires <mauricioppires at gmail dot com>
"""

import tkinter as tk
from tkinter import Label, Button, messagebox
import random

modulos_necessarios = ['pandas', 'openpyxl']

missing_modules = []
for module in modulos_necessarios:
    try:
        if module == 'pandas':
            alias = 'pd'
            imported_module = __import__(module)
            globals()[alias] = imported_module
        else:
            __import__(module)
    except ImportError:
        missing_modules.append(module)

if missing_modules:
    messagebox.showerror("Módulos ausentes", f"Os seguintes módulos estão faltando: {', '.join(missing_modules)}")
    exit()

caminho_arquivo = 'rifa.xlsx'
df = pd.read_excel(caminho_arquivo, sheet_name='lista')
lista = df.iloc[0:100, 0:3].values.tolist()

class GeradorRandomico:
    def __init__(self, root):
        self.root = root
        self.root.title("Rifa - Sorteio")
        self.root.attributes('-fullscreen', True)

        self.numero_label = Label(root, text="", font=("Helvetica", 90))
        self.numero_label.pack(pady=(root.winfo_screenheight() - 200) / 2)

        self.gerar_numero_button = Button(root, text="Sortear", command=self.iniciar_geracao)
        self.gerar_numero_button.pack()

        self.sair_button = Button(root, text="Sair", command=self.sair)
        self.sair_button.pack()

        self.giros_restantes = 18

    def iniciar_geracao(self):
        self.giros_restantes = 18
        self.gerar_numero()

    def gerar_numero(self):
        if self.giros_restantes > 0:
            numero_aleatorio = random.randint(1, 100)
            self.numero_label.config(text=str(numero_aleatorio))
            self.giros_restantes -= 1
            self.root.after(100, self.gerar_numero)
        else:
            ultimo_numero = int(self.numero_label.cget("text"))
            mensagem_final = f"{ultimo_numero} - "
            # Localizar na lista
            for item in lista:
                if item[0] == ultimo_numero:
                    nome_mulher = item[1]
                    nome_terceira_posicao = item[2]
                    mensagem_final += f"{nome_mulher} - {nome_terceira_posicao}"
            self.numero_label.config(text=mensagem_final)

    def sair(self):
        resposta = messagebox.askyesno("Sair", "Deseja realmente sair?")
        if resposta:
            self.root.destroy()

if __name__ == "__main__":
    rroot = tk.Tk()
    app = GeradorRandomico(rroot)
    rroot.mainloop()
