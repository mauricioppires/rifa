#!C:\\python3\\Python312\\python.exe
# -*- coding: utf-8 -*-

"""
Mauricio P Pires <mauricioppires at gmail dot com>
"""

import tkinter as tk
from tkinter import messagebox

modulos_necessarios = ['pandas']

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

def listar_linhas_e_renomear():
    try:
        df = pd.read_excel("rifa.xlsx", sheet_name="lista", header=None)
        df_selecionado = df.iloc[0:100, 0:3]
        df_selecionado.columns = ['Número', 'Opção', 'Participante']

        root = tk.Tk()
        root.title("Listagem dos Participantes")
        row = 0
        col = 0
        for index, row_data in df_selecionado.iterrows():
            label_text = f"{' | '.join(map(str, row_data.tolist()))}"
            label = tk.Label(root, text=label_text, padx=10, pady=5)
            label.grid(row=row, column=col, sticky='w')
            row += 1
            if row == 25:
                row = 0
                col += 1

        fechar_button = tk.Button(root, text="Fechar", command=root.destroy)
        fechar_button.grid(row=25, column=0, columnspan=4, pady=10)

        root.mainloop()

    except FileNotFoundError:
        messagebox.showerror("Erro", "Arquivo rifa.xlsx não encontrado")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")


root = tk.Tk()
root.title("Rifa - Listagem do Participantes")

listar_button = tk.Button(root, text="Listar", command=listar_linhas_e_renomear)
listar_button.pack(pady=20)

root.mainloop()
