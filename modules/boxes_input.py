import tkinter as tk
from tkinter import simpledialog

def app_option(): 
    result = {"value": None}
    # Cria a janela principal
    root = tk.Tk()
    root.title("Selecione a opção")
    root.geometry("320x150")
    root.resizable(False, False)
    root.eval('tk::PlaceWindow . center')
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass
    def choose(value: str):
        result["value"] = value
        root.destroy()
    root.bind("<Escape>", lambda e: root.destroy())
    frm = tk.Frame(root, padx=12, pady=12)
    frm.pack(fill="both", expand=True)
    lbl = tk.Label(frm, text="O que você deseja baixar?", font=("Segoe UI", 11))
    lbl.pack(pady=(0, 10))
    btn_extratos = tk.Button(
        frm, text="Extratos", width=24, command=lambda: choose("extratos")
    )
    btn_extratos.pack(pady=5)
    btn_informe = tk.Button(
        frm, text="Informe de Rendimento", width=24, command=lambda: choose("informe")
    )
    btn_informe.pack(pady=5)
    btn_extratos.focus_set()
    # Loop da janela
    root.mainloop()
    return result["value"]


def input_token_btg():
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    token = simpledialog.askstring("Input Token", "Digite o token (apenas números):")
    
    if token and token.isdigit():
        return token
    else:
        print("Token inválido. Certifique-se de digitar apenas números.")
        return None
