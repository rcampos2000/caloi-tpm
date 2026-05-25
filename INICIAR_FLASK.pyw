"""
Launcher do Sistema de Manutenção TPM - Caloi
Duplo clique para iniciar o servidor Flask
"""
import subprocess
import sys
import os
import webbrowser
import time
import tkinter as tk
from tkinter import messagebox

def instalar_dependencias():
    try:
        import flask
        import openpyxl
    except ImportError:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("TPM Caloi", "Instalando dependências (Flask, openpyxl)...\nAguarde alguns segundos.")
        subprocess.run([sys.executable, "-m", "pip", "install", "flask", "openpyxl", "--quiet"], check=True)

def main():
    # Mudar para o diretório do script
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # Instalar dependências se necessário
    instalar_dependencias()

    # Abrir navegador após 2 segundos
    def abrir_navegador():
        time.sleep(2)
        webbrowser.open("http://localhost:5000")

    import threading
    t = threading.Thread(target=abrir_navegador)
    t.daemon = True
    t.start()

    # Iniciar Flask
    try:
        import app
        app.app.run(debug=False, host='0.0.0.0', port=5000)
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Erro", f"Erro ao iniciar servidor:\n{e}")

if __name__ == "__main__":
    main()
