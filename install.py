import subprocess, sys, os

DEPENDENCIAS = ["requests", "tkinter"]

def instalar_dependencias():
    for pacote in DEPENDENCIAS:
        subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", pacote])

def iniciar_atualizador():
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    updater = os.path.join(APP_DIR, "updater_gui.py")
    subprocess.run([sys.executable, updater])

if __name__ == "__main__":
    instalar_dependencias()
    iniciar_atualizador()
