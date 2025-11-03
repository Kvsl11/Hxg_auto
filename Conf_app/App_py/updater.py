import requests, os, ssl, subprocess, sys, tkinter as tk
from tkinter import messagebox

# --- Ignora SSL corporativo (seguro em rede interna) ---
ssl._create_default_https_context = ssl._create_unverified_context
requests.packages.urllib3.disable_warnings()

# --- Configurações principais ---
REPO = "Kvsl11/Hxg_auto"
URL_VERSION = f"https://raw.githubusercontent.com/{REPO}/main/version.txt"
URL_SCRIPT = f"https://raw.githubusercontent.com/{REPO}/main/main.py"
LOCAL_SCRIPT = "main.py"
LOCAL_VERSION_FILE = "version_local.txt"

# Caminho do Python interno (sem console)
PYTHONW_PATH = os.path.join(os.getcwd(), "Python313", "pythonw.exe")

# --- Funções auxiliares ---
def get_local_version():
    """Lê a versão local salva no arquivo."""
    if os.path.exists(LOCAL_VERSION_FILE):
        with open(LOCAL_VERSION_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "0.0.0"

def get_online_version():
    """Obtém versão mais recente do GitHub."""
    try:
        r = requests.get(URL_VERSION, timeout=10, verify=False)
        if r.status_code == 200:
            return r.text.strip()
    except Exception:
        pass
    return None

def atualizar_script():
    """Baixa nova versão do main.py."""
    try:
        r = requests.get(URL_SCRIPT, timeout=20, verify=False)
        r.raise_for_status()
        if os.path.exists(LOCAL_SCRIPT):
            os.remove(LOCAL_SCRIPT)
        with open(LOCAL_SCRIPT, "wb") as f:
            f.write(r.content)
        return True
    except Exception:
        return False

def save_local_version(ver):
    """Salva versão local após atualização."""
    with open(LOCAL_VERSION_FILE, "w", encoding="utf-8") as f:
        f.write(ver)

# --- Interface visual de atualização ---
def show_update_ui(local_v, online_v):
    win = tk.Tk()
    win.title("Atualização disponível - Auto-Ficha-OPE")
    win.geometry("400x280")
    win.configure(bg="#232323")
    win.resizable(False, False)

    tk.Label(win, text="🔄 Atualização disponível!",
             fg="white", bg="#232323", font=("Segoe UI", 13, "bold")).pack(pady=15)
    tk.Label(win, text=f"Versão atual: {local_v}\nNova versão: {online_v}",
             fg="#ccc", bg="#232323").pack(pady=5)
    status_label = tk.Label(win, text="", fg="#3ba55d", bg="#232323")
    status_label.pack(pady=5)

    def atualizar_agora():
        btn_update.config(state="disabled", text="Baixando...")
        win.update()
        ok = atualizar_script()
        if ok:
            save_local_version(online_v)
            status_label.config(text="✅ Atualização concluída!")
            win.update()
            messagebox.showinfo("Atualizado", "O aplicativo foi atualizado com sucesso!")
            win.destroy()
            iniciar_app()
        else:
            messagebox.showerror("Erro", "⚠️ Falha ao baixar atualização. Tente novamente mais tarde.")
            win.destroy()
            iniciar_app()

    btn_update = tk.Button(win, text="Atualizar agora", command=atualizar_agora,
                           bg="#3ba55d", fg="white", font=("Segoe UI", 11, "bold"), width=20)
    btn_update.pack(pady=15)

    tk.Button(win, text="Ignorar", command=lambda: [win.destroy(), iniciar_app()],
              bg="#555", fg="white", width=20).pack(pady=5)

    win.protocol("WM_DELETE_WINDOW", lambda: [win.destroy(), iniciar_app()])
    win.mainloop()

# --- Função para iniciar o app ---
def iniciar_app():
    """Executa o app principal usando pythonw.exe sem console."""
    python_exe = PYTHONW_PATH if os.path.exists(PYTHONW_PATH) else sys.executable
    try:
        subprocess.Popen(
            [python_exe, LOCAL_SCRIPT],
            creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
    except Exception:
        pass
    finally:
        os._exit(0)

# --- Execução principal ---
def main():
    # Garante que não haja console mesmo se rodar via python.exe
    if sys.stdout is not None:
        try:
            sys.stdout.flush()
            sys.stdout = open(os.devnull, 'w')
            sys.stderr = open(os.devnull, 'w')
        except Exception:
            pass

    local_v = get_local_version()
    online_v = get_online_version()

    if not online_v:
        iniciar_app()
        return

    if online_v != local_v:
        show_update_ui(local_v, online_v)
    else:
        iniciar_app()

if __name__ == "__main__":
    main()
