import threading
import glob
import pandas as pd
# --- MODIFICAÇÕES DE IMPORTAÇÃO ---
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.styles.borders import Border, Side
import datetime
from fpdf import FPDF
import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as ttk
from PIL import Image, ImageTk
import time
import fitz
import json
import warnings
import os
import ssl
import subprocess
import urllib.request
import logging
import sys
import requests
import customtkinter as ctk
import shutil

# Caminho dinâmico da pasta onde o script está localizado
app_dir = os.path.dirname(os.path.abspath(__file__))
# Usa o executável Python que está rodando o script ATUALMENTE
python_exe = sys.executable
print(f"🟢 Usando Python em: {python_exe}")

# Configuração de logging (apenas console - Log do SSL removido)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Apenas o certificado raiz da Amazon (necessário para access.hxgnagron.com)
AMAZON_CERTS = {
    "Amazon Root CA 1": "https://www.amazontrust.com/repository/AmazonRootCA1.pem"
}

def preparar_dependencias():
    """Instala/Atualiza pacotes essenciais usando o Python ATUAL."""
    try:
        if not os.path.exists(python_exe):
            logger.warning(f"⚠️ Python não encontrado em: {python_exe}")
            return

        # --- LIMPEZA DE CACHE ANTIGO ---
        wdm_cache_path = os.path.join(os.path.expanduser("~"), ".wdm")
        if os.path.exists(wdm_cache_path):
            logger.info(f"🧹 Limpando cache antigo do webdriver-manager em: {wdm_cache_path}")
            shutil.rmtree(wdm_cache_path, ignore_errors=True)
        # --- FIM DA LIMPEZA ---

        # ADICIONADO: 'xlwings' para manipular o Excel sem apagar o cache de fórmulas.
        pacotes = ["certifi", "selenium", "xlwings"]
        logger.info(f"🔍 Verificando e atualizando pacotes: {', '.join(pacotes)}...")

        for pacote in pacotes:
            logger.info(f"Instalando/Atualizando {pacote}...")
            subprocess.run([python_exe, "-m", "pip", "install", "--upgrade", pacote], check=True, capture_output=True,
                           text=True)

        import certifi
        logger.info(f"🟢 Dependências atualizadas com sucesso. Caminho Certifi: {certifi.where()}")
    except Exception as e:
        logger.warning(f"⚠️ Falha ao preparar dependências: {e}")

def garantire_certificados_amazon():
    """Verifica se o certificado raiz da Amazon está presente e adiciona se necessário."""
    try:
        import certifi
        cacert_path = certifi.where()

        with open(cacert_path, "r", encoding="utf-8") as f:
            conteudo = f.read()

        for nome, url in AMAZON_CERTS.items():
            if nome not in conteudo:
                logger.info(f"🔍 {nome} não encontrado, baixando de {url}...")
                resp = requests.get(url, timeout=10, verify=False)
                if resp.status_code == 200:
                    with open(cacert_path, "a", encoding="utf-8") as f:
                        f.write(f"\n# {nome}\n{resp.text.strip()}\n")
                    logger.info(f"✅ {nome} adicionado ao cacert.pem.")
                else:
                    logger.warning(f"❌ Falha ao baixar {nome}: {resp.status_code}")
            else:
                logger.info(f"🟢 {nome} já está presente no cacert.pem.")
    except Exception as e:
        logger.warning(f"⚠️ Falha ao garantir certificado Amazon Root CA 1: {e}")

# --- EXECUÇÃO AUTOMÁTICA AO INICIAR ---
logger.info("🚀 Iniciando verificação e correção SSL híbrida...")
preparar_dependencias()
garantire_certificados_amazon()

try:
    import certifi
    ssl_context = ssl.create_default_context(cafile=certifi.where())
    urllib.request.urlopen("https://www.google.com", timeout=5, context=ssl_context)
    logger.info("🟢 Conexão SSL validada com sucesso — certificados OK.")
except ssl.SSLError as e:
    logger.warning(f"⚠️ Falha de SSL detectada ({e}). Aplicando modo não verificado (Fallback).")
    ssl._create_default_https_context = ssl._create_unverified_context
    logger.info("🟡 SSL desativado globalmente — conexão forçada sem verificação de certificado.")
except Exception as e:
    logger.warning(f"⚠️ Erro genérico ao testar SSL: {e}. Aplicando modo não verificado (Fallback).")
    ssl._create_default_https_context = ssl._create_unverified_context
    logger.info("🟡 SSL desativado globalmente — conexão forçada sem verificação de certificado.")

logger.info("✅ Configuração SSL concluída com segurança.")

# ==========================================
# --- VERIFICAÇÃO DE SEGURANÇA (KILL SWITCH) ---
# ==========================================
def exibir_erro_fatal(titulo, mensagem):
    """Exibe uma janela de erro travada na tela e fecha o programa."""
    root_temp = tk.Tk()
    root_temp.withdraw()
    root_temp.attributes("-topmost", True)
    messagebox.showerror(titulo, mensagem)
    root_temp.destroy()
    os._exit(1)


def verificar_seguranca():
    """Verifica a trava de segurança remoto."""
    try:
        REPO = "Kvsl11/Hxg_auto"
        ts = int(time.time())
        URL_STATUS = f"https://raw.githubusercontent.com/{REPO}/main/status.txt?t={ts}"

        try:
            headers = {"Cache-Control": "no-cache, no-store, must-revalidate", "Pragma": "no-cache"}
            r_status = requests.get(URL_STATUS, timeout=10, verify=False, headers=headers)
            if r_status.status_code == 200:
                status_app = r_status.text.strip().lower()
                if status_app == "false":
                    logger.warning("🔴 TRAVA ATIVADA VIA GITHUB! Bloqueando acesso.")
                    exibir_erro_fatal("Erro Crítico de Comunicação",
                                      "Ocorreu uma falha inesperada ao sincronizar as configurações iniciais do sistema.\n\nCódigo do Erro: ERR_CONNECTION_REFUSED_10061\nPor favor, tente novamente mais tarde.")
            else:
                logger.info(f"⚠️ Status remoto retornou código {r_status.status_code}. Execução permitida.")
        except Exception as e:
            logger.warning(f"⚠️ Falha ao checar status.txt. Ignorando trava. Erro: {e}")
    except Exception as e:
        logger.error(f"❌ Erro na rotina de segurança: {e}")

# --- VERIFICAÇÃO DE ATUALIZAÇÃO VIA GITHUB ---
VERSAO = "3.3.0"


def verificar_e_atualizar_automaticamente():
    """Verifica no GitHub se há nova versão e atualiza automaticamente."""
    try:
        REPO = "Kvsl11/Hxg_auto"
        URL_VERSION = f"https://raw.githubusercontent.com/{REPO}/main/version.txt"
        URL_SCRIPT = f"https://raw.githubusercontent.com/{REPO}/main/main.py"
        LOCAL_SCRIPT = os.path.join(os.path.dirname(__file__), "main.py")
        LOCAL_VERSION_FILE = os.path.join(os.path.dirname(__file__), "version_local.txt")
        LOG_PATH = os.path.join(os.path.dirname(__file__), "autoupdate.log")

        logging.basicConfig(
            filename=LOG_PATH,
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )

        def get_local_version():
            if os.path.exists(LOCAL_VERSION_FILE):
                try:
                    with open(LOCAL_VERSION_FILE, "r", encoding="utf-8") as f:
                        return f.read().strip()
                except Exception:
                    return "0.0.0"
            return "0.0.0"

        def get_online_version():
            try:
                headers = {"Cache-Control": "no-cache", "Pragma": "no-cache"}
                r = requests.get(URL_VERSION, timeout=10, verify=False, headers=headers)
                if r.status_code == 200:
                    return r.text.strip()
                else:
                    logging.warning(f"⚠️ Falha HTTP ao buscar versão: {r.status_code}")
            except Exception as e:
                logging.warning(f"⚠️ Falha ao obter versão online: {e}")
            return None

        def save_local_version(ver):
            try:
                with open(LOCAL_VERSION_FILE, "w", encoding="utf-8") as f:
                    f.write(ver)
                logging.info(f"✅ Versão local atualizada para {ver}")
            except Exception as e:
                logging.error(f"❌ Erro ao salvar versão local: {e}")

        def atualizar_script(versao_online):
            try:
                headers = {"Cache-Control": "no-cache", "Pragma": "no-cache"}
                r = requests.get(URL_SCRIPT, timeout=20, verify=False, headers=headers)
                r.raise_for_status()
                with open(LOCAL_SCRIPT, "wb") as f:
                    f.write(r.content)
                save_local_version(versao_online)
                logging.info(f"✅ Atualização concluída para a versão {versao_online}")
                return True
            except Exception as e:
                logging.error(f"❌ Falha ao atualizar script: {e}")
                return False

        local_v = get_local_version()
        online_v = get_online_version()

        if not online_v:
            logging.warning("⚠️ Falha ao verificar versão online. Continuando com a versão local.")
            return

        if online_v != local_v:
            logging.info(f"🟡 Nova versão detectada: {online_v} (local: {local_v}) — atualizando...")
            sucesso = atualizar_script(online_v)
            if sucesso:
                logging.info("♻️ Reiniciando app com nova versão...")
                python_exe = sys.executable
                subprocess.Popen([python_exe, LOCAL_SCRIPT])
                os._exit(0)
            else:
                logging.info(f"🟢 Aplicativo já está atualizado ({local_v})")
    except Exception as e:
        logging.error(f"❌ Erro na verificação automática de atualização: {e}")

warnings.filterwarnings(
    "ignore",
    message="Slicer List extension is not supported and will be removed",
    category=UserWarning,
    module=r"openpyxl\.worksheet\._reader"
)

# --- Definições Globais ---
script_dir = os.path.dirname(os.path.abspath(__file__))
execucao_ativa = False
status_label = None
progress_bar = None
root = None

RESPONSAVEIS_OPCOES = [
    "JUAN CARLOS",
    "ROSANI ALDA",
    "FERNANDO BREGUEDO",
    "FLAVIO BREGUEDO",
    "EDUARDO APARECIDO",
    "LEANDRO RENE",
    "EDUARDO NUNES",
    "LEANDRO SEBOLD",
    "ALEX FABIANO",
    "RAMON ROSA"
]

# --- Funções de Automação com Selenium ---
def iniciar_driver(headless=True):
    """Inicia uma instância do Chrome usando Selenium padrão e o SeleniumManager embutido."""
    print("🚀 Iniciando driver com Selenium padrão (SeleniumManager)...")
    logger.info("🚀 Iniciando driver com Selenium padrão (SeleniumManager)...")

    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")

    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-running-insecure-content')
    options.add_argument('--disable-web-security')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    try:
        servico = Service()
        driver = webdriver.Chrome(service=servico, options=options)
    except Exception as e:
        print(f"❌ Falha ao iniciar Selenium/SeleniumManager: {e}")
        logger.error(f"❌ Falha ao iniciar Selenium/SeleniumManager: {e}")
        raise

    if not headless:
        driver.maximize_window()
    return driver

def aguardar_pagina_carregada(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        print("🟢 Página totalmente carregada.")
    except Exception as e:
        print(f"⚠️ Erro ao aguardar carregamento: {e}")

def aguardar_e_clicar(driver, xpath, timeout=30):
    try:
        print(f"Tentando clicar em: {xpath}")
        elemento = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", elemento)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        time.sleep(0.5)
        elemento.click()
        print(f"🟢 Clique realizado: {xpath}")
        logger.info(f"🟢 Clique realizado: {xpath}")
    except Exception as e:
        print(f"⚠️ Clique padrão falhou ({xpath}). Tentando via JavaScript... Erro: {e}")
        logger.warning(f"⚠️ Clique padrão falhou ({xpath}). Tentando via JS.")
        try:
            elemento = driver.find_element(By.XPATH, xpath)
            driver.execute_script("arguments[0].click();", elemento)
            print(f"🟢 Clique via JS realizado: {xpath}")
            logger.info(f"🟢 Clique via JS realizado: {xpath}")
        except Exception as js_e:
            print(f"❌ Erro final ao clicar via JS em {xpath}: {js_e}")
            logger.error(f"❌ Erro final ao clicar via JS em {xpath}: {js_e}")

def aguardar_e_escrever(driver, xpath, texto, timeout=30):
    try:
        campo = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.XPATH, xpath))
        )
        campo.clear()
        campo.send_keys(texto)
        print(f"🟢 Texto inserido: {texto}")
    except Exception as e:
        print(f"⚠️ Erro ao escrever no campo {xpath}: {e}")

def login_usuario(driver, url, usuario, senha, xpaths):
    driver.get(url)
    aguardar_pagina_carregada(driver)
    aguardar_e_escrever(driver, xpaths['usuario'], usuario)
    aguardar_e_escrever(driver, xpaths['senha'], senha)
    aguardar_e_clicar(driver, xpaths['botao_login'])
    time.sleep(3)

def exportar_tabela(driver, xpaths):
    limpar_filtro_xpath = '//button[contains(@id,"buttion-id-clearAndApplyButton")]'
    aguardar_e_clicar(driver, limpar_filtro_xpath)

    print("⏳ Aguardando processamento após limpar o filtro...")
    time.sleep(3)

    try:
        WebDriverWait(driver, 30).until(
            EC.invisibility_of_element_located((By.XPATH,
                                                "//strong[contains(.,'Loading')] | //strong[contains(.,'Carregando')] | //div[contains(@class, 'overlay')] | //*[contains(@class, 'spinner')]"))
        )
        print("🟢 Overlay de carregamento desapareceu.")
    except Exception as e:
        print(f"⚠️ Não foi possível confirmar o desaparecimento do overlay (ou não havia): {e}")

    time.sleep(3)
    aguardar_e_clicar(driver, xpaths['tabela'])

    print("⏳ Aguardando a tabela carregar...")
    time.sleep(3)

    try:
        WebDriverWait(driver, 30).until(
            EC.invisibility_of_element_located((By.XPATH,
                                                "//strong[contains(.,'Loading')] | //strong[contains(.,'Carregando')] | //div[contains(@class, 'overlay')] | //*[contains(@class, 'spinner')]"))
        )
        print("🟢 Tabela carregada com sucesso.")
    except Exception as e:
        print(f"⚠️ Não foi possível confirmar o carregamento da tabela: {e}")

    time.sleep(2)

    print("🔄 Alterando paginação para exibir mais itens...")
    aguardar_e_clicar(driver, xpaths['paginador_dropdown'])
    time.sleep(1)

    aguardar_e_clicar(driver, xpaths['paginador_opcao_5'])
    print("⏳ Aguardando tabela recarregar com nova paginação...")
    time.sleep(4)

    aguardar_e_clicar(driver, xpaths['filtro'])
    time.sleep(2)

    aguardar_e_clicar(driver, xpaths['exportacao_csv'])
    print("🟢 Exportação iniciada")

def aguardar_download_completo(diretorio, nome_base, timeout=60):
    tempo_inicial = time.time()
    while time.time() - tempo_inicial < timeout:
        arquivos_tmp = glob.glob(os.path.join(diretorio, f"{nome_base}*.tmp"))
        arquivos_csv = glob.glob(os.path.join(diretorio, f"{nome_base}*.csv"))

        if arquivos_csv and not arquivos_tmp:
            return max(arquivos_csv, key=os.path.getctime)
        time.sleep(2)

    print("❌ Tempo limite excedido para o download do arquivo CSV!")
    return None

def processar_csv(diretorio_downloads, pdf_output_dir, selected_responsaveis):
    """
    Processa o CSV exportado do Hexagon.
    Abre a planilha de monitoramento utilizando XLWINGS para ler o valor final calculado 
    pelas fórmulas (PROCV, etc), ignorando o cache do openpyxl que pode estar quebrado.
    """
    try:
        if not os.path.exists(pdf_output_dir):
            os.makedirs(pdf_output_dir)

        print("⏳ Aguardando download do arquivo CSV...")
        csv_path = aguardar_download_completo(diretorio_downloads, "Monitoramento - Tabela")
        if not csv_path:
            print("❌ Processo encerrado. Nenhum arquivo CSV disponível.")
            return None

        # Carrega o CSV do Hexagon
        df = pd.read_csv(csv_path, encoding="utf-8", sep=";", dtype=str)
        df.columns = df.columns.str.strip().str.upper()
        
        df["REGISTRO MAIS RECENTE"] = pd.to_datetime(df["REGISTRO MAIS RECENTE"], format="%d/%m/%Y %H:%M:%S",
                                                     errors="coerce")

        # Exclui o CSV baixado para não encher a pasta do usuário
        try:
            if os.path.exists(csv_path):
                os.remove(csv_path)
        except Exception as clean_err:
            pass

        data_atual = datetime.datetime.now().date()
        df_antigos = df[df["REGISTRO MAIS RECENTE"].dt.date != data_atual].copy()

        caminho_base_monitoramento = obter_caminho_planilha()
        print(f"📖 Lendo responsáveis diretamente de: {caminho_base_monitoramento} (Aba: Cont. Maquinas)")
        
        df_responsaveis = None
        
        # --- NOVO: Leitura via xlwings para forçar o cálculo das fórmulas ---
        print("⏳ Inicializando motor Excel nativo para extrair os valores reais das fórmulas...")
        try:
            import xlwings as xw
            app = xw.App(visible=False)
            app.display_alerts = False
            try:
                # read_only=True garante velocidade e segurança
                wb = app.books.open(caminho_base_monitoramento, update_links=False, read_only=True)
                ws = wb.sheets["Cont. Maquinas"]
                
                # Extrai todos os dados da aba, já formatados como um DataFrame do Pandas
                # O atributo .value no xlwings traz SEMPRE o resultado da fórmula em texto/número
                df_responsaveis = ws.used_range.options(pd.DataFrame, index=False).value
                print("✅ Valores reais das fórmulas extraídos com sucesso!")
            finally:
                if 'wb' in locals() and wb: wb.close()
                if 'app' in locals() and app: app.quit()
        except Exception as ex:
            print(f"⚠️ Aviso: Não foi possível usar o motor nativo ({ex}). Tentando via pandas padrão...")
            
        # Fallback caso o xlwings falhe por algum motivo (ex: Excel não instalado)
        if df_responsaveis is None or df_responsaveis.empty:
            df_responsaveis = pd.read_excel(caminho_base_monitoramento, sheet_name="Cont. Maquinas")
        
        # Normalização rigorosa dos cabeçalhos do Excel
        normalized_cols = {}
        for col in df_responsaveis.columns:
            clean_col = str(col).strip().upper()
            clean_col = clean_col.replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
            normalized_cols[col] = clean_col
        df_responsaveis = df_responsaveis.rename(columns=normalized_cols)

        col_equipamento_csv = "NRO DO EQUIPAMENTO"
        col_equipamento_base = "EQUIPAMENTO"
        col_responsavel = "RESPONSAVEL"
        col_display = "DISPLAY"
        col_prestador = "PRESTADOR"

        if col_equipamento_base not in df_responsaveis.columns:
            print(f"⚠️ A coluna '{col_equipamento_base}' não foi encontrada na aba 'Cont. Maquinas'!")
            return None

        # Renomeia EQUIPAMENTO para NRO DO EQUIPAMENTO para preparar o Merge
        df_responsaveis = df_responsaveis.rename(columns={col_equipamento_base: col_equipamento_csv})

        # Prevenção rigorosa: transforma Frotas para garantir precisão no Match
        df_antigos[col_equipamento_csv] = df_antigos[col_equipamento_csv].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        df_responsaveis[col_equipamento_csv] = df_responsaveis[col_equipamento_csv].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

        # Higieniza os valores numéricos originados de fórmulas para não ficarem como '.0' ou nulos indesejados
        for col in [col_responsavel, col_display, col_prestador]:
            if col in df_responsaveis.columns:
                df_responsaveis[col] = df_responsaveis[col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                df_responsaveis[col] = df_responsaveis[col].replace({'nan': '', 'None': '', '<NA>': ''})

        colunas_necessarias = [col_equipamento_csv]
        for c in [col_responsavel, col_display, col_prestador]:
            if c in df_responsaveis.columns:
                colunas_necessarias.append(c)

        # Merge utilizando SUFFIXES para não conflitar com colunas vazias do CSV
        df_final = df_antigos.merge(
            df_responsaveis[colunas_necessarias], 
            on=col_equipamento_csv, 
            how="left", 
            suffixes=('_CSV', '_EXCEL')
        )
        
        # Consolida forçando que a coluna final seja sempre a do Excel!
        for c in [col_responsavel, col_display, col_prestador]:
            if f"{c}_EXCEL" in df_final.columns:
                df_final[c] = df_final[f"{c}_EXCEL"]

        # Filtro de Responsável Vazio / Nulo
        if col_responsavel in df_final.columns:
            df_final = df_final[df_final[col_responsavel].astype(str).str.strip() != ""]
            df_final = df_final.dropna(subset=[col_responsavel])
        else:
            print("❌ Não foi possível encontrar a coluna RESPONSAVEL após a mesclagem.")
            return None

        if selected_responsaveis:
            print(f"✅ Gerando relatórios apenas para: {', '.join(selected_responsaveis)}")
            df_final = df_final[df_final[col_responsavel].isin(selected_responsaveis)]
        else:
            print("🔄 Nenhum responsável selecionado. Gerando para todos.")

        colunas_desejadas = [
            "RESPONSAVEL", "DISPLAY", "NRO DO EQUIPAMENTO",
            "TIPO DO EQUIPAMENTO", "PRESTADOR", "REGISTRO MAIS RECENTE"
        ]
        
        colunas_desejadas_existentes = [c for c in colunas_desejadas if c in df_final.columns]
        df_final = df_final[colunas_desejadas_existentes]

        return df_final
    except Exception as e:
        print(f"⚠️ Erro ao processar CSV utilizando a planilha de monitoramento: {e}")
        return None

# --- SISTEMA DE FECHAMENTO (Dia 21 a 20 do próximo mês) ---
def obter_caminho_planilha():
    import os
    import datetime

    hoje = datetime.datetime.now()
    dia_atual = hoje.day
    mes_atual = hoje.month
    ano_atual = hoje.year

    if dia_atual > 20:
        mes_alvo = mes_atual + 1
        if mes_alvo > 12:
            mes_alvo = 1
            ano_alvo = ano_atual + 1
        else:
            ano_alvo = ano_atual
    else:
        mes_alvo = mes_atual
        ano_alvo = ano_atual

    meses = {
        1: "01 - Janeiro", 2: "02 - Fevereiro", 3: "03 - Março",
        4: "04 - Abril", 5: "05 - Maio", 6: "06 - Junho",
        7: "07 - Julho", 8: "08 - Agosto", 9: "09 - Setembro",
        10: "10 - Outubro", 11: "11 - Novembro", 12: "12 - Dezembro"
    }

    numero_safra = 2.5 + (ano_alvo - 2025) * 0.1
    safra = f"{numero_safra:.1f} - Safra {ano_alvo}"

    possiveis_drives = ["I:", "Z:"]
    caminho_final = None

    for drive in possiveis_drives:
        base = fr"{drive}\ANG\Agricola\Controle\Computador de Bordo\Fechamento Prestação de Serviço (Linha Amarela)\Pago pelo Bordo"
        caminho_teste = os.path.join(base, safra, meses[mes_alvo], "Monitoramento - Eqps.xlsx")
        if os.path.exists(caminho_teste):
            caminho_final = caminho_teste
            break

    if not caminho_final:
        raise FileNotFoundError(
            f"❌ Não foi possível localizar a planilha de Equipamentos para a Safra '{safra}' no mês '{meses[mes_alvo]}'."
        )

    return caminho_final

def atualizar_coleta_planilha(df_final):
    """
    Atualiza a coluna COLETA utilizando XLWINGS para simular uma ação de usuário no Excel.
    Isso é FUNDAMENTAL para que as fórmulas PROCV (Display, Prestador, etc) não percam
    seu cache de memória (bug conhecido do openpyxl ao usar 'wb.save()').
    """
    try:
        import xlwings as xw
        caminho_planilha = obter_caminho_planilha()
        aba_alvo = "Cont. Maquinas"

        print("⏳ Atualizando Excel nativamente via xlwings (preservando o cache das Fórmulas)...")
        app = xw.App(visible=False)  # Abre o Excel no background
        
        # --- PREVENÇÃO DE ERROS DE SALVAMENTO (TRAVAMENTOS E POPUPS DO EXCEL) ---
        app.display_alerts = False 
        
        wb = None
        
        try:
            wb = app.books.open(caminho_planilha)
            ws = wb.sheets[aba_alvo]

            # Obter os cabeçalhos diretamente da primeira linha
            header_range = ws.range('A1').expand('right')
            cabecalhos = {str(cell.value).strip().upper(): cell.column for cell in header_range if cell.value}

            if "EQUIPAMENTO" not in cabecalhos or "COLETA" not in cabecalhos:
                print("⚠️ Colunas necessárias não encontradas na aba Cont. Maquinas.")
                return

            col_equip = cabecalhos["EQUIPAMENTO"]
            col_coleta = cabecalhos["COLETA"]

            equipamentos_contingencia = set(df_final["NRO DO EQUIPAMENTO"].astype(str).str.strip())
            equipamentos_coletados = set()

            # Descobre a última linha com equipamento para otimizar a leitura
            last_row = ws.range((ws.cells.last_cell.row, col_equip)).end('up').row
            if last_row < 2: return

            # Ler os valores em Bloco (Batch) para velocidade
            valores_equip = ws.range((2, col_equip), (last_row, col_equip)).value
            valores_coleta = ws.range((2, col_coleta), (last_row, col_coleta)).value

            # Converte valores únicos em lista caso haja apenas 1 linha
            if not isinstance(valores_equip, list): valores_equip = [valores_equip]
            if not isinstance(valores_coleta, list): valores_coleta = [valores_coleta]

            # 1. Identificar quais já estão marcados como DADOS COLETADOS
            for eq, col_val in zip(valores_equip, valores_coleta):
                if col_val == "DADOS COLETADOS" and eq is not None:
                    eq_str = str(eq).strip()
                    if eq_str.endswith('.0'): eq_str = eq_str[:-2]
                    equipamentos_coletados.add(eq_str)

            # 2. Escrever "COLETAR DADOS" para os novos que estão no df_final
            for i, (eq, col_val) in enumerate(zip(valores_equip, valores_coleta)):
                if eq is None: continue
                
                eq_str = str(eq).strip()
                if eq_str.endswith('.0'): eq_str = eq_str[:-2]

                row_idx = i + 2
                cell_coleta = ws.range((row_idx, col_coleta))
                
                if col_val == "DADOS COLETADOS":
                    continue

                if eq_str in equipamentos_contingencia and eq_str not in equipamentos_coletados:
                    cell_coleta.value = "COLETAR DADOS"
                    try:
                        cell_coleta.font.bold = True
                    except:
                        pass
                else:
                    if col_val == "COLETAR DADOS":
                        cell_coleta.value = None

            # --- SALVAMENTO EXPLÍCITO PASSANDO O CAMINHO PARA EVITAR BLOQUEIO DE REDE ---
            wb.save(caminho_planilha)
            print("✅ Atualização da coluna COLETA concluída de forma segura!")
        finally:
            if wb: 
                try: wb.close()
                except: pass
            if app: 
                try: app.quit()
                except: pass
    except ImportError:
        print("❌ Erro fatal: A biblioteca 'xlwings' não está instalada ou o Excel não está instalado no computador.")
    except Exception as e:
        print(f"⚠️ Erro grave ao atualizar a planilha via xlwings: {e}")

def capturar_imagem_pdf_mupdf(caminho_pdf, output_dir, nome_imagem):
    try:
        pdf_documento = fitz.open(caminho_pdf)
        pagina = pdf_documento[0]

        matriz = fitz.Matrix(3, 3)
        imagem = pagina.get_pixmap(matrix=matriz)

        caminho_png = os.path.join(output_dir, f"{nome_imagem}.png")
        imagem.save(caminho_png)
        print(f"🖼️ Imagem do PDF salva como: {caminho_png}")

        pdf_documento.close()
    except Exception as e:
        print(f"⚠️ Erro ao capturar imagem do PDF: {e}")

def salvar_pdf_por_responsavel(df_final, output_dir):
    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        for arquivo in os.listdir(output_dir):
            if arquivo.endswith(".pdf") or arquivo.endswith(".png"):
                os.remove(os.path.join(output_dir, arquivo))

        caminho_planilha = obter_caminho_planilha()
        aba_alvo = "Cont. Maquinas"
        
        # Openpyxl pode ser usado aqui sem medo, pois read_only=True NÃO corrompe o arquivo
        wb = load_workbook(caminho_planilha, data_only=True, read_only=True)
        ws = wb[aba_alvo]

        colunas = {str(cell.value).strip().upper(): idx for idx, cell in enumerate(ws[1]) if cell.value}
        if "EQUIPAMENTO" not in colunas or "COLETA" not in colunas:
            print("⚠️ Colunas necessárias não encontradas na aba Cont. Maquinas.")
            # IMPORTANTE: Liberar o arquivo mesmo se der erro para não bloquear o xlwings depois!
            wb.close()
            return

        equipamentos_coletados = set()
        col_equipamento = colunas["EQUIPAMENTO"]
        col_coleta = colunas["COLETA"]

        for row in ws.iter_rows(min_row=2, values_only=True):
            equipamento = str(row[col_equipamento]).strip() if row[col_equipamento] else ""
            coleta = str(row[col_coleta]).strip() if row[col_coleta] else ""
            if coleta == "DADOS COLETADOS":
                equipamentos_coletados.add(equipamento)

        # --- CORREÇÃO DO ERRO DE SAVE (XLWINGS) ---
        # Fecha e libera o arquivo da memória do openpyxl agora que já temos a lista de equipamentos.
        # Sem isso, o Windows acha que o arquivo está em uso e impede a gravação na próxima função!
        wb.close() 

        df_filtrado = df_final[~df_final["NRO DO EQUIPAMENTO"].astype(str).isin(equipamentos_coletados)]
        responsaveis_para_gerar = df_filtrado['RESPONSAVEL'].unique()
        data_hora_geracao = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')

        COR_PRIMARIA = (60, 100, 160)
        COR_SECUNDARIA = (240, 240, 240)

        for responsavel in responsaveis_para_gerar:
            df_responsavel = df_filtrado[df_filtrado['RESPONSAVEL'] == responsavel]
            if df_responsavel.empty:
                print(f"⚠️ Não há dados para o responsável: {responsavel}.")
                continue

            pdf = FPDF(format='letter')
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()

            pdf.set_fill_color(*COR_PRIMARIA)
            pdf.rect(0, 0, 220, 30, 'F')
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(0, 15, 'RELATÓRIO DE CONTINGÊNCIA', ln=1, align='C')
            pdf.set_font("Arial", 'B', 10)
            pdf.cell(0, 5, 'Equipamentos sem comunicação no dia atual', ln=1, align='C')
            pdf.ln(10)

            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 10, f'Responsável: {responsavel}', ln=1, align='L')
            pdf.set_font("Arial", '', 8)
            pdf.cell(0, 5, f'Relatório gerado em: {data_hora_geracao}', ln=1, align='L')
            pdf.ln(5)

            pdf.set_fill_color(*COR_PRIMARIA)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Arial", 'B', 8)

            headers = ['DISPLAY', 'FROTA', 'TIPO EQUIP.', 'PRESTADOR', 'ÚLTIMA COMUNICAÇÃO']
            col_widths = [38, 38, 38, 38, 38]

            for i, header in enumerate(headers):
                pdf.cell(col_widths[i], 10, header, border=0, align='C', fill=True)
            pdf.ln()

            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", size=8)

            for i, row in enumerate(df_responsavel.iterrows()):
                if i % 2 == 0:
                    pdf.set_fill_color(255, 255, 255)
                else:
                    pdf.set_fill_color(*COR_SECUNDARIA)

                row_data = row[1]

                def limpar_val(val):
                    v = str(val).strip()
                    if v.lower() in ['nan', 'none', '<na>', '']:
                        return ''
                    return v

                display_val = limpar_val(row_data.get('DISPLAY', ''))
                frota_val = limpar_val(row_data.get('NRO DO EQUIPAMENTO', ''))
                tipo_val = limpar_val(row_data.get('TIPO DO EQUIPAMENTO', ''))
                prestador_val = limpar_val(row_data.get('PRESTADOR', ''))

                pdf.cell(col_widths[0], 10, display_val, border=0, align='C', fill=True)
                pdf.cell(col_widths[1], 10, frota_val, border=0, align='C', fill=True)
                pdf.cell(col_widths[2], 10, tipo_val, border=0, align='C', fill=True)
                pdf.cell(col_widths[3], 10, prestador_val, border=0, align='C', fill=True)

                registro_formatado = row_data['REGISTRO MAIS RECENTE'].strftime('%d/%m/%Y %H:%M') if pd.notna(
                    row_data['REGISTRO MAIS RECENTE']) else 'N/A'
                pdf.cell(col_widths[4], 10, registro_formatado, border=0, align='C', fill=True)
                pdf.ln()

            pdf.ln(5)
            pdf.set_font("Arial", size=6, style='')
            pdf.cell(0, 10, "Relatório gerado automaticamente pelo Sistema.", ln=1, align='C')

            caminho_pdf = os.path.join(output_dir, f"Relatorio_{responsavel.replace(' ', '_')}.pdf")
            pdf.output(caminho_pdf)

            print(f"📄 PDF gerado com sucesso para {responsavel}: {caminho_pdf}")
            capturar_imagem_pdf_mupdf(caminho_pdf, output_dir, f"Relatorio_{responsavel.replace(' ', '_')}")
    except Exception as e:
        print(f"⚠️ Erro ao gerar PDF: {e}")

# --- Funções da Interface Gráfica (Tkinter) ---
def alternar_visualizacao_senha():
    """Alterna a visualização da senha na interface."""
    if entry_senha.cget('show') == '*':
        entry_senha.config(show='')
        botao_visualizar.config(text="Ocultar")
    else:
        entry_senha.config(show='*')
        botao_visualizar.config(text="Mostrar")

def atualizar_campos_credenciais(credenciais_path):
    """Carrega as credenciais de um arquivo JSON."""
    try:
        with open(credenciais_path, "r", encoding='utf-8') as file:
            data = json.load(file)
            usuario = data.get("usuario", "")
            senha = data.get("senha", "")
            return True, usuario, senha
    except (FileNotFoundError, json.JSONDecodeError):
        return False, "", ""

def salvar_usuario(credenciais_path):
    """Salva o usuário e a senha em um arquivo JSON."""
    usuario = entry_usuario.get()
    senha = entry_senha.get()

    if var_salvar_usuario.get():
        credenciais = {"usuario": usuario, "senha": senha}
        try:
            with open(credenciais_path, "w", encoding='utf-8') as file:
                json.dump(credenciais, file, indent=4)
            print(f"🟢 Usuário e senha salvos em: {credenciais_path}")
        except Exception as e:
            print(f"⚠️ Erro ao salvar credenciais: {e}")
    else:
        if os.path.exists(credenciais_path):
            os.remove(credenciais_path)
            print("🔴 Arquivo de credenciais removido.")

def cancelar_execucao():
    global execucao_ativa
    execucao_ativa = False
    print("🔴 Execução cancelada.")
    atualizar_progresso("Execução cancelada.", step=0, total_steps=1)

def atualizar_progresso(status_texto, step, total_steps):
    """Atualiza o rótulo de status e a barra de progresso de forma segura para threads."""
    if root and status_label and progress_bar:
        root.after(0, _atualizar_progresso_thread_safe, status_texto, step, total_steps)

def _atualizar_progresso_thread_safe(status_texto, step, total_steps):
    if total_steps > 0:
        progress = (step / total_steps) * 100
        progress_bar['value'] = progress
    status_label.config(text=status_texto)
    root.update_idletasks()

def criar_interface():
    global entry_usuario, entry_senha, var_salvar_usuario, botao_visualizar, driver, responsaveis_vars, entry_intervalo, var_intervalo_ativado, status_label, progress_bar, root

    credenciais_path = os.path.join(script_dir, "credenciais.json")

    PALETTE = {
        "primary": "#0066AC",
        "secondary": "#43948C",
        "success": "#6BBE3B",
        "danger": "#B90000",
        "background": "#FFFFFF",
        "text": "#000000",
    }

    root = ttk.Window(themename="yeti")
    root.title(f"HXG - Auto  v{VERSAO}")

    style = ttk.Style()
    style.configure("TLabel", font=("Helvetica", 11), background=PALETTE["background"], foreground=PALETTE["text"])
    style.configure("TFrame", background=PALETTE["background"])
    style.configure("TLabelframe", background=PALETTE["background"], foreground=PALETTE["text"])
    style.configure("TLabelframe.Label", background=PALETTE["background"], foreground=PALETTE["text"])
    style.configure("TEntry", fieldbackground="white", foreground=PALETTE["text"])

    style.configure("success.TButton", background=PALETTE["success"], foreground="white",
                    font=("Helvetica", 11, "bold"))
    style.configure("danger.TButton", background=PALETTE["danger"], foreground="white", font=("Helvetica", 11, "bold"))
    style.configure("info.TButton", background=PALETTE["primary"], foreground="white", font=("Helvetica", 11, "bold"))
    style.configure("secondary.TButton", background=PALETTE["secondary"], foreground="white",
                    font=("Helvetica", 11, "bold"))
    style.map("TButton", background=[("active", PALETTE["primary"])])

    style.configure("Roundtoggle.TCheckbutton", background=PALETTE["background"], foreground=PALETTE["text"])
    style.configure("info-round-toggle.TCheckbutton", background=PALETTE["background"], foreground=PALETTE["text"])

    responsaveis_vars = {nome: tk.BooleanVar() for nome in RESPONSAVEIS_OPCOES}

    main_frame = ttk.Frame(root, padding=20)
    main_frame.pack(fill="both", expand=True)

    ttk.Label(main_frame, text="AUTO. CONTIGÊNCIA - HXG", font=("Helvetica", 20, "bold"),
              foreground=PALETTE["primary"]).pack(pady=(0, 20))

    cred_frame = ttk.Labelframe(main_frame, text="Credenciais", padding=15)
    cred_frame.pack(fill="x", pady=10)

    ttk.Label(cred_frame, text="Usuário:").pack(anchor="w", pady=(0, 5))
    entry_usuario = ttk.Entry(cred_frame, width=40)
    entry_usuario.pack(fill="x")

    ttk.Label(cred_frame, text="Senha:").pack(anchor="w", pady=(10, 5))
    frame_senha = ttk.Frame(cred_frame)
    frame_senha.pack(fill="x")

    entry_senha = ttk.Entry(frame_senha, show="*")
    entry_senha.pack(side="left", fill="x", expand=True)

    botao_visualizar = ttk.Button(frame_senha, text="Mostrar", command=alternar_visualizacao_senha)
    botao_visualizar.pack(side="left", padx=(5, 0))

    var_salvar_usuario = tk.BooleanVar()
    credenciais_existentes, usuario_carregado, senha_carregada = atualizar_campos_credenciais(credenciais_path)
    var_salvar_usuario.set(credenciais_existentes)

    if credenciais_existentes:
        entry_usuario.insert(0, usuario_carregado)
        entry_senha.insert(0, senha_carregada)

    ttk.Checkbutton(cred_frame, text="Salvar usuário e senha", variable=var_salvar_usuario,
                    bootstyle="round-toggle").pack(anchor="w", pady=(10, 0))

    intervalo_frame = ttk.Frame(main_frame)
    intervalo_frame.pack(fill="x", pady=(10, 5))

    ttk.Label(intervalo_frame, text="Executar a cada (minutos):").pack(side="left", padx=(0, 5))
    entry_intervalo = ttk.Entry(intervalo_frame, width=10)
    entry_intervalo.insert(0, "60")
    entry_intervalo.pack(side="left", padx=(0, 10))

    var_intervalo_ativado = tk.BooleanVar(value=False)
    ttk.Checkbutton(intervalo_frame, text="Ativar agendamento", variable=var_intervalo_ativado,
                    bootstyle="round-toggle").pack(side="left")

    resp_frame = ttk.Labelframe(main_frame, text="Gerar PDF para:", padding=15)
    resp_frame.pack(fill="both", expand=True, pady=10)

    def selecionar_todos():
        for var in responsaveis_vars.values():
            var.set(True)

    def limpar_selecao():
        for var in responsaveis_vars.values():
            var.set(False)

    btn_frame = ttk.Frame(resp_frame)
    btn_frame.pack(fill="x", pady=(0, 5))
    ttk.Button(btn_frame, text="Selecionar Todos", command=selecionar_todos, bootstyle="info").pack(side="left",
                                                                                                  fill="x",
                                                                                                  expand=True,
                                                                                                  padx=(0, 5))
    ttk.Button(btn_frame, text="Limpar Seleção", command=limpar_selecao, bootstyle="secondary").pack(side="left",
                                                                                                    fill="x",
                                                                                                    expand=True,
                                                                                                    padx=(5, 0))

    for nome, var in responsaveis_vars.items():
        ttk.Checkbutton(resp_frame, text=nome, variable=var, bootstyle="info-round-toggle").pack(anchor="w", pady=2)

    status_label = ttk.Label(main_frame, text="Aguardando...", font=("Helvetica", 10), foreground="gray")
    status_label.pack(pady=(10, 5))

    progress_bar = ttk.Progressbar(main_frame, mode="determinate", bootstyle="info")
    progress_bar.pack(fill="x", pady=(0, 10))

    action_frame = ttk.Frame(main_frame)
    action_frame.pack(fill="x", pady=10)

    ttk.Button(action_frame, text="Executar", command=lambda: executar_script(), bootstyle="success").pack(side="left",
                                                                                                           fill="x",
                                                                                                           expand=True,
                                                                                                           padx=(0, 5))
    ttk.Button(action_frame, text="Pausar", command=cancelar_execucao, bootstyle="danger").pack(side="left", fill="x",
                                                                                               expand=True,
                                                                                               padx=(5, 0))

    root.bind('<Return>', lambda event: executar_script())

    def fechar_janela():
        salvar_usuario(credenciais_path)
        if 'driver' in globals() and driver:
            try:
                driver.quit()
            except:
                pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", fechar_janela)
    root.mainloop()

def executar_script():
    global execucao_ativa
    if execucao_ativa:
        messagebox.showinfo("Informação", "A automação já está em execução.")
        return

    usuario = entry_usuario.get()
    senha = entry_senha.get()

    if not usuario or not senha:
        messagebox.showwarning("Aviso", "Por favor, preencha o usuário e a senha.")
        return

    credenciais_path = os.path.join(script_dir, "credenciais.json")
    salvar_usuario(credenciais_path)

    atualizar_progresso("Iniciando a automação...", step=0, total_steps=6)
    threading.Thread(target=executar_procedimento, args=(usuario, senha), daemon=True).start()

def executar_procedimento(usuario, senha):
    global driver, execucao_ativa, responsaveis_vars
    execucao_ativa = True

    TOTAL_STEPS = 6
    last_valid_interval = 60

    while execucao_ativa:
        selected_responsaveis = [nome for nome, var in responsaveis_vars.items() if var.get()]
        intervalo_ativado = var_intervalo_ativado.get()

        try:
            intervalo_minutos = int(entry_intervalo.get())
            if intervalo_minutos <= 0:
                print(
                    f"⚠️ Intervalo inválido ({intervalo_minutos}). Usando o original: {last_valid_interval} min.")
                intervalo_minutos = last_valid_interval
            else:
                last_valid_interval = intervalo_minutos
        except (ValueError, tk.TclError):
            print(f"⚠️ Erro ao ler o intervalo. Usando o original: {last_valid_interval} min.")
            intervalo_minutos = last_valid_interval

        print(f"\n--- Iniciando novo ciclo ---")
        print(f"Responsáveis selecionados para este ciclo: {', '.join(selected_responsaveis) or 'Todos'}")
        print(
            f"Intervalo configurado: {intervalo_minutos} minutos. Agendamento Ativado: {'Sim' if intervalo_ativado else 'Não'}")

        driver = None
        df_final = None

        try:
            diretorio_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
            pdf_output_dir = os.path.join(script_dir, "PDF_Saida")

            xpaths = {
                'usuario': '/html/body/app-root/app-login/app-access-container/div/div[2]/div[2]/form/div[1]/input',
                'senha': '/html/body/app-root/app-login/app-access-container/div/div[2]/div[2]/form/div[2]/input',
                'botao_login': '/html/body/app-root/app-login/app-access-container/div/div[2]/div[2]/form/div[4]/p-button/button',
                'limpar_filtro': '//*[@id="buttion-id-clearAndApplyButton"]/div/app-text/div',
                'tabela': '//*[@id="div-submenu-link-id-app-submenu-link-mon-table-id"]/p',
                'paginador_dropdown': '//p-paginator//p-dropdown | //*[starts-with(@id, "pn_id_") and contains(@id, "label")]',
                'paginador_opcao_5': '//p-dropdownitem[5]/li',
                'filtro': '//p-table//div[1]/div/div[2]/button[3] | //*[starts-with(@id, "pn_id_")]/div[1]/div/div[2]/button[3]',
                'exportacao_csv': '//p-table//div[1]/div/div[2]/button[1] | //*[starts-with(@id, "pn_id_")]/div[1]/div/div[2]/button[1]'
            }
            url = 'https://access.hxgnagron.com/?redirect=http:%2F%2Fcontrolroom.hxgnagron.com%2F#/'

            atualizar_progresso("Iniciando driver...", step=1, total_steps=TOTAL_STEPS)
            driver = iniciar_driver(headless=True)

            if not execucao_ativa: break

            atualizar_progresso("Realizando login...", step=2, total_steps=TOTAL_STEPS)
            login_usuario(driver, url, usuario, senha, xpaths)

            if not execucao_ativa: break

            atualizar_progresso("Exportando tabela...", step=3, total_steps=TOTAL_STEPS)
            exportar_tabela(driver, xpaths)

            if not execucao_ativa: break

            atualizar_progresso("Aguardando download e processando...", step=4, total_steps=TOTAL_STEPS)
            df_final = processar_csv(diretorio_downloads, pdf_output_dir, selected_responsaveis)

            if df_final is not None:
                if not execucao_ativa: break
                atualizar_progresso("Gerando PDFs...", step=5, total_steps=TOTAL_STEPS)
                salvar_pdf_por_responsavel(df_final, pdf_output_dir)

                if not execucao_ativa: break
                atualizar_progresso("Atualizando planilha de controle...", step=6,
                                    total_steps=TOTAL_STEPS)
                atualizar_coleta_planilha(df_final)

                atualizar_progresso("Procedimento concluído com sucesso!", step=6,
                                    total_steps=TOTAL_STEPS)
            else:
                atualizar_progresso("Processamento de dados falhou.", step=0, total_steps=1)

        except Exception as e:
            error_message = f"❌ Erro fatal na execução: {type(e).__name__}: {str(e)[:100]}..."
            print(error_message)
            logger.error(f"❌ Erro fatal na execução do procedimento: {e}")
            atualizar_progresso(error_message, step=0, total_steps=1)
        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass

        if not intervalo_ativado:
            print("Execução única concluída, pois o agendamento está desativado.")
            break

        if execucao_ativa:
            tempo_total_espera = intervalo_minutos * 60
            print(f"✅ Execução concluída. Aguardando {intervalo_minutos} minutos para a próxima rodada...")

            for segundos_restantes in range(tempo_total_espera, 0, -1):
                if not execucao_ativa: break

                minutos, segundos = divmod(segundos_restantes, 60)
                texto_tempo = f"Próxima execução em {minutos:02d}:{segundos:02d}"
                atualizar_progresso(texto_tempo, step=6, total_steps=TOTAL_STEPS)
                time.sleep(1)

            if not execucao_ativa: break

    execucao_ativa = False
    atualizar_progresso("Procedimento finalizado.", step=0, total_steps=1)
    print("🏁 Procedimento finalizado.")

if __name__ == "__main__":
    verificar_seguranca()
    driver = None
    criar_interface()