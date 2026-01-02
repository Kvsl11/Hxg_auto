import threading
import glob
import pandas as pd
# --- MODIFICAÇÕES DE IMPORTAÇÃO ---
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager (REMOVIDO - Usaremos o SeleniumManager embutido)
# --- FIM MODIFICAÇÕES ---
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
# --- CORREÇÃO DE CAMINHO ---
# Usa o executável Python que está rodando o script ATUALMENTE
python_exe = sys.executable 
print(f"🟢 Usando Python em: {python_exe}")

# A lógica de relançar com o Python interno foi removida,
# pois estava causando a falha na atualização de dependências.


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
        # Remove o cache do webdriver-manager para evitar que o Selenium o use por engano.
        wdm_cache_path = os.path.join(os.path.expanduser("~"), ".wdm")
        if os.path.exists(wdm_cache_path):
            logger.info(f"🧹 Limpando cache antigo do webdriver-manager em: {wdm_cache_path}")
            shutil.rmtree(wdm_cache_path, ignore_errors=True)
        # --- FIM DA LIMPEZA ---

        # Remove 'webdriver-manager' da lista, pois não é mais necessário
        pacotes = ["certifi", "selenium"] 
        logger.info(f"🔍 Verificando e atualizando pacotes: {', '.join(pacotes)}...")
        
        for pacote in pacotes:
            logger.info(f"Instalando/Atualizando {pacote}...")
            # Esta chamada agora usará o 'sys.executable' correto
            subprocess.run([python_exe, "-m", "pip", "install", "--upgrade", pacote], check=True, capture_output=True, text=True)
        
        import certifi
        logger.info(f"🟢 Dependências atualizadas com sucesso. Caminho Certifi: {certifi.where()}")
    except Exception as e:
        logger.warning(f"⚠️ Falha ao preparar dependências: {e}")

def garantir_certificados_amazon():
# ... (código existente sem alterações) ...
    """Verifica se o certificado raiz da Amazon está presente e adiciona se necessário."""
    try:
        import certifi
        cacert_path = certifi.where()

        with open(cacert_path, "r", encoding="utf-8") as f:
            conteudo = f.read()

        for nome, url in AMAZON_CERTS.items():
            if nome not in conteudo:
                logger.info(f"🔍 {nome} não encontrado, baixando de {url}...")
                # Usamos verify=False aqui se o teste SSL global ainda não foi feito e tiver problemas
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

# --- FUNÇÃO DE LIMPEZA DE CACHE (REMOVIDA, NÃO É MAIS NECESSÁRIA) ---
# ... (código existente sem alterações) ...
# def limpar_cache_uc(): ...


# --- EXECUÇÃO AUTOMÁTICA AO INICIAR ---
logger.info("🚀 Iniciando verificação e correção SSL híbrida...")
# ... (código existente sem alterações) ...
preparar_dependencias() # Renomeado de atualizar_certifi()
garantir_certificados_amazon()

# --- NOVO BLOCO DE TESTE E FALLBACK SSL GLOBAL ---
# ... (código existente sem alterações) ...
# Tenta o modo verificado. Se falhar com SSL Error, seta o modo não verificado
# de forma global e persistente para evitar o erro CERTIFICATE_VERIFY_FAILED
# em chamadas subsequentes (UC, requests, etc.).
try:
    import certifi
    ssl_context = ssl.create_default_context(cafile=certifi.where())
    # Testa uma URL genérica para validar o contexto atual
    urllib.request.urlopen("https://www.google.com", timeout=5, context=ssl_context)
    logger.info("🟢 Conexão SSL validada com sucesso — certificados OK.")
except ssl.SSLError as e:
    # Se falhar com erro SSL (como CERTIFICATE_VERIFY_FAILED), aplica o fallback.
    logger.warning(f"⚠️ Falha de SSL detectada ({e}). Aplicando modo não verificado (Fallback).")
    ssl._create_default_https_context = ssl._create_unverified_context
    logger.info("🟡 SSL desativado globalmente — conexão forçada sem verificação de certificado.")
except Exception as e:
    # Falha genérica, ainda assim aplica o fallback preventivamente.
    logger.warning(f"⚠️ Erro genérico ao testar SSL: {e}. Aplicando modo não verificado (Fallback).")
    ssl._create_default_https_context = ssl._create_unverified_context
    logger.info("🟡 SSL desativado globalmente — conexão forçada sem verificação de certificado.")

logger.info("✅ Configuração SSL concluída com segurança.")

# --- VERIFICAÇÃO DE ATUALIZAÇÃO VIA GITHUB ---
# ... (código existente sem alterações) ...
VERSAO = "3.1.5"

def verificar_e_atualizar_automaticamente():
# ... (código existente sem alterações) ...
    """
    Verifica no GitHub se há nova versão e atualiza automaticamente sem interação do usuário.
    """
    try:
        REPO = "Kvsl11/Hxg_auto"
        URL_VERSION = f"https://raw.githubusercontent.com/{REPO}/main/version.txt"
        URL_SCRIPT = f"https://raw.githubusercontent.com/{REPO}/main/main.py"
        LOCAL_SCRIPT = os.path.join(os.path.dirname(__file__), "main.py")
        LOCAL_VERSION_FILE = os.path.join(os.path.dirname(__file__), "version_local.txt")
        LOG_PATH = os.path.join(os.path.dirname(__file__), "autoupdate.log")

        # Mantém a configuração de log para o arquivo de autoupdate, pois é uma thread separada.
        logging.basicConfig(
            filename=LOG_PATH,
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )

        def get_local_version():
# ... (código existente sem alterações) ...
            if os.path.exists(LOCAL_VERSION_FILE):
                try:
                    with open(LOCAL_VERSION_FILE, "r", encoding="utf-8") as f:
                        return f.read().strip()
                except Exception:
                    return "0.0.0"
            return "0.0.0"

        def get_online_version():
# ... (código existente sem alterações) ...
            try:
                headers = {"Cache-Control": "no-cache", "Pragma": "no-cache"}
                # Mantém verify=False aqui para garantir que a atualização funcione mesmo com problemas SSL
                r = requests.get(URL_VERSION, timeout=10, verify=False, headers=headers)
                if r.status_code == 200:
                    return r.text.strip()
                else:
                    logging.warning(f"⚠️ Falha HTTP ao buscar versão: {r.status_code}")
            except Exception as e:
                logging.warning(f"⚠️ Falha ao obter versão online: {e}")
            return None

        def save_local_version(ver):
# ... (código existente sem alterações) ...
            try:
                with open(LOCAL_VERSION_FILE, "w", encoding="utf-8") as f:
                    f.write(ver)
                logging.info(f"✅ Versão local atualizada para {ver}")
            except Exception as e:
                logging.error(f"❌ Erro ao salvar versão local: {e}")

        def atualizar_script(versao_online):
# ... (código existente sem alterações) ...
            try:
                headers = {"Cache-Control": "no-cache", "Pragma": "no-cache"}
                # Mantém verify=False aqui para garantir que a atualização funcione mesmo com problemas SSL
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
# ... (código existente sem alterações) ...
        online_v = get_online_version()

        if not online_v:
# ... (código existente sem alterações) ...
            logging.warning("⚠️ Falha ao verificar versão online. Continuando com a versão local.")
            return

        if online_v != local_v:
# ... (código existente sem alterações) ...
            logging.info(f"🟡 Nova versão detectada: {online_v} (local: {local_v}) — atualizando...")
            sucesso = atualizar_script(online_v)
            if sucesso:
# ... (código existente sem alterações) ...
                logging.info("♻️ Reiniciando app com nova versão...")
                python_exe = sys.executable
                subprocess.Popen([python_exe, LOCAL_SCRIPT])
                os._exit(0)
            else:
# ... (código existente sem alterações) ...
                logging.info(f"🟢 Aplicativo já está atualizado ({local_v})")

    except Exception as e:
# ... (código existente sem alterações) ...
        logging.error(f"❌ Erro na verificação automática de atualização: {e}")

warnings.filterwarnings(
# ... (código existente sem alterações) ...
    "ignore",
    message="Slicer List extension is not supported and will be removed",
    category=UserWarning,
    module=r"openpyxl\.worksheet\._reader"
)

# --- Definições Globais ---
# ... (código existente sem alterações) ...
# Descobre dinamicamente o diretório onde o script está localizado
script_dir = os.path.dirname(os.path.abspath(__file__))
execucao_ativa = False
# Declarações globais para os widgets da barra de progresso e status
status_label = None
progress_bar = None
root = None

# Lista de responsáveis para o seletor da interface
RESPONSAVEIS_OPCOES = [
# ... (código existente sem alterações) ...
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

# --- Funções de Automação com Selenium (usando undetected_chromedriver) ---
def iniciar_driver(headless=True):
# ... (código existente sem alterações) ...
    """
    Inicia uma instância do Chrome usando Selenium padrão e o SeleniumManager embutido,
    com opções para ignorar erros de SSL no navegador.
    """
    print("🚀 Iniciando driver com Selenium padrão (SeleniumManager)...")
    logger.info("🚀 Iniciando driver com Selenium padrão (SeleniumManager)...")

    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        # --- CORREÇÃO HEADLESS 1: Definir tamanho da janela ---
        # Evita que elementos responsivos (menus) cubram os botões.
        options.add_argument("--window-size=1920,1080")
    
    # --- CORREÇÃO SSL NÍVEL NAVEGADOR ---
# ... (código existente sem alterações) ...
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-running-insecure-content')
    options.add_argument('--disable-web-security') # Adicional
    
    # Desativa logs desnecessários do Selenium
# ... (código existente sem alterações) ...
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    try:
# ... (código existente sem alterações) ...
        # --- MUDANÇA PRINCIPAL ---
        # Ao chamar Service() vazio, o Selenium (4.6.0+) usa seu 
        # SeleniumManager embutido para baixar o driver correto.
        servico = Service() 
        
        driver = webdriver.Chrome(service=servico, options=options)
        # --- FIM DA MUDANÇA ---
        
    except Exception as e:
# ... (código existente sem alterações) ...
        # Se o SeleniumManager falhar (mesmo com patch SSL),
        # é um erro de rede/firewall ou permissão.
        print(f"❌ Falha ao iniciar Selenium/SeleniumManager: {e}")
        logger.error(f"❌ Falha ao iniciar Selenium/SeleniumManager: {e}")
        raise # Levanta a exceção para ser tratada em 'executar_procedimento'

    if not headless:
# ... (código existente sem alterações) ...
        driver.maximize_window()
    return driver

    
def aguardar_pagina_carregada(driver, timeout=30):
# ... (código existente sem alterações) ...
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        print("🟢 Página totalmente carregada.")
    except Exception as e:
        print(f"⚠️ Erro ao aguardar carregamento: {e}")

def aguardar_e_clicar(driver, xpath, timeout=30):
# ... (código existente sem alterações) ...
    try:
        print(f"Tentando clicar em: {xpath}")
        elemento = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento)
        
        # AGUARDA A CLICABILIDADE
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        
        # --- CORREÇÃO HEADLESS 2: PEQUENA PAUSA ---
        # Adiciona uma pequena pausa para o DOM 'assentar' em modo headless
        # antes de tentar o clique, ajudando a evitar "interceptações".
        time.sleep(0.5) 
        
        elemento.click()
        print(f"🟢 Clique realizado: {xpath}")
        logger.info(f"🟢 Clique realizado: {xpath}")
    except Exception as e:
# ... (código existente sem alterações) ...
        print(f"⚠️ Clique padrão falhou ({xpath}). Tentando via JavaScript... Erro: {e}")
        logger.warning(f"⚠️ Clique padrão falhou ({xpath}). Tentando via JS.")
        try:
# ... (código existente sem alterações) ...
            elemento = driver.find_element(By.XPATH, xpath)
            driver.execute_script("arguments[0].click();", elemento)
            print(f"🟢 Clique via JS realizado: {xpath}")
            logger.info(f"🟢 Clique via JS realizado: {xpath}")
        except Exception as js_e:
# ... (código existente sem alterações) ...
            print(f"❌ Erro final ao clicar via JS em {xpath}: {js_e}")
            logger.error(f"❌ Erro final ao clicar via JS em {xpath}: {js_e}")

def aguardar_e_escrever(driver, xpath, texto, timeout=30):
# ... (código existente sem alterações) ...
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
# ... (código existente sem alterações) ...
    driver.get(url)
    aguardar_pagina_carregada(driver)
    aguardar_e_escrever(driver, xpaths['usuario'], usuario)
    aguardar_e_escrever(driver, xpaths['senha'], senha)
    aguardar_e_clicar(driver, xpaths['botao_login'])
    time.sleep(3) # Manter este sleep para a página de login processar.

def exportar_tabela(driver, xpaths):
    aguardar_e_clicar(driver, xpaths['control_room'])
    
    limpar_filtro_xpath = '//button[contains(@id,"buttion-id-clearAndApplyButton")]'
    aguardar_e_clicar(driver, limpar_filtro_xpath)
    
    # --- CORREÇÃO HEADLESS 3: ESPERAR OVERLAY SUMIR ---
    # O log mostrou que um <strong> (provavelmente "Loading...") interceptou
    # o clique em "Tabela". Vamos esperar ele desaparecer após limpar o filtro.
    print("⏳ Aguardando overlay de filtro desaparecer...")
    try:
        # Aumentei o timeout para 20s caso a rede esteja lenta
        WebDriverWait(driver, 20).until(
            EC.invisibility_of_element_located((By.XPATH, "//strong[contains(.,'Loading')] | //strong[contains(.,'Carregando')] | //div[contains(@class, 'overlay')]"))
        )
        print("🟢 Overlay desapareceu.")
    except Exception as e:
        # Se não encontrar o overlay (ou ele sumir rápido), apenas avisa e continua.
        print(f"⚠️ Não foi possível confirmar o desaparecimento do overlay (ou não havia): {e}")
    # --- FIM DA CORREÇÃO ---
    
    aguardar_e_clicar(driver, xpaths['tabela'])
    aguardar_e_clicar(driver, xpaths['filtro'])
    time.sleep(2)
    aguardar_e_clicar(driver, xpaths['exportacao_csv'])
    print("🟢 Exportação iniciada")

def aguardar_download_completo(diretorio, nome_base, timeout=60):
# ... (código existente sem alterações) ...
    tempo_inicial = time.time()
    while time.time() - tempo_inicial < timeout:
        arquivos_tmp = glob.glob(os.path.join(diretorio, f"{nome_base}*.tmp"))
        arquivos_csv = glob.glob(os.path.join(diretorio, f"{nome_base}*.csv"))

        if arquivos_csv and not arquivos_tmp:
            return max(arquivos_csv, key=os.path.getctime)
        time.sleep(2)

    print("❌ Tempo limite excedido para o download do arquivo CSV!")
    return None

# --- Funções de Processamento de Dados (Pandas, OpenPyXL) ---
# ... (código existente sem alterações) ...

def processar_csv(diretorio_downloads, excel_output, base_responsaveis_path, pdf_output_dir, selected_responsaveis):
# ... (código existente sem alterações) ...
    try:
        if not os.path.exists(pdf_output_dir):
            os.makedirs(pdf_output_dir)
        
        print("⏳ Aguardando download do arquivo CSV...")
        csv_path = aguardar_download_completo(diretorio_downloads, "Monitoramento - Tabela")
        if not csv_path:
            print("❌ Processo encerrado. Nenhum arquivo CSV disponível.")
            return None
        
        df = pd.read_csv(csv_path, encoding="utf-8", sep=";", dtype=str)
# ... (código existente sem alterações) ...
        df["Registro mais recente"] = pd.to_datetime(df["Registro mais recente"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
        
        data_atual = datetime.datetime.now().date()
# ... (código existente sem alterações) ...
        df_antigos = df[df["Registro mais recente"].dt.date != data_atual]
        df_responsaveis = pd.read_excel(base_responsaveis_path, dtype=str)

        col_equipamento = "Nro do equipamento"
# ... (código existente sem alterações) ...
        col_responsavel = "RESPONSAVEL"
        col_display = "DISPLAY"
        col_prestador = "PRESTADOR"
        
        if col_equipamento not in df_antigos.columns or col_equipamento not in df_responsaveis.columns:
# ... (código existente sem alterações) ...
            print(f"⚠️ A coluna '{col_equipamento}' não foi encontrada em uma das planilhas!")
            return None
        
        df_final = df_antigos.merge(df_responsaveis[[col_equipamento, col_responsavel, col_display, col_prestador]], on=col_equipamento, how="left")
# ... (código existente sem alterações) ...
        df_final = df_final.dropna(subset=[col_responsavel])
        
        # Filtra por responsáveis selecionados se a lista não estiver vazia
# ... (código existente sem alterações) ...
        if selected_responsaveis:
            print(f"✅ Gerando relatórios apenas para: {', '.join(selected_responsaveis)}")
            df_final = df_final[df_final[col_responsavel].isin(selected_responsaveis)]
        else:
# ... (código existente sem alterações) ...
            print("🔄 Nenhum responsável selecionado. Gerando relatórios para todos os responsáveis.")

        colunas_desejadas = [
# ... (código existente sem alterações) ...
            "RESPONSAVEL", "DISPLAY", "Nro do equipamento",
            "Tipo do equipamento", "PRESTADOR", "Registro mais recente"
        ]
        df_final = df_final[colunas_desejadas]
        
        try:
# ... (código existente sem alterações) ...
            with pd.ExcelWriter(excel_output, engine='openpyxl', mode='w') as writer:
                df_final.to_excel(writer, index=False, sheet_name="Contingência")
            print(f"🟢 Registros antigos com responsáveis salvos em: {excel_output}")
        except Exception as save_error:
# ... (código existente sem alterações) ...
            print(f"⚠️ Erro ao salvar Excel: {save_error}")
            return None

        return df_final
# ... (código existente sem alterações) ...
    except Exception as e:
        print(f"⚠️ Erro ao processar CSV: {e}")
        return None

def obter_caminho_planilha():
# ... (código existente sem alterações) ...
    import os, datetime
    
    ano_atual = datetime.datetime.now().year
# ... (código existente sem alterações) ...
    mes_atual = datetime.datetime.now().month

    meses = {
# ... (código existente sem alterações) ...
        1: "01 - Janeiro", 2: "02 - Fevereiro", 3: "03 - Março",
        4: "04 - Abril", 5: "05 - Maio", 6: "06 - Junho",
        7: "07 - Julho", 8: "08 - Agosto", 9: "09 - Setembro",
        10: "10 - Outubro", 11: "11 - Novembro", 12: "12 - Dezembro"
    }

    numero_safra = 2.5 + (ano_atual - 2025) * 0.1
# ... (código existente sem alterações) ...
    safra = f"{numero_safra:.1f} - Safra {ano_atual}"

    possiveis_drives = ["I:", "Z:"]

    caminho_final = None
# ... (código existente sem alterações) ...
    for drive in possiveis_drives:
        base = fr"{drive}\ANG\Agricola\Controle\Computador de Bordo\Fechamento Prestação de Serviço (Linha Amarela)\Pago pelo Bordo"
        caminho_teste = os.path.join(base, safra, meses[mes_atual], "Monitoramento - Eqps.xlsx")
        if os.path.exists(caminho_teste):
# ... (código existente sem alterações) ...
            caminho_final = caminho_teste
            break

    if not caminho_final:
# ... (código existente sem alterações) ...
        raise FileNotFoundError("❌ Não foi possível localizar a planilha em nenhum dos caminhos (I: ou Z:).")

    return caminho_final

def atualizar_coleta_planilha(df_final):
# ... (código existente sem alterações) ...
    try:
        caminho_planilha = obter_caminho_planilha()

        aba_alvo = "Cont. Maquinas"
        
        wb = load_workbook(caminho_planilha)
        ws = wb[aba_alvo]
        
        col_equipamento_final = "Nro do equipamento"
# ... (código existente sem alterações) ...
        cabecalhos = {cell.value.strip().upper(): cell.column for cell in ws[1] if cell.value}

        if "EQUIPAMENTO" not in cabecalhos or "COLETA" not in cabecalhos:
# ... (código existente sem alterações) ...
            print("⚠️ Colunas necessárias não encontradas na aba Cont. Maquinas.")
            return
        
        col_equipamento_contingencia = cabecalhos["EQUIPAMENTO"]
# ... (código existente sem alterações) ...
        col_coleta = cabecalhos["COLETA"]
        
        equipamentos_contingencia = set(df_final[col_equipamento_final].astype(str).str.strip())
# ... (código existente sem alterações) ...
        equipamentos_coletados = set()
        
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=False):
# ... (código existente sem alterações) ...
            equipamento_cell = row[col_equipamento_contingencia - 1]
            coleta_cell = row[col_coleta - 1]
            if coleta_cell.value == "DADOS COLETADOS" and equipamento_cell.value:
# ... (código existente sem alterações) ...
                equipamentos_coletados.add(str(equipamento_cell.value).strip())

        bold_font = Font(bold=True)
        
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=False):
# ... (código existente sem alterações) ...
            equipamento_cell = row[col_equipamento_contingencia - 1]
            cell = row[col_coleta - 1]
            
            if cell.value == "DADOS COLETADOS":
# ... (código existente sem alterações) ...
                continue
            
            if (equipamento_cell.value and 
                str(equipamento_cell.value).strip() in equipamentos_contingencia and 
                str(equipamento_cell.value).strip() not in equipamentos_coletados):
# ... (código existente sem alterações) ...
                cell.value = "COLETAR DADOS"
                cell.font = bold_font
            else:
# ... (código existente sem alterações) ...
                cell.value = None
        
        wb.save(caminho_planilha)
# ... (código existente sem alterações) ...
        print("✅ Atualização da coluna COLETA concluída com sucesso.")

    except Exception as e:
# ... (código existente sem alterações) ...
        print(f"⚠️ Erro ao atualizar a planilha de Equipamentos: {e}") 

# --- Funções de Geração de Arquivos (PDF, Imagem) ---
def capturar_imagem_pdf_mupdf(caminho_pdf, output_dir, nome_imagem):
# ... (código existente sem alterações) ...
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
# ... (código existente sem alterações) ...
    try:
        # Garante que o diretório de saída existe
        if not os.path.exists(output_dir):
# ... (código existente sem alterações) ...
            os.makedirs(output_dir)

        # Remove todos os arquivos PDF existentes no diretório antes de gerar novos
        for arquivo in os.listdir(output_dir):
# ... (código existente sem alterações) ...
            if arquivo.endswith(".pdf") or arquivo.endswith(".png"):
                os.remove(os.path.join(output_dir, arquivo))

        caminho_planilha = obter_caminho_planilha()
# ... (código existente sem alterações) ...
        aba_alvo = "Cont. Maquinas"
        wb = load_workbook(caminho_planilha, data_only=True, read_only=True)
        ws = wb[aba_alvo]

        colunas = {str(cell.value).strip().upper(): idx for idx, cell in enumerate(ws[1]) if cell.value}
        if "EQUIPAMENTO" not in colunas or "COLETA" not in colunas:
# ... (código existente sem alterações) ...
            print("⚠️ Colunas necessárias não encontradas na aba Cont. Maquinas.")
            return

        equipamentos_coletados = set()
# ... (código existente sem alterações) ...
        col_equipamento = colunas["EQUIPAMENTO"]
        col_coleta = colunas["COLETA"]

        for row in ws.iter_rows(min_row=2, values_only=True):
# ... (código existente sem alterações) ...
            equipamento = str(row[col_equipamento]).strip() if row[col_equipamento] else ""
            coleta = str(row[col_coleta]).strip() if row[col_coleta] else ""
            if coleta == "DADOS COLETADOS":
# ... (código existente sem alterações) ...
                equipamentos_coletados.add(equipamento)

        df_filtrado = df_final[~df_final["Nro do equipamento"].astype(str).isin(equipamentos_coletados)]
# ... (código existente sem alterações) ...
        responsaveis_para_gerar = df_filtrado['RESPONSAVEL'].unique()
        data_hora_geracao = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
        
        # --- ESTILOS PARA O PDF ---
# ... (código existente sem alterações) ...
        COR_PRIMARIA = (60, 100, 160) # Azul escuro
        COR_SECUNDARIA = (240, 240, 240) # Cinza claro
        
        for responsavel in responsaveis_para_gerar:
# ... (código existente sem alterações) ...
            df_responsavel = df_filtrado[df_filtrado['RESPONSAVEL'] == responsavel]
            if df_responsavel.empty:
# ... (código existente sem alterações) ...
                print(f"⚠️ Não há dados para o responsável: {responsavel}.")
                continue

            pdf = FPDF(format='letter')
# ... (código existente sem alterações) ...
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            
            # Cabeçalho profissional
# ... (código existente sem alterações) ...
            pdf.set_fill_color(*COR_PRIMARIA)
            pdf.rect(0, 0, 220, 30, 'F')
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(0, 15, 'RELATÓRIO DE CONTINGÊNCIA', ln=1, align='C')
            pdf.set_font("Arial", 'B', 10)
            pdf.cell(0, 5, 'Equipamentos sem comunicação no dia atual', ln=1, align='C')
            pdf.ln(10)
            
            # Título e Informações do Responsável
# ... (código existente sem alterações) ...
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 10, f'Responsável: {responsavel}', ln=1, align='L')
            pdf.set_font("Arial", '', 8)
            pdf.cell(0, 5, f'Relatório gerado em: {data_hora_geracao}', ln=1, align='L')
            pdf.ln(5)

            # Cabeçalho da tabela
# ... (código existente sem alterações) ...
            pdf.set_fill_color(*COR_PRIMARIA)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Arial", 'B', 8)
            
            headers = ['DISPLAY', 'FROTA', 'TIPO EQUIP.', 'PRESTADOR', 'ÚLTIMA COMUNICAÇÃO']
# ... (código existente sem alterações) ...
            col_widths = [38, 38, 38, 38, 38]
            
            for i, header in enumerate(headers):
# ... (código existente sem alterações) ...
                pdf.cell(col_widths[i], 10, header, border=0, align='C', fill=True)
            pdf.ln()

            # Dados da tabela
# ... (código existente sem alterações) ...
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", size=8)
            
            for i, row in enumerate(df_responsavel.iterrows()):
# ... (código existente sem alterações) ...
                # Cor de fundo alternada
                if i % 2 == 0:
# ... (código existente sem alterações) ...
                    pdf.set_fill_color(255, 255, 255)
                else:
# ... (código existente sem alterações) ...
                    pdf.set_fill_color(*COR_SECUNDARIA)
                
                row_data = row[1]
                
                pdf.cell(col_widths[0], 10, str(row_data['DISPLAY']), border=0, align='C', fill=True)
                pdf.cell(col_widths[1], 10, str(row_data['Nro do equipamento']), border=0, align='C', fill=True)
                pdf.cell(col_widths[2], 10, str(row_data['Tipo do equipamento']), border=0, align='C', fill=True)
                pdf.cell(col_widths[3], 10, str(row_data['PRESTADOR']), border=0, align='C', fill=True)
                
                registro_formatado = row_data['Registro mais recente'].strftime('%d/%m/%Y %H:%M') if pd.notna(row_data['Registro mais recente']) else 'N/A'
                pdf.cell(col_widths[4], 10, registro_formatado, border=0, align='C', fill=True)
                pdf.ln()

            # Linha de rodapé
# ... (código existente sem alterações) ...
            pdf.ln(5)
            pdf.set_font("Arial", size=6, style='')
            pdf.cell(0, 10, "Relatório gerado automaticamente pelo Sistema.", ln=1, align='C')

            caminho_pdf = os.path.join(output_dir, f"Relatorio_{responsavel.replace(' ', '_')}.pdf")
# ... (código existente sem alterações) ...
            pdf.output(caminho_pdf)

            print(f"📄 PDF gerado com sucesso para {responsavel}: {caminho_pdf}")
# ... (código existente sem alterações) ...
            capturar_imagem_pdf_mupdf(caminho_pdf, output_dir, f"Relatorio_{responsavel.replace(' ', '_')}")

    except Exception as e:
# ... (código existente sem alterações) ...
        print(f"⚠️ Erro ao gerar PDF: {e}")

def formatar_excel(excel_output):
# ... (código existente sem alterações) ...
    """Formata a planilha Excel de saída, ajustando larguras e aplicando estilos."""
    try:
        wb = load_workbook(excel_output)
        ws = wb['Contingência']

        data_hora_geracao = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
        ws.insert_rows(1)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)
        cell_data = ws['A1']
        cell_data.value = f"Relatório gerado em: {data_hora_geracao}"
        cell_data.font = Font(bold=True, color="000000")
        cell_data.alignment = Alignment(horizontal='center', vertical='center')
        
        header_row = 2
# ... (código existente sem alterações) ...
        for cell in ws[header_row]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

        for col in ws.columns:
# ... (código existente sem alterações) ...
            max_length = 0
            column = col[1].column_letter
            
            for cell in col:
# ... (código existente sem alterações) ...
                try:
                    if cell.value:
                        if len(str(cell.value)) > max_length:
# ... (código existente sem alterações) ...
                            max_length = len(str(cell.value))
                except:
# ... (código existente sem alterações) ...
                    pass
            
            adjusted_width = max_length + 2
# ... (código existente sem alterações) ...
            ws.column_dimensions[column].width = adjusted_width

        border_style = Border(
# ... (código existente sem alterações) ...
            left=Side(border_style="thin"),
            right=Side(border_style="thin"),
            top=Side(border_style="thin"),
            bottom=Side(border_style="thin")
        )

        for row in ws.iter_rows(min_row=header_row, max_row=ws.max_row):
# ... (código existente sem alterações) ...
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border_style

        col_registro = 6 
# ... (código existente sem alterações) ...
        data_atual = datetime.datetime.now().date()
        for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
            registro_data = row[col_registro - 1].value 
            if registro_data and isinstance(registro_data, datetime.datetime):
                if registro_data.date() < data_atual:
# ... (código existente sem alterações) ...
                    for cell in row:
                        cell.fill = PatternFill(fill_type="none")

        wb.save(excel_output)
# ... (código existente sem alterações) ...
        print(f"🟢 Excel formatado com sucesso: {excel_output}")

    except Exception as e:
# ... (código existente sem alterações) ...
        print(f"⚠️ Erro ao formatar Excel: {e}")

# --- Funções da Interface Gráfica (Tkinter) ---
def alternar_visualizacao_senha():
# ... (código existente sem alterações) ...
    """Alterna a visualização da senha na interface."""
    if entry_senha.cget('show') == '*':
        entry_senha.config(show='')
        botao_visualizar.config(text="Ocultar")
    else:
        entry_senha.config(show='*')
        botao_visualizar.config(text="Mostrar")

def atualizar_campos_credenciais(credenciais_path):
# ... (código existente sem alterações) ...
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
# ... (código existente sem alterações) ...
    """Salva o usuário e a senha em um arquivo JSON."""
    usuario = entry_usuario.get()
    senha = entry_senha.get()
    
    if var_salvar_usuario.get():
# ... (código existente sem alterações) ...
        credenciais = {"usuario": usuario, "senha": senha}
        try:
            with open(credenciais_path, "w", encoding='utf-8') as file:
                json.dump(credenciais, file, indent=4)
            print(f"🟢 Usuário e senha salvos em: {credenciais_path}")
        except Exception as e:
            print(f"⚠️ Erro ao salvar credenciais: {e}")
    else:
# ... (código existente sem alterações) ...
        if os.path.exists(credenciais_path):
            os.remove(credenciais_path)
            print("🔴 Arquivo de credenciais removido.")

def cancelar_execucao():
# ... (código existente sem alterações) ...
    global execucao_ativa
    execucao_ativa = False
    print("🔴 Execução cancelada.")
    atualizar_progresso("Execução cancelada.", step=0, total_steps=1)


def atualizar_progresso(status_texto, step, total_steps):
# ... (código existente sem alterações) ...
    """
    Atualiza o rótulo de status e a barra de progresso.
    Esta função deve ser chamada na thread principal do Tkinter.
    """
    if root and status_label and progress_bar:
# ... (código existente sem alterações) ...
        # Garante que a atualização seja feita na thread principal
        root.after(0, _atualizar_progresso_thread_safe, status_texto, step, total_steps)

def _atualizar_progresso_thread_safe(status_texto, step, total_steps):
# ... (código existente sem alterações) ...
    """Função interna para atualização segura da interface."""
    if total_steps > 0:
        progress = (step / total_steps) * 100
        progress_bar['value'] = progress
    status_label.config(text=status_texto)
    root.update_idletasks() # Força a atualização da interface

def criar_interface():
# ... (código existente sem alterações) ...
    global entry_usuario, entry_senha, var_salvar_usuario, botao_visualizar, driver, responsaveis_vars, entry_intervalo, var_intervalo_ativado, status_label, progress_bar, root
    
    credenciais_path = os.path.join(script_dir, "credenciais.json")

    # Paleta de cores
# ... (código existente sem alterações) ...
    PALETTE = {
        "primary": "#0066AC",       # Azul Escuro
        "secondary": "#43948C",     # Verde Acinzentado
        "success": "#6BBE3B",       # Verde Claro
        "danger": "#B90000",        # Vermelho
        "background": "#FFFFFF",    # Branco
        "text": "#000000",          # Preto
    }
    
    root = ttk.Window(themename="yeti")
# ... (código existente sem alterações) ...
    root.title(f"HXG - Auto  v{VERSAO}")
    
    # Configura o estilo para usar a nova paleta de cores
    style = ttk.Style()
# ... (código existente sem alterações) ...
    style.configure("TLabel", font=("Helvetica", 11), background=PALETTE["background"], foreground=PALETTE["text"])
    style.configure("TFrame", background=PALETTE["background"])
    style.configure("TLabelframe", background=PALETTE["background"], foreground=PALETTE["text"])
    style.configure("TLabelframe.Label", background=PALETTE["background"], foreground=PALETTE["text"])
    style.configure("TEntry", fieldbackground="white", foreground=PALETTE["text"])

    # Estilizando os botões
# ... (código existente sem alterações) ...
    style.configure("success.TButton", background=PALETTE["success"], foreground="white", font=("Helvetica", 11, "bold"))
    style.configure("danger.TButton", background=PALETTE["danger"], foreground="white", font=("Helvetica", 11, "bold"))
    style.configure("info.TButton", background=PALETTE["primary"], foreground="white", font=("Helvetica", 11, "bold"))
    style.configure("secondary.TButton", background=PALETTE["secondary"], foreground="white", font=("Helvetica", 11, "bold"))
    style.map("TButton", background=[("active", PALETTE["primary"])])
    
    # Estilizando o checkbox
# ... (código existente sem alterações) ...
    style.configure("Roundtoggle.TCheckbutton", background=PALETTE["background"], foreground=PALETTE["text"])
    style.configure("info-round-toggle.TCheckbutton", background=PALETTE["background"], foreground=PALETTE["text"])

    # Define as variáveis de controle para os checkboxes
# ... (código existente sem alterações) ...
    responsaveis_vars = {nome: tk.BooleanVar() for nome in RESPONSAVEIS_OPCOES}
    
    # Frame principal
    main_frame = ttk.Frame(root, padding=20)
# ... (código existente sem alterações) ...
    main_frame.pack(fill="both", expand=True)
    
    ttk.Label(main_frame, text="AUTO. CONTIGÊNCIA - HXG", font=("Helvetica", 20, "bold"), foreground=PALETTE["primary"]).pack(pady=(0, 20))

    # Frame de credenciais
# ... (código existente sem alterações) ...
    cred_frame = ttk.Labelframe(main_frame, text="Credenciais", padding=15)
    cred_frame.pack(fill="x", pady=10)
    
    ttk.Label(cred_frame, text="Usuário:").pack(anchor="w", pady=(0, 5))
# ... (código existente sem alterações) ...
    entry_usuario = ttk.Entry(cred_frame, width=40)
    entry_usuario.pack(fill="x")

    ttk.Label(cred_frame, text="Senha:").pack(anchor="w", pady=(10, 5))
# ... (código existente sem alterações) ...
    frame_senha = ttk.Frame(cred_frame)
    frame_senha.pack(fill="x")
    
    entry_senha = ttk.Entry(frame_senha, show="*")
# ... (código existente sem alterações) ...
    entry_senha.pack(side="left", fill="x", expand=True)

    botao_visualizar = ttk.Button(frame_senha, text="Mostrar", command=alternar_visualizacao_senha)
# ... (código existente sem alterações) ...
    botao_visualizar.pack(side="left", padx=(5,0))

    var_salvar_usuario = tk.BooleanVar()
# ... (código existente sem alterações) ...
    credenciais_existentes, usuario_carregado, senha_carregada = atualizar_campos_credenciais(credenciais_path)
    var_salvar_usuario.set(credenciais_existentes)

    if credenciais_existentes:
# ... (código existente sem alterações) ...
        entry_usuario.insert(0, usuario_carregado)
        entry_senha.insert(0, senha_carregada)
    
    ttk.Checkbutton(cred_frame, text="Salvar usuário e senha", variable=var_salvar_usuario, bootstyle="round-toggle").pack(anchor="w", pady=(10, 0))

    # Campo para o intervalo de tempo
# ... (código existente sem alterações) ...
    intervalo_frame = ttk.Frame(main_frame)
    intervalo_frame.pack(fill="x", pady=(10, 5))

    ttk.Label(intervalo_frame, text="Executar a cada (minutos):").pack(side="left", padx=(0, 5))
# ... (código existente sem alterações) ...
    entry_intervalo = ttk.Entry(intervalo_frame, width=10)
    entry_intervalo.insert(0, "60") # Valor padrão de 60 minutos
    entry_intervalo.pack(side="left", padx=(0, 10))
    
    # Checkbox para ativar/desativar o agendamento
# ... (código existente sem alterações) ...
    var_intervalo_ativado = tk.BooleanVar(value=False) # Por padrão, o agendamento fica ativo
    ttk.Checkbutton(intervalo_frame, text="Ativar agendamento", variable=var_intervalo_ativado, bootstyle="round-toggle").pack(side="left")


    # Seletor de Responsáveis
# ... (código existente sem alterações) ...
    resp_frame = ttk.Labelframe(main_frame, text="Gerar PDF para:", padding=15)
    resp_frame.pack(fill="both", expand=True, pady=10)

    def selecionar_todos():
# ... (código existente sem alterações) ...
        for var in responsaveis_vars.values():
            var.set(True)

    def limpar_selecao():
# ... (código existente sem alterações) ...
        for var in responsaveis_vars.values():
            var.set(False)

    btn_frame = ttk.Frame(resp_frame)
# ... (código existente sem alterações) ...
    btn_frame.pack(fill="x", pady=(0, 5))
    ttk.Button(btn_frame, text="Selecionar Todos", command=selecionar_todos, bootstyle="info").pack(side="left", fill="x", expand=True, padx=(0, 5))
    ttk.Button(btn_frame, text="Limpar Seleção", command=limpar_selecao, bootstyle="secondary").pack(side="left", fill="x", expand=True, padx=(5, 0))
    
    # Criação dos checkboxes
# ... (código existente sem alterações) ...
    for nome, var in responsaveis_vars.items():
        ttk.Checkbutton(resp_frame, text=nome, variable=var, bootstyle="info-round-toggle").pack(anchor="w", pady=2)

    # Barra de progresso e status
# ... (código existente sem alterações) ...
    status_label = ttk.Label(main_frame, text="Aguardando...", font=("Helvetica", 10), foreground="gray")
    status_label.pack(pady=(10, 5))

    progress_bar = ttk.Progressbar(main_frame, mode="determinate", bootstyle="info")
# ... (código existente sem alterações) ...
    progress_bar.pack(fill="x", pady=(0, 10))

    # Botões de ação
    action_frame = ttk.Frame(main_frame)
# ... (código existente sem alterações) ...
    action_frame.pack(fill="x", pady=10)

    ttk.Button(action_frame, text="Executar", command=lambda: executar_script(), bootstyle="success").pack(side="left", fill="x", expand=True, padx=(0, 5))
    ttk.Button(action_frame, text="Pausar", command=cancelar_execucao, bootstyle="danger").pack(side="left", fill="x", expand=True, padx=(5, 0))

    root.bind('<Return>', lambda event: executar_script())
    
    def fechar_janela():
# ... (código existente sem alterações) ...
        salvar_usuario(credenciais_path)
        if 'driver' in globals() and driver:
            try:
# ... (código existente sem alterações) ...
                driver.quit()
            except:
# ... (código existente sem alterações) ...
                pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", fechar_janela)
# ... (código existente sem alterações) ...
    root.mainloop()

# --- Funções de Execução Principal ---
def executar_script():
# ... (código existente sem alterações) ...
    global execucao_ativa
    if execucao_ativa:
        messagebox.showinfo("Informação", "A automação já está em execução.")
        return

    usuario = entry_usuario.get()
# ... (código existente sem alterações) ...
    senha = entry_senha.get()
    
    if not usuario or not senha:
# ... (código existente sem alterações) ...
        messagebox.showwarning("Aviso", "Por favor, preencha o usuário e a senha.")
        return

    credenciais_path = os.path.join(script_dir, "credenciais.json")
# ... (código existente sem alterações) ...
    salvar_usuario(credenciais_path)
    
    # Inicia a barra de progresso antes de iniciar o procedimento
# ... (código existente sem alterações) ...
    atualizar_progresso("Iniciando a automação...", step=0, total_steps=7) # Alterado de 8 para 7
    
    threading.Thread(target=executar_procedimento, args=(usuario, senha), daemon=True).start()

def executar_procedimento(usuario, senha):
# ... (código existente sem alterações) ...
    global driver, execucao_ativa, responsaveis_vars
    execucao_ativa = True
    
    # Definição das etapas para a barra de progresso
# ... (código existente sem alterações) ...
    TOTAL_STEPS = 7 # Alterado de 8 para 7
    last_valid_interval = 60 # Valor padrão
    
    while execucao_ativa:
# ... (código existente sem alterações) ...
        # Coleta os responsáveis e as configurações da UI a cada novo ciclo
        selected_responsaveis = [nome for nome, var in responsaveis_vars.items() if var.get()]
        intervalo_ativado = var_intervalo_ativado.get()
        
        try:
# ... (código existente sem alterações) ...
            intervalo_minutos = int(entry_intervalo.get())
            if intervalo_minutos <= 0:
                print(f"⚠️ Intervalo inválido ({intervalo_minutos}). Usando o último valor válido: {last_valid_interval} min.")
                intervalo_minutos = last_valid_interval
            else:
# ... (código existente sem alterações) ...
                last_valid_interval = intervalo_minutos
        except (ValueError, tk.TclError):
            print(f"⚠️ Erro ao ler o intervalo. Usando o último valor válido: {last_valid_interval} min.")
# ... (código existente sem alterações) ...
            intervalo_minutos = last_valid_interval
            
        print(f"\n--- Iniciando novo ciclo ---")
# ... (código existente sem alterações) ...
        print(f"Responsáveis selecionados para este ciclo: {', '.join(selected_responsaveis) or 'Nenhum'}")
        print(f"Intervalo configurado: {intervalo_minutos} minutos. Agendamento Ativado: {'Sim' if intervalo_ativado else 'Não'}")

        driver = None
# ... (código existente sem alterações) ...
        df_final = None 

        try:
# ... (código existente sem alterações) ...
            # Definição de caminhos
            diretorio_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
            excel_output = os.path.join(script_dir, "Contingencia - Final.xlsx")
            base_responsaveis_path = os.path.join(script_dir, "Base - Respon.xlsx")
            pdf_output_dir = os.path.join(script_dir, "PDF_Saida")

            xpaths = {
# ... (código existente sem alterações) ...
                'usuario': '/html/body/app-root/app-login/app-access-container/div/div[2]/div[3]/form/div[1]/input',
                'senha': '/html/body/app-root/app-login/app-access-container/div/div[2]/div[3]/form/div[2]/input',
                'botao_login': '/html/body/app-root/app-login/app-access-container/div/div[2]/div[3]/form/div[3]/p-button/button/span',
                'control_room': '/html/body/app-root/app-home/div/div/div[1]/div[1]/div/div/div/div/div[3]/a[1]',
                'limpar_filtro': '//button[contains(@id,"buttion-id-clearAndApplyButton")]',
                'tabela': '//*[@id="div-submenu-link-id-app-submenu-link-mon-table-id"]/p',
                'filtro': '//*[@id="pn_id_5"]/div[1]/div/div[2]/button[3]',
                'exportacao_csv': '//*[@id="pn_id_5"]/div[1]/div/div[2]/button[1]/i'
            }
            url = "https://access.hxgnagron.com/"
            
            # --- ETAPA DE LIMPEZA DE CACHE REMOVIDA ---
# ... (código existente sem alterações) ...
            # atualizar_progresso("Limpando cache...", step=1, ...)

            atualizar_progresso("Iniciando driver...", step=1, total_steps=TOTAL_STEPS) # Alterado de step=2
            driver = iniciar_driver(headless=True)
            
            if not execucao_ativa: break
            
            atualizar_progresso("Realizando login...", step=2, total_steps=TOTAL_STEPS) # Alterado de step=3
# ... (código existente sem alterações) ...
            login_usuario(driver, url, usuario, senha, xpaths)
            
            if not execucao_ativa: break
            
            atualizar_progresso("Exportando tabela...", step=3, total_steps=TOTAL_STEPS) # Alterado de step=4
# ... (código existente sem alterações) ...
            exportar_tabela(driver, xpaths)
            
            if not execucao_ativa: break

            atualizar_progresso("Aguardando download do CSV...", step=4, total_steps=TOTAL_STEPS) # Alterado de step=5
# ... (código existente sem alterações) ...
            df_final = processar_csv(diretorio_downloads, excel_output, base_responsaveis_path, pdf_output_dir, selected_responsaveis)
            
            if df_final is not None:
# ... (código existente sem alterações) ...
                if not execucao_ativa: break
                atualizar_progresso("Formatando Excel...", step=5, total_steps=TOTAL_STEPS) # Alterado de step=6
                formatar_excel(excel_output)
                
                if not execucao_ativa: break
                atualizar_progresso("Gerando PDFs...", step=6, total_steps=TOTAL_STEPS) # Alterado de step=7
# ... (código existente sem alterações) ...
                salvar_pdf_por_responsavel(df_final, pdf_output_dir)

                if not execucao_ativa: break
                atualizar_progresso("Atualizando planilha de controle...", step=7, total_steps=TOTAL_STEPS) # Alterado de step=8
# ... (código existente sem alterações) ...
                atualizar_coleta_planilha(df_final)

                atualizar_progresso("Procedimento concluído com sucesso!", step=7, total_steps=TOTAL_STEPS) # Alterado de step=8
            else:
# ... (código existente sem alterações) ...
                atualizar_progresso("Processamento de dados falhou.", step=0, total_steps=1)
        
        except Exception as e:
# ... (código existente sem alterações) ...
            # Captura exceções e informa na interface
            error_message = f"❌ Erro fatal na execução: {type(e).__name__}: {str(e)[:100]}..."
            print(error_message)
            logger.error(f"❌ Erro fatal na execução do procedimento: {e}")
            atualizar_progresso(error_message, step=0, total_steps=1)
        finally:
# ... (código existente sem alterações) ...
            if driver:
                try:
# ... (código existente sem alterações) ...
                    driver.quit()
                except:
# ... (código existente sem alterações) ...
                    pass

        # Decide se continua o loop ou encerra
# ... (código existente sem alterações) ...
        if not intervalo_ativado:
            print("Execução única concluída, pois o agendamento está desativado.")
            break # Encerra o loop principal

        # Pausa para o intervalo definido, com verificação de cancelamento e contagem regressiva
# ... (código existente sem alterações) ...
        if execucao_ativa:
            tempo_total_espera = intervalo_minutos * 60
            print(f"✅ Execução concluída. Aguardando {intervalo_minutos} minutos para a próxima rodada...")

            for segundos_restantes in range(tempo_total_espera, 0, -1):
# ... (código existente sem alterações) ...
                if not execucao_ativa: break
                
                minutos, segundos = divmod(segundos_restantes, 60)
# ... (código existente sem alterações) ...
                texto_tempo = f"Próxima execução em {minutos:02d}:{segundos:02d}"
                atualizar_progresso(texto_tempo, step=7, total_steps=TOTAL_STEPS) # Mantém a barra cheia
                time.sleep(1)
            
            if not execucao_ativa: break
# ... (código existente sem alterações) ...
    
    execucao_ativa = False
# ... (código existente sem alterações) ...
    atualizar_progresso("Procedimento finalizado.", step=0, total_steps=1)
    print("🏁 Procedimento finalizado.")


if __name__ == "__main__":
# ... (código existente sem alterações) ...
    driver = None
    criar_interface()