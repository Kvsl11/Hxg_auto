import threading
import glob
import pandas as pd
import undetected_chromedriver as uc
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
import warnings  # << adicione

# --- CORREÇÃO SSL HÍBRIDA E ATUALIZAÇÃO AUTOMÁTICA DO CERTIFICADO ---
import os
import ssl
import subprocess
import urllib.request
import logging
import sys
import requests
import customtkinter as ctk

# Caminho dinâmico da pasta onde o script está localizado
app_dir = os.path.dirname(os.path.abspath(__file__))
python_exe = os.path.join(app_dir, "App_py", "Python313", "Python313", "pythonw.exe")

# Configuração de logging (arquivo + console)
log_path = os.path.join(app_dir, "ssl_patch.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(log_path, encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Apenas o certificado raiz da Amazon (necessário para access.hxgnagron.com)
AMAZON_CERTS = {
    "Amazon Root CA 1": "https://www.amazontrust.com/repository/AmazonRootCA1.pem"
}

def atualizar_certifi():
    """Verifica e atualiza o pacote certifi usando o Python interno (Python313)."""
    try:
        if not os.path.exists(python_exe):
            logger.warning(f"⚠️ Python interno não encontrado em: {python_exe}")
            return
        logger.info("🔍 Verificando e atualizando pacote certifi no Python interno...")
        subprocess.run([python_exe, "-m", "pip", "install", "--upgrade", "certifi"], check=True)
        import certifi
        logger.info(f"🟢 Certifi atualizado com sucesso. Caminho: {certifi.where()}")
    except Exception as e:
        logger.warning(f"⚠️ Falha ao atualizar certifi: {e}")

def garantir_certificados_amazon():
    """Verifica se o certificado raiz da Amazon está presente e adiciona se necessário."""
    try:
        import certifi
        cacert_path = certifi.where()

        with open(cacert_path, "r", encoding="utf-8") as f:
            conteudo = f.read()

        for nome, url in AMAZON_CERTS.items():
            if nome not in conteudo:
                logger.info(f"🔍 {nome} não encontrado, baixando de {url}...")
                resp = requests.get(url, timeout=10)
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

def testar_ssl():
    """Testa a verificação SSL e aplica fallback automático caso falhe."""
    try:
        import certifi
        ssl_context = ssl.create_default_context(cafile=certifi.where())
        urllib.request.urlopen("https://www.google.com", timeout=5, context=ssl_context)
        logger.info("🟢 Conexão SSL validada com sucesso — certificados OK.")
    except ssl.SSLError as e:
        logger.warning(f"⚠️ Falha de SSL detectada ({e}). Aplicando modo não verificado.")
        ssl._create_default_https_context = ssl._create_unverified_context
        try:
            urllib.request.urlopen("https://www.google.com", timeout=5)
            logger.info("🟡 SSL desativado — conexão forçada sem verificação de certificado.")
        except Exception as e2:
            logger.error(f"❌ Mesmo após desativar SSL, a conexão falhou: {e2}")
    except Exception as e:
        logger.warning(f"⚠️ Erro genérico ao testar SSL: {e}. Aplicando fallback.")
        ssl._create_default_https_context = ssl._create_unverified_context

# --- EXECUÇÃO AUTOMÁTICA AO INICIAR ---
logger.info("🚀 Iniciando verificação e correção SSL híbrida...")
atualizar_certifi()
garantir_certificados_amazon()
testar_ssl()
logger.info("✅ Configuração SSL concluída com segurança.")

# --- VERIFICAÇÃO DE ATUALIZAÇÃO VIA GITHUB ---
VERSAO = "3.0.6"

def verificar_atualizacao_disponivel(root=None, frame_status=None):
    """Verifica no GitHub se há nova versão e atualiza automaticamente, se desejado."""
    try:
        repo_url = "https://raw.githubusercontent.com/Kvsl11/Hxg_auto/main/version.txt"
        script_url = "https://raw.githubusercontent.com/Kvsl11/Hxg_auto/main/main.py"
        versao_local = VERSAO

        # Mostra status inicial de verificação
        if frame_status:
            for widget in frame_status.winfo_children():
                widget.destroy()
            status_label = ctk.CTkLabel(frame_status, text="🔄 Verificando atualizações...", text_color="#ffffff")
            status_label.pack(pady=2)

        # Busca versão online
        resposta = requests.get(repo_url, timeout=8, verify=False)
        if resposta.status_code != 200:
            raise Exception(f"Erro HTTP {resposta.status_code}")

        versao_online = resposta.text.strip()

        # Limpa frame
        if frame_status:
            for widget in frame_status.winfo_children():
                widget.destroy()

        if versao_online != versao_local:
            # Nova versão detectada
            label = ctk.CTkLabel(
                frame_status,
                text=f"🟡 Nova versão disponível: v{versao_online}",
                text_color="#fff8dc",
                font=ctk.CTkFont(weight="bold")
            )
            label.pack(side="left", padx=10, pady=3)

            def baixar_e_atualizar():
                try:
                    label.configure(text="⬇ Baixando atualização...")
                    btn_update.configure(state="disabled")
                    frame_status.update()

                    # Baixa o novo main.py
                    r = requests.get(script_url, timeout=15, verify=False)
                    r.raise_for_status()

                    # Substitui o arquivo local
                    local_path = os.path.join(os.path.dirname(__file__), "main.py")
                    with open(local_path, "wb") as f:
                        f.write(r.content)

                    # Atualiza a versão no arquivo version_local.txt
                    version_local = os.path.join(os.path.dirname(__file__), "version_local.txt")
                    with open(version_local, "w", encoding="utf-8") as vf:
                        vf.write(versao_online)

                    messagebox.showinfo("Atualização concluída", f"✅ Atualizado para v{versao_online}.\nO app será reiniciado.")
                    subprocess.Popen(["python", local_path])
                    os._exit(0)
                except Exception as e:
                    messagebox.showerror("Erro", f"⚠️ Falha ao atualizar: {e}")

            btn_update = ctk.CTkButton(
                frame_status,
                text="⬇ Atualizar agora",
                fg_color="#ffaa00",
                hover_color="#cc8800",
                text_color="#000000",
                width=150,
                command=baixar_e_atualizar
            )
            btn_update.pack(side="right", padx=10, pady=3)

        else:
            # Já está atualizado
            label = ctk.CTkLabel(
                frame_status,
                text=f"🟢 Atualizado — v{VERSAO}",
                text_color="#43948c",
                font=ctk.CTkFont(weight="bold")
            )
            label.pack(pady=3)

    except Exception as e:
        if frame_status:
            for widget in frame_status.winfo_children():
                widget.destroy()
            ctk.CTkLabel(
                frame_status,
                text=f"⚠️ Falha ao verificar atualização: {e}",
                text_color="#ffcc00"
            ).pack(pady=3)

warnings.filterwarnings(
    "ignore",
    message="Slicer List extension is not supported and will be removed",
    category=UserWarning,
    module=r"openpyxl\.worksheet\._reader"
)

# --- Definições Globais ---
# Descobre dinamicamente o diretório onde o script está localizado
script_dir = os.path.dirname(os.path.abspath(__file__))
execucao_ativa = False
# Declarações globais para os widgets da barra de progresso e status
status_label = None
progress_bar = None
root = None

# Lista de responsáveis para o seletor da interface
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

# --- Funções de Automação com Selenium (usando undetected_chromedriver) ---
def iniciar_driver(headless=True):
    """Inicia uma instância do Chrome usando undetected_chromedriver."""
    print("🚀 Iniciando driver com undetected_chromedriver...")
    logger.info("🚀 Iniciando driver com undetected_chromedriver...")

    options = uc.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")

    try:
        driver = uc.Chrome(options=options, use_subprocess=True)
    except Exception as e:
        print(f"⚠️ Falha no SSL ou no driver, tentando ignorar validação: {e}")
        import ssl
        ssl._create_default_https_context = ssl._create_unverified_context
        driver = uc.Chrome(options=options, use_subprocess=True)

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
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath)))
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
    aguardar_e_clicar(driver, xpaths['control_room'])
    
    limpar_filtro_xpath = '//button[contains(@id,"buttion-id-clearAndApplyButton")]'
    aguardar_e_clicar(driver, limpar_filtro_xpath)
    
    aguardar_e_clicar(driver, xpaths['tabela'])
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

# --- Funções de Processamento de Dados (Pandas, OpenPyXL) ---

def processar_csv(diretorio_downloads, excel_output, base_responsaveis_path, pdf_output_dir, selected_responsaveis):
    try:
        if not os.path.exists(pdf_output_dir):
            os.makedirs(pdf_output_dir)
        
        print("⏳ Aguardando download do arquivo CSV...")
        csv_path = aguardar_download_completo(diretorio_downloads, "Monitoramento - Tabela")
        if not csv_path:
            print("❌ Processo encerrado. Nenhum arquivo CSV disponível.")
            return None
        
        df = pd.read_csv(csv_path, encoding="utf-8", sep=";", dtype=str)
        df["Registro mais recente"] = pd.to_datetime(df["Registro mais recente"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
        
        data_atual = datetime.datetime.now().date()
        df_antigos = df[df["Registro mais recente"].dt.date != data_atual]
        df_responsaveis = pd.read_excel(base_responsaveis_path, dtype=str)

        col_equipamento = "Nro do equipamento"
        col_responsavel = "RESPONSAVEL"
        col_display = "DISPLAY"
        col_prestador = "PRESTADOR"
        
        if col_equipamento not in df_antigos.columns or col_equipamento not in df_responsaveis.columns:
            print(f"⚠️ A coluna '{col_equipamento}' não foi encontrada em uma das planilhas!")
            return None
        
        df_final = df_antigos.merge(df_responsaveis[[col_equipamento, col_responsavel, col_display, col_prestador]], on=col_equipamento, how="left")
        df_final = df_final.dropna(subset=[col_responsavel])
        
        # Filtra por responsáveis selecionados se a lista não estiver vazia
        if selected_responsaveis:
            print(f"✅ Gerando relatórios apenas para: {', '.join(selected_responsaveis)}")
            df_final = df_final[df_final[col_responsavel].isin(selected_responsaveis)]
        else:
            print("🔄 Nenhum responsável selecionado. Gerando relatórios para todos os responsáveis.")

        colunas_desejadas = [
            "RESPONSAVEL", "DISPLAY", "Nro do equipamento",
            "Tipo do equipamento", "PRESTADOR", "Registro mais recente"
        ]
        df_final = df_final[colunas_desejadas]
        
        try:
            with pd.ExcelWriter(excel_output, engine='openpyxl', mode='w') as writer:
                df_final.to_excel(writer, index=False, sheet_name="Contingência")
            print(f"🟢 Registros antigos com responsáveis salvos em: {excel_output}")
        except Exception as save_error:
            print(f"⚠️ Erro ao salvar Excel: {save_error}")
            return None

        return df_final
    except Exception as e:
        print(f"⚠️ Erro ao processar CSV: {e}")
        return None

def obter_caminho_planilha():
    import os, datetime
    
    ano_atual = datetime.datetime.now().year
    mes_atual = datetime.datetime.now().month

    meses = {
        1: "01 - Janeiro", 2: "02 - Fevereiro", 3: "03 - Março",
        4: "04 - Abril", 5: "05 - Maio", 6: "06 - Junho",
        7: "07 - Julho", 8: "08 - Agosto", 9: "09 - Setembro",
        10: "10 - Outubro", 11: "11 - Novembro", 12: "12 - Dezembro"
    }

    numero_safra = 2.5 + (ano_atual - 2025)
    safra = f"{numero_safra:.1f} - Safra {ano_atual}"

    possiveis_drives = ["I:", "Z:"]

    caminho_final = None
    for drive in possiveis_drives:
        base = fr"{drive}\ANG\Agricola\Controle\Computador de Bordo\Fechamento Prestação de Serviço (Linha Amarela)\Pago pelo Bordo"
        caminho_teste = os.path.join(base, safra, meses[mes_atual], "Monitoramento - Eqps.xlsx")
        if os.path.exists(caminho_teste):
            caminho_final = caminho_teste
            break

    if not caminho_final:
        raise FileNotFoundError("❌ Não foi possível localizar a planilha em nenhum dos caminhos (I: ou Z:).")

    return caminho_final

def atualizar_coleta_planilha(df_final):
    try:
        caminho_planilha = obter_caminho_planilha()

        aba_alvo = "Cont. Maquinas"
        
        wb = load_workbook(caminho_planilha)
        ws = wb[aba_alvo]
        
        col_equipamento_final = "Nro do equipamento"
        cabecalhos = {cell.value.strip().upper(): cell.column for cell in ws[1] if cell.value}

        if "EQUIPAMENTO" not in cabecalhos or "COLETA" not in cabecalhos:
            print("⚠️ Colunas necessárias não encontradas na aba Cont. Maquinas.")
            return
        
        col_equipamento_contingencia = cabecalhos["EQUIPAMENTO"]
        col_coleta = cabecalhos["COLETA"]
        
        equipamentos_contingencia = set(df_final[col_equipamento_final].astype(str).str.strip())
        equipamentos_coletados = set()
        
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=False):
            equipamento_cell = row[col_equipamento_contingencia - 1]
            coleta_cell = row[col_coleta - 1]
            if coleta_cell.value == "DADOS COLETADOS" and equipamento_cell.value:
                equipamentos_coletados.add(str(equipamento_cell.value).strip())

        bold_font = Font(bold=True)
        
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=False):
            equipamento_cell = row[col_equipamento_contingencia - 1]
            cell = row[col_coleta - 1]
            
            if cell.value == "DADOS COLETADOS":
                continue
            
            if (equipamento_cell.value and 
                str(equipamento_cell.value).strip() in equipamentos_contingencia and 
                str(equipamento_cell.value).strip() not in equipamentos_coletados):
                cell.value = "COLETAR DADOS"
                cell.font = bold_font
            else:
                cell.value = None
        
        wb.save(caminho_planilha)
        print("✅ Atualização da coluna COLETA concluída com sucesso.")

    except Exception as e:
        print(f"⚠️ Erro ao atualizar a planilha de Equipamentos: {e}") 

# --- Funções de Geração de Arquivos (PDF, Imagem) ---
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
        # Garante que o diretório de saída existe
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Remove todos os arquivos PDF existentes no diretório antes de gerar novos
        for arquivo in os.listdir(output_dir):
            if arquivo.endswith(".pdf") or arquivo.endswith(".png"):
                os.remove(os.path.join(output_dir, arquivo))

        caminho_planilha = obter_caminho_planilha()
        aba_alvo = "Cont. Maquinas"
        wb = load_workbook(caminho_planilha, data_only=True, read_only=True)
        ws = wb[aba_alvo]

        colunas = {str(cell.value).strip().upper(): idx for idx, cell in enumerate(ws[1]) if cell.value}
        if "EQUIPAMENTO" not in colunas or "COLETA" not in colunas:
            print("⚠️ Colunas necessárias não encontradas na aba Cont. Maquinas.")
            return

        equipamentos_coletados = set()
        col_equipamento = colunas["EQUIPAMENTO"]
        col_coleta = colunas["COLETA"]

        for row in ws.iter_rows(min_row=2, values_only=True):
            equipamento = str(row[col_equipamento]).strip()
            coleta = str(row[col_coleta]).strip() if row[col_coleta] else ""
            if coleta == "DADOS COLETADOS":
                equipamentos_coletados.add(equipamento)

        df_filtrado = df_final[~df_final["Nro do equipamento"].astype(str).isin(equipamentos_coletados)]
        responsaveis_para_gerar = df_filtrado['RESPONSAVEL'].unique()
        data_hora_geracao = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
        
        # --- ESTILOS PARA O PDF ---
        COR_PRIMARIA = (60, 100, 160) # Azul escuro
        COR_SECUNDARIA = (240, 240, 240) # Cinza claro
        
        for responsavel in responsaveis_para_gerar:
            df_responsavel = df_filtrado[df_filtrado['RESPONSAVEL'] == responsavel]
            if df_responsavel.empty:
                print(f"⚠️ Não há dados para o responsável: {responsavel}.")
                continue

            pdf = FPDF(format='letter')
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            
            # Cabeçalho profissional
            pdf.set_fill_color(*COR_PRIMARIA)
            pdf.rect(0, 0, 220, 30, 'F')
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(0, 15, 'RELATÓRIO DE CONTINGÊNCIA', ln=1, align='C')
            pdf.set_font("Arial", 'B', 10)
            pdf.cell(0, 5, 'Equipamentos sem comunicação no dia atual', ln=1, align='C')
            pdf.ln(10)
            
            # Título e Informações do Responsável
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 10, f'Responsável: {responsavel}', ln=1, align='L')
            pdf.set_font("Arial", '', 8)
            pdf.cell(0, 5, f'Relatório gerado em: {data_hora_geracao}', ln=1, align='L')
            pdf.ln(5)

            # Cabeçalho da tabela
            pdf.set_fill_color(*COR_PRIMARIA)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Arial", 'B', 8)
            
            headers = ['DISPLAY', 'FROTA', 'TIPO EQUIP.', 'PRESTADOR', 'ÚLTIMA COMUNICAÇÃO']
            col_widths = [38, 38, 38, 38, 38]
            
            for i, header in enumerate(headers):
                pdf.cell(col_widths[i], 10, header, border=0, align='C', fill=True)
            pdf.ln()

            # Dados da tabela
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", size=8)
            
            for i, row in enumerate(df_responsavel.iterrows()):
                # Cor de fundo alternada
                if i % 2 == 0:
                    pdf.set_fill_color(255, 255, 255)
                else:
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
            pdf.ln(5)
            pdf.set_font("Arial", size=6, style='')
            pdf.cell(0, 10, "Relatório gerado automaticamente pelo Sistema.", ln=1, align='C')

            caminho_pdf = os.path.join(output_dir, f"Relatorio_{responsavel.replace(' ', '_')}.pdf")
            pdf.output(caminho_pdf)

            print(f"📄 PDF gerado com sucesso para {responsavel}: {caminho_pdf}")
            capturar_imagem_pdf_mupdf(caminho_pdf, output_dir, f"Relatorio_{responsavel.replace(' ', '_')}")

    except Exception as e:
        print(f"⚠️ Erro ao gerar PDF: {e}")

def formatar_excel(excel_output):
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
        for cell in ws[header_row]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

        for col in ws.columns:
            max_length = 0
            column = col[1].column_letter
            
            for cell in col:
                try:
                    if cell.value:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = max_length + 2
            ws.column_dimensions[column].width = adjusted_width

        border_style = Border(
            left=Side(border_style="thin"),
            right=Side(border_style="thin"),
            top=Side(border_style="thin"),
            bottom=Side(border_style="thin")
        )

        for row in ws.iter_rows(min_row=header_row, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border_style

        col_registro = 6 
        data_atual = datetime.datetime.now().date()
        for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
            registro_data = row[col_registro - 1].value 
            if registro_data and isinstance(registro_data, datetime.datetime):
                if registro_data.date() < data_atual:
                    for cell in row:
                        cell.fill = PatternFill(fill_type="none")

        wb.save(excel_output)
        print(f"🟢 Excel formatado com sucesso: {excel_output}")

    except Exception as e:
        print(f"⚠️ Erro ao formatar Excel: {e}")

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
    """
    Atualiza o rótulo de status e a barra de progresso.
    Esta função deve ser chamada na thread principal do Tkinter.
    """
    if root and status_label and progress_bar:
        # Garante que a atualização seja feita na thread principal
        root.after(0, _atualizar_progresso_thread_safe, status_texto, step, total_steps)

def _atualizar_progresso_thread_safe(status_texto, step, total_steps):
    """Função interna para atualização segura da interface."""
    if total_steps > 0:
        progress = (step / total_steps) * 100
        progress_bar['value'] = progress
    status_label.config(text=status_texto)
    root.update_idletasks() # Força a atualização da interface

def criar_interface():
    global entry_usuario, entry_senha, var_salvar_usuario, botao_visualizar, driver, responsaveis_vars, entry_intervalo, var_intervalo_ativado, status_label, progress_bar, root
    
    credenciais_path = os.path.join(script_dir, "credenciais.json")

    # Paleta de cores
    PALETTE = {
        "primary": "#0066AC",      # Azul Escuro
        "secondary": "#43948C",    # Verde Acinzentado
        "success": "#6BBE3B",      # Verde Claro
        "danger": "#B90000",       # Vermelho
        "background": "#FFFFFF",   # Branco
        "text": "#000000",         # Preto
    }
    
    root = ttk.Window(themename="yeti")
    root.title(f"HXG - Auto  v{VERSAO}")
    
    # Configura o estilo para usar a nova paleta de cores
    style = ttk.Style()
    style.configure("TLabel", font=("Helvetica", 11), background=PALETTE["background"], foreground=PALETTE["text"])
    style.configure("TFrame", background=PALETTE["background"])
    style.configure("TLabelframe", background=PALETTE["background"], foreground=PALETTE["text"])
    style.configure("TLabelframe.Label", background=PALETTE["background"], foreground=PALETTE["text"])
    style.configure("TEntry", fieldbackground="white", foreground=PALETTE["text"])

    # Estilizando os botões
    style.configure("success.TButton", background=PALETTE["success"], foreground="white", font=("Helvetica", 11, "bold"))
    style.configure("danger.TButton", background=PALETTE["danger"], foreground="white", font=("Helvetica", 11, "bold"))
    style.configure("info.TButton", background=PALETTE["primary"], foreground="white", font=("Helvetica", 11, "bold"))
    style.configure("secondary.TButton", background=PALETTE["secondary"], foreground="white", font=("Helvetica", 11, "bold"))
    style.map("TButton", background=[("active", PALETTE["primary"])])
    
    # Estilizando o checkbox
    style.configure("Roundtoggle.TCheckbutton", background=PALETTE["background"], foreground=PALETTE["text"])
    style.configure("info-round-toggle.TCheckbutton", background=PALETTE["background"], foreground=PALETTE["text"])

    # Define as variáveis de controle para os checkboxes
    responsaveis_vars = {nome: tk.BooleanVar() for nome in RESPONSAVEIS_OPCOES}
    
    # Frame principal
    main_frame = ttk.Frame(root, padding=20)
    main_frame.pack(fill="both", expand=True)
    
    ttk.Label(main_frame, text="AUTO. CONTIGÊNCIA - HXG", font=("Helvetica", 20, "bold"), foreground=PALETTE["primary"]).pack(pady=(0, 20))

    # Frame de credenciais
    cred_frame = ttk.LabelFrame(main_frame, text="Credenciais", padding=15)
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
    botao_visualizar.pack(side="left", padx=(5,0))

    var_salvar_usuario = tk.BooleanVar()
    credenciais_existentes, usuario_carregado, senha_carregada = atualizar_campos_credenciais(credenciais_path)
    var_salvar_usuario.set(credenciais_existentes)

    if credenciais_existentes:
        entry_usuario.insert(0, usuario_carregado)
        entry_senha.insert(0, senha_carregada)
    
    ttk.Checkbutton(cred_frame, text="Salvar usuário e senha", variable=var_salvar_usuario, bootstyle="round-toggle").pack(anchor="w", pady=(10, 0))

    # Campo para o intervalo de tempo
    intervalo_frame = ttk.Frame(main_frame)
    intervalo_frame.pack(fill="x", pady=(10, 5))

    ttk.Label(intervalo_frame, text="Executar a cada (minutos):").pack(side="left", padx=(0, 5))
    entry_intervalo = ttk.Entry(intervalo_frame, width=10)
    entry_intervalo.insert(0, "60") # Valor padrão de 60 minutos
    entry_intervalo.pack(side="left", padx=(0, 10))
    
    # Checkbox para ativar/desativar o agendamento
    var_intervalo_ativado = tk.BooleanVar(value=False) # Por padrão, o agendamento fica ativo
    ttk.Checkbutton(intervalo_frame, text="Ativar agendamento", variable=var_intervalo_ativado, bootstyle="round-toggle").pack(side="left")


    # Seletor de Responsáveis
    resp_frame = ttk.LabelFrame(main_frame, text="Gerar PDF para:", padding=15)
    resp_frame.pack(fill="both", expand=True, pady=10)

    def selecionar_todos():
        for var in responsaveis_vars.values():
            var.set(True)

    def limpar_selecao():
        for var in responsaveis_vars.values():
            var.set(False)

    btn_frame = ttk.Frame(resp_frame)
    btn_frame.pack(fill="x", pady=(0, 5))
    ttk.Button(btn_frame, text="Selecionar Todos", command=selecionar_todos, bootstyle="info").pack(side="left", fill="x", expand=True, padx=(0, 5))
    ttk.Button(btn_frame, text="Limpar Seleção", command=limpar_selecao, bootstyle="secondary").pack(side="left", fill="x", expand=True, padx=(5, 0))
    
    # Criação dos checkboxes
    for nome, var in responsaveis_vars.items():
        ttk.Checkbutton(resp_frame, text=nome, variable=var, bootstyle="info-round-toggle").pack(anchor="w", pady=2)

    # Barra de progresso e status
    status_label = ttk.Label(main_frame, text="Aguardando...", font=("Helvetica", 10), foreground="gray")
    status_label.pack(pady=(10, 5))

    progress_bar = ttk.Progressbar(main_frame, mode="determinate", bootstyle="info")
    progress_bar.pack(fill="x", pady=(0, 10))

    # Botões de ação
    action_frame = ttk.Frame(main_frame)
    action_frame.pack(fill="x", pady=10)

    ttk.Button(action_frame, text="Executar", command=lambda: executar_script(), bootstyle="success").pack(side="left", fill="x", expand=True, padx=(0, 5))
    ttk.Button(action_frame, text="Pausar", command=cancelar_execucao, bootstyle="danger").pack(side="left", fill="x", expand=True, padx=(5, 0))

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

# --- Funções de Execução Principal ---
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
    
    # Inicia a barra de progresso antes de iniciar o procedimento
    atualizar_progresso("Iniciando a automação...", step=0, total_steps=8)
    
    threading.Thread(target=executar_procedimento, args=(usuario, senha), daemon=True).start()

def executar_procedimento(usuario, senha):
    global driver, execucao_ativa, responsaveis_vars
    execucao_ativa = True
    
    # Definição das etapas para a barra de progresso
    TOTAL_STEPS = 8
    last_valid_interval = 60 # Valor padrão
    
    while execucao_ativa:
        # Coleta os responsáveis e as configurações da UI a cada novo ciclo
        selected_responsaveis = [nome for nome, var in responsaveis_vars.items() if var.get()]
        intervalo_ativado = var_intervalo_ativado.get()
        
        try:
            intervalo_minutos = int(entry_intervalo.get())
            if intervalo_minutos <= 0:
                print(f"⚠️ Intervalo inválido ({intervalo_minutos}). Usando o último valor válido: {last_valid_interval} min.")
                intervalo_minutos = last_valid_interval
            else:
                last_valid_interval = intervalo_minutos
        except (ValueError, tk.TclError):
            print(f"⚠️ Erro ao ler o intervalo. Usando o último valor válido: {last_valid_interval} min.")
            intervalo_minutos = last_valid_interval
            
        print(f"\n--- Iniciando novo ciclo ---")
        print(f"Responsáveis selecionados para este ciclo: {', '.join(selected_responsaveis) or 'Nenhum'}")
        print(f"Intervalo configurado: {intervalo_minutos} minutos. Agendamento Ativado: {'Sim' if intervalo_ativado else 'Não'}")

        driver = None
        df_final = None 

        try:
            # Definição de caminhos
            diretorio_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
            excel_output = os.path.join(script_dir, "Contingencia - Final.xlsx")
            base_responsaveis_path = os.path.join(script_dir, "Base - Respon.xlsx")
            pdf_output_dir = os.path.join(script_dir, "PDF_Saida")

            xpaths = {
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
            
            atualizar_progresso("Iniciando driver...", step=1, total_steps=TOTAL_STEPS)
            driver = iniciar_driver(headless=False)
            
            atualizar_progresso("Realizando login...", step=2, total_steps=TOTAL_STEPS)
            login_usuario(driver, url, usuario, senha, xpaths)
            
            if not execucao_ativa: break
            
            atualizar_progresso("Exportando tabela...", step=3, total_steps=TOTAL_STEPS)
            exportar_tabela(driver, xpaths)
            
            if not execucao_ativa: break

            atualizar_progresso("Aguardando download do CSV...", step=4, total_steps=TOTAL_STEPS)
            df_final = processar_csv(diretorio_downloads, excel_output, base_responsaveis_path, pdf_output_dir, selected_responsaveis)
            
            if df_final is not None:
                if not execucao_ativa: break
                atualizar_progresso("Formatando Excel...", step=5, total_steps=TOTAL_STEPS)
                formatar_excel(excel_output)
                
                if not execucao_ativa: break
                atualizar_progresso("Gerando PDFs...", step=6, total_steps=TOTAL_STEPS)
                salvar_pdf_por_responsavel(df_final, pdf_output_dir)

                if not execucao_ativa: break
                atualizar_progresso("Atualizando planilha de controle...", step=7, total_steps=TOTAL_STEPS)
                atualizar_coleta_planilha(df_final)

                atualizar_progresso("Procedimento concluído com sucesso!", step=8, total_steps=TOTAL_STEPS)
            else:
                atualizar_progresso("Processamento de dados falhou.", step=0, total_steps=1)
        
        except Exception as e:
            print(f"❌ Erro fatal na execução do procedimento: {e}")
            atualizar_progresso(f"Erro: {e}", step=0, total_steps=1)
        finally:
            if driver:
                driver.quit()

        # Decide se continua o loop ou encerra
        if not intervalo_ativado:
            print("Execução única concluída, pois o agendamento está desativado.")
            break # Encerra o loop principal

        # Pausa para o intervalo definido, com verificação de cancelamento e contagem regressiva
        if execucao_ativa:
            tempo_total_espera = intervalo_minutos * 60
            print(f"✅ Execução concluída. Aguardando {intervalo_minutos} minutos para a próxima rodada...")

            for segundos_restantes in range(tempo_total_espera, 0, -1):
                if not execucao_ativa: break
                
                minutos, segundos = divmod(segundos_restantes, 60)
                texto_tempo = f"Próxima execução em {minutos:02d}:{segundos:02d}"
                atualizar_progresso(texto_tempo, step=8, total_steps=TOTAL_STEPS) # Mantém a barra cheia
                time.sleep(1)
            
            if not execucao_ativa: break
    
    execucao_ativa = False
    atualizar_progresso("Procedimento finalizado.", step=0, total_steps=1)
    print("🏁 Procedimento finalizado.")


if __name__ == "__main__":
    driver = None
    criar_interface()

