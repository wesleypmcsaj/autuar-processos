#Em 03/12/24 daqui para baixo contém os códigos de 0 até 5 para testes

from datetime import datetime
import os
import time
from pathlib import Path
import shutil
import re
import fitz  # PyMuPDF
from PyPDF2 import PdfReader, PdfWriter
import json
import threading
import win32print
import win32api

from reportlab.pdfgen import canvas
from io import BytesIO
import gspread  # Para acessar o Google Planilhas
from oauth2client.service_account import ServiceAccountCredentials  # Para autenticação com conta de serviço
from openpyxl import load_workbook
import openpyxl  # Para manipular o Excel sem alterar a formatação
from PyPDF2 import PdfMerger
from send2trash import send2trash  # Biblioteca para mover arquivos para a lixeira
import pandas as pd
from docxtpl import DocxTemplate

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#Código 0, copiar dados do Google Planilhas


print("Copiando os dados disponíveis da planilha Google, aguarde.")

# Configuração da autenticação
def autenticar_google():
    # Define o escopo para acessar o Google Drive e Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

    # Carrega as credenciais do arquivo JSON da conta de serviço
    creds = ServiceAccountCredentials.from_json_keyfile_name('chave.json', scope)
    cliente = gspread.authorize(creds)
    return cliente

# Função principal para copiar dados do Google Planilhas para o Excel
def copiar_dados_para_excel():
    # Autentica e conecta ao Google Planilhas
    cliente = autenticar_google()

    # Abre a planilha pelo ID e seleciona a aba "CAPTACOES"
    planilha = cliente.open_by_key("16_zlC5bRdyGTqFcVFvRIBCYzP-fjoPN9i64tD5DGe5c")
    aba = planilha.worksheet("CAPTACOES")

    # Define os cabeçalhos esperados, caso algum esteja duplicado ou em branco
    expected_headers = ["ID processo", "Procurador", "Orgão"]
    dados = aba.get_all_records(expected_headers=expected_headers)

    # Abre o arquivo Excel existente sem alterar a formatação
    arquivo_excel = "dados_para_autuar_processos.xlsx"
    workbook = openpyxl.load_workbook(arquivo_excel)
    sheet = workbook["CAPTACOES"]  # Nome da aba onde os dados serão adicionados

    # Limpa as linhas de dados antigas (sem apagar o cabeçalho)
    sheet.delete_rows(2, sheet.max_row)

    # Insere os dados a partir da segunda linha
    for i, linha in enumerate(dados, start=2):  # Começa na linha 2 para manter o cabeçalho original
        sheet.cell(row=i, column=1, value=linha["ID processo"])
        sheet.cell(row=i, column=2, value=linha["Procurador"])
        sheet.cell(row=i, column=3, value=linha["Orgão"])

    # Salva o arquivo Excel com os dados atualizados
    workbook.save(arquivo_excel)
    print("Dados copiados com sucesso do Google Planilhas para o Excel.")

# Executa a função para copiar os dados
copiar_dados_para_excel()



#Código 1 inserir livro SAJ nas células correspondentes



def gerar_registros():
    # Cria o arquivo de registros com o padrão necessário
    with open("livro_de_registros_pa_procuradores_judiciais.txt", "w") as arquivo:
        for livro in range(19, 51):  # Começa no Livro 19 e vai até o Livro 50
            for pagina in range(1, 101):  # Vai da página 1 até a 100 (verso incluído)
                # Formata cada linha como "L <livro> P <pagina>, (utilizável)" e duplica a entrada
                arquivo.write(f"L {livro} P {pagina},(utilizável)\n")
                arquivo.write(f"L {livro} P {pagina},(utilizável)\n")  # Duplicação da página com "(utilizável)"
                arquivo.write(f"L {livro} P {pagina} (v),(utilizável)\n")
                arquivo.write(f"L {livro} P {pagina} (v),(utilizável)\n")  # Duplicação da página com "(v),(utilizável)"
    print("Arquivo de registros criado com sucesso.")


def verificar_e_preencher_excel():
    # Verifica se o arquivo "livro_de_registros_pa_procuradores_judiciais.txt" existe; se não, gera o arquivo com o padrão especificado
    if not os.path.exists("livro_de_registros_pa_procuradores_judiciais.txt"):
        print("Arquivo 'livro_de_registros_pa_procuradores_judiciais.txt' não encontrado. Gerando arquivo...")
        gerar_registros()

    # Abre o arquivo de Excel e a aba relevante
    arquivo_excel = "dados_para_autuar_processos.xlsx"
    workbook = openpyxl.load_workbook(arquivo_excel)
    sheet = workbook["CAPTACOES"]

    # Lê o arquivo de registros
    with open("livro_de_registros_pa_procuradores_judiciais.txt", "r") as arquivo:
        registros = arquivo.readlines()

    # Itera sobre as linhas da planilha do Excel para preencher "Livro saj" com registros disponíveis
    for row in range(2, sheet.max_row + 1):
        numero_processo = sheet.cell(row=row, column=1).value
        procurador = sheet.cell(row=row, column=2).value

        # Verifica se "numero processo" e "Procurador" estão preenchidos
        if numero_processo and procurador:
            # Encontra o primeiro registro com "(utilizável)"
            registro_disponivel = None
            for i, linha in enumerate(registros):
                if "(utilizável)" in linha:
                    registro_disponivel = linha.split(",")[0]
                    registros[i] = linha.replace("(utilizável)", "(já utilizado)")
                    break

            # Se não houver registros disponíveis, gera mais 50 livros e reinicia o preenchimento
            if not registro_disponivel:
                print("Todos os registros foram utilizados. Gerando mais 50 livros.")
                with open("livro_de_registros_pa_procuradores_judiciais.txt", "a") as arquivo:
                    for livro in range(51, 101):
                        for pagina in range(1, 101):
                            arquivo.write(f"L {livro} P {pagina},(utilizável)\n")
                            arquivo.write(f"L {livro} P {pagina},(utilizável)\n")
                            arquivo.write(f"L {livro} P {pagina} (v),(utilizável)\n")
                            arquivo.write(f"L {livro} P {pagina} (v),(utilizável)\n")
                return verificar_e_preencher_excel()

            # Insere o registro disponível na coluna "Livro saj" (sexta coluna)
            sheet.cell(row=row, column=6, value=registro_disponivel)

    # Atualiza o arquivo de registros
    with open("livro_de_registros_pa_procuradores_judiciais.txt", "w") as arquivo:
        arquivo.writelines(registros)

    # Salva o arquivo Excel com as alterações
    workbook.save(arquivo_excel)
    print("Livros de registros dos Pa's da procuradoria com sucesso!")


# Gera o arquivo de registros na primeira execução, se necessário
verificar_e_preencher_excel()




#Código 2 baixar arquivos do ESAJ e inicio de completar informações nescessarias;


# Função para atualizar o WebDriver e garantir que está atualizado
def update_webdriver():
    print("Atualizando o WebDriver e completando os dados da tabela. Por favor, aguarde.")
    service = Service(ChromeDriverManager().install())
    return service

# Função para realizar o login no site
def realizar_login(driver):
    with open("login_esaj_TJSP.txt", "r") as file:
        login = file.readline().strip()
        senha = file.readline().strip()

    driver.get(
        "https://esaj.tjsp.jus.br/sajcas/login?service=https%3A%2F%2Fesaj.tjsp.jus.br%2Fesaj%2Fj_spring_cas_security_check"
    )

    driver.find_element(By.XPATH, '//*[@id="usernameForm"]').send_keys(login)
    driver.find_element(By.XPATH, '//*[@id="passwordForm"]').send_keys(senha)
    driver.find_element(By.XPATH, '//*[@id="pbEntrar"]').click()
    time.sleep(1)

# Função para acessar o site de consulta
def acessar_site(driver):
    driver.get("https://esaj.tjsp.jus.br/cpopg/open.do")
    time.sleep(1)

# Função para fechar janelas extras, mantendo apenas a janela principal aberta
def fechar_janelas_extras(driver, janela_principal):
    for handle in driver.window_handles:
        if handle != janela_principal:
            driver.switch_to.window(handle)
            driver.close()
    driver.switch_to.window(janela_principal)

# Caminho da pasta de downloads
downloads_path = Path.home() / "Downloads"

# Função para mover o arquivo baixado para a pasta "docs PJs" e salvar o caminho em um bloco de notas
def mover_arquivo_downloads(destino_pasta):
    docs_pjs_path = Path(destino_pasta)

    # Verifica se a pasta "docs PJs" existe, se não, cria
    docs_pjs_path.mkdir(parents=True, exist_ok=True)

    # Filtra apenas arquivos PDF com a data atual no diretório de downloads
    today = datetime.now().date()
    pdf_files_today = [
        f for f in downloads_path.glob('*.pdf')
        if datetime.fromtimestamp(f.stat().st_mtime).date() == today
    ]

    # Se houver arquivos PDF da data atual, move o mais recente para "docs PJs"
    if pdf_files_today:
        latest_file = max(pdf_files_today, key=os.path.getmtime)
        shutil.move(str(latest_file), str(docs_pjs_path / latest_file.name))

    # Salva o caminho da pasta "docs PJs" em um bloco de notas
    with open("diretorio_docs_pjs.txt", "w") as file:
        file.write(str(docs_pjs_path))

# Função para consultar o processo e salvar as informações no Excel
def consultar_processo(driver, id_processo, row_index, sheet):
    if id_processo:
        match = re.match(r"(\d{7})-(\d{2})\.(\d{4})\.8\.26\.(\d{4})", id_processo)
        if match:
            numero_digito_ano = f"{match.group(1)}-{match.group(2)}.{match.group(3)}"
            foro_numero_unificado = match.group(4)

            driver.get("https://esaj.tjsp.jus.br/cpopg/open.do?gateway=true")
            time.sleep(1)

            # Seleciona a opção "Número do Processo" no dropdown de pesquisa
            select_element = driver.find_element(By.XPATH, '//*[@id="cbPesquisa"]')
            select = Select(select_element)
            select.select_by_visible_text("Número do Processo")

            # Preenche os campos com os valores formatados
            driver.find_element(By.XPATH, '//*[@id="numeroDigitoAnoUnificado"]').send_keys(numero_digito_ano)
            driver.find_element(By.XPATH, '//*[@id="foroNumeroUnificado"]').send_keys(foro_numero_unificado)

            # Clica no botão de consultar processo
            driver.find_element(By.XPATH, '//*[@id="botaoConsultarProcessos"]').click()
            time.sleep(10)

            # Classe do processo
            classe_texto = driver.find_element(By.XPATH, '//*[@id="classeProcesso"]').text
            sheet.cell(row=row_index, column=4, value=classe_texto)

            # Assunto do processo
            assunto_texto = driver.find_element(By.XPATH, '//*[@id="assuntoProcesso"]').text
            sheet.cell(row=row_index, column=5, value=assunto_texto)

            # Nome da parte
            nome_parte_completo = driver.find_element(By.XPATH,
                                                      '//*[@id="tablePartesPrincipais"]/tbody/tr[1]/td[2]').text
            if "Advogada:" in nome_parte_completo:
                nome_parte = nome_parte_completo.split("Advogada:")[0].strip()
            elif "Advogado:" in nome_parte_completo:
                nome_parte = nome_parte_completo.split("Advogado:")[0].strip()
            else:
                nome_parte = nome_parte_completo
            sheet.cell(row=row_index, column=7, value=nome_parte)

            # Salva o arquivo atualizado
            sheet.parent.save("dados_para_autuar_processos.xlsx")

            try:
                # Abre a pasta de documentos do processo
                driver.find_element(By.XPATH, '//*[@id="linkPasta"]').click()
                WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)

                janela_principal = driver.current_window_handle

                # Alterna para a nova janela (pop-up)
                for handle in driver.window_handles:
                    if handle != driver.current_window_handle:
                        driver.switch_to.window(handle)
                        break

                # Clica no botão "Selecionar todos os documentos"
                driver.find_element(By.XPATH, '//*[@id="selecionarButton"]').click()
                time.sleep(3)

                # Clica no botão "Salvar"
                driver.find_element(By.XPATH, '//*[@id="salvarButton"]').click()
                time.sleep(3)

                # Aguarda até que o botão "Continuar" esteja visível e clicável e clica nele
                continuar_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="botaoContinuar"]'))
                )
                continuar_button.click()
                time.sleep(10)

                # Aguarda até que o botão de download esteja disponível e clica nele
                download_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="btnDownloadDocumento"]'))
                )
                download_button.click()
                time.sleep(5)

                # Fecha a janela pop-up após o download
                driver.close()
                driver.switch_to.window(janela_principal)

                # Espera até que o arquivo .crdownload seja removido
                while any(downloads_path.glob("*.crdownload")):
                    time.sleep(1)

                mover_arquivo_downloads("docs PJs")

            except Exception as e:
                print(f"Erro durante o processo de download: {e}")

        else:
            print(f"Aviso: ID do processo '{id_processo}' não está no formato esperado.")
    else:
        print("Aviso: O ID do processo está vazio. Linha ignorada.")
# Função principal para processar todos os dados
def processar_dados():
    wb = load_workbook("dados_para_autuar_processos.xlsx")
    sheet = wb["CAPTACOES"]

    chrome_options = Options()
    service = update_webdriver()
    driver = webdriver.Chrome(service=service, options=chrome_options)

    realizar_login(driver)

    for i, row in enumerate(sheet.iter_rows(min_row=2, max_col=8, values_only=True), start=2):
        id_processo, procurador, orgao = row[0], row[1], row[2]

        # Verificação se "ID processo" começa com "///"
        if id_processo and id_processo.startswith("///"):
            # Extração e divisão das informações
            partes = id_processo[3:].strip().split(" / ")
            if len(partes) >= 4:
                numero_processo, classe, assunto, nome_parte = partes[:4]

                # Preenche as colunas específicas
                sheet.cell(row=i, column=4, value=classe)
                sheet.cell(row=i, column=5, value=assunto)
                sheet.cell(row=i, column=7, value=nome_parte)

                # Preenche a 8ª coluna com a formatação solicitada
                if procurador:
                    informacoes_completas = (
                        f"Ação envolvendo {nome_parte} e PREFEITURA MUNICIPAL DE CARAPICUIBA {classe} {assunto} {numero_processo}"
                    )
                else:
                    informacoes_completas = (
                        f"Solicitação de providências/ solicitação de informações envolvendo {nome_parte} e PREFEITURA MUNICIPAL DE CARAPICUIBA {classe} {assunto} {numero_processo}"
                    )
                sheet.cell(row=i, column=8, value=informacoes_completas)

            # Pula o processamento no Selenium para essas linhas
            print(f"Linha {i} iniciada com '///' foi processada sem usar o Selenium.")
            continue

        # Apenas processa no Selenium se "ID processo" não começar com "///"
        if orgao == "TJSP" and id_processo:
            consultar_processo(driver, id_processo, i, sheet)

    # Salva as alterações no arquivo Excel
    wb.save("dados_para_autuar_processos.xlsx")
    driver.quit()

# Chama a função principal
if __name__ == "__main__":
    processar_dados()


# Código 3 - Finalizar as informações nescessarias para autuar os processos no GIAP



# Mensagem inicial
print("As informações necessárias para autuar os processos administrativos Judiciais estão sendo completadas, por favor aguarde.")

# Carregar o arquivo Excel
arquivo_excel = "dados_para_autuar_processos.xlsx"
workbook = openpyxl.load_workbook(arquivo_excel)
planilha = workbook.active

# Identificar os índices das colunas (usando os nomes fornecidos)
coluna_id_processo = 1  # "ID processo"
coluna_procurador = 2  # "Procurador"
coluna_orgao = 3  # "Orgão"
coluna_classe = 4  # "Classe"
coluna_assunto = 5  # "Assunto"
coluna_livro_saj = 6  # "Livro saj"
coluna_nome_parte = 7  # "Nome da parte"
coluna_info_completa = 8  # "Informacoes completas para autuar o PA"

# Função para extrair o número do processo judicial
def extrair_numero_processo(id_processo):
    if id_processo is None:
        return ""
    id_processo = str(id_processo)
    if "///" in id_processo:
        partes = id_processo.split("/")
        for parte in partes:
            if parte and len(parte) >= 25 and "-" in parte and "." in parte:
                return parte
    elif len(id_processo) >= 25 and "-" in id_processo and "." in id_processo:
        return id_processo
    return ""

# Iterar pelas linhas (ignorar a primeira linha, que é o cabeçalho)
for linha in range(2, planilha.max_row + 1):
    procurador = planilha.cell(row=linha, column=coluna_procurador).value
    orgao = planilha.cell(row=linha, column=coluna_orgao).value
    nome_parte = planilha.cell(row=linha, column=coluna_nome_parte).value
    classe = planilha.cell(row=linha, column=coluna_classe).value
    assunto = planilha.cell(row=linha, column=coluna_assunto).value
    id_processo = planilha.cell(row=linha, column=coluna_id_processo).value
    livro_saj = planilha.cell(row=linha, column=coluna_livro_saj).value

    numero_processo = extrair_numero_processo(id_processo)  # Extrair número do processo judicial

    # Adicionar espaço simples e valor da coluna "Livro saj" (ou apenas espaço se vazio)
    complemento_livro_saj = f" {livro_saj}" if livro_saj else " "

    # Condicional para quando o "Orgão" for "TJSP"
    if orgao == "TJSP":
        if procurador:  # Se o "Procurador" estiver preenchido
            texto = (f"Ação Judicial envolvendo {nome_parte} e PREFEITURA MUNICIPAL DE CARAPICUIBA. "
                     f"{classe}   {assunto}   número do processo: {numero_processo} {complemento_livro_saj}")
        else:  # Se o "Procurador" estiver vazio
            texto = (f"Solicitação de providências/ Solicitação de Informações envolvendo {nome_parte} "
                     f"e PREFEITURA MUNICIPAL DE CARAPICUIBA   {classe}   {assunto} número do processo: {numero_processo} {complemento_livro_saj}")
    else:  # Quando o "Orgão" for diferente de "TJSP"
        texto = f"Solicitação de providências/ Solicitação de Informações envolvendo {id_processo} {complemento_livro_saj}"

    # Inserir o texto na 8ª coluna ("Informacoes completas para autuar o PA")
    planilha.cell(row=linha, column=coluna_info_completa).value = texto

# Salvar o arquivo Excel atualizado
workbook.save(arquivo_excel)

# Mensagem final
print("Todas as informações necessárias para autuar os processos foram inseridas.")



#Código 4 - dividir os documentos baixados em PDF


# Configuração da pasta de entrada e saída
PASTA_ORIGEM = "docs PJs"
PASTA_SAIDA = "docs processados"

# Certifique-se de que a pasta de saída existe
os.makedirs(PASTA_SAIDA, exist_ok=True)

# Dicionário de termos de busca
TERMINOS_PETICAO = [
    "Nestes termos p. deferimento",
    "Termos em que, Pede Deferimento.",
    "Termos em que, Pede Deferimento",
    "Termos em que, Pede e Espera Deferimento.",
    "Termosemque, PedeeEsperaDeferimento.",
    "Termosemque, PedeeEsperaDeferimento",
    "Termos em que, Pede e Espera Deferimento",
]
TERMO_DECISAO = "DECISÃO"
TERMO_DESPACHO = "DESPACHO"
TERMO_OFICIO = "OFÍCIO"
TERMO_JUIZ = "Juiz(a) de Direito: Dr(a)."
TERMO_TRIBUNAL = "TRIBUNAL DE JUSTIÇA DO ESTADO DE SÃO PAULO"

# Função para localizar páginas baseadas em termos
def localizar_paginas(doc, termos, termos_adicionais=None, ultima_ocorrencia=False):
    paginas_encontradas = []
    for i in range(len(doc)):
        pagina = doc[i]
        texto = pagina.get_text()

        # Verifica se os termos principais estão na página
        if any(termo in texto for termo in termos):
            if termos_adicionais is None or all(termo in texto for termo in termos_adicionais):
                paginas_encontradas.append(i)

    # Retorna as páginas encontradas ou apenas a última se necessário
    return [paginas_encontradas[-1]] if ultima_ocorrencia and paginas_encontradas else paginas_encontradas

# Processar cada arquivo PDF na pasta
for arquivo in os.listdir(PASTA_ORIGEM):
    if arquivo.endswith(".pdf"):
        caminho_arquivo = os.path.join(PASTA_ORIGEM, arquivo)
        doc = fitz.open(caminho_arquivo)

        try:
            # Localizar páginas da petição inicial
            paginas_peticao = []
            for i, pagina in enumerate(doc):
                texto = pagina.get_text()
                if any(termo in texto for termo in TERMINOS_PETICAO):
                    paginas_peticao = list(range(0, i + 1))  # Da página 0 até a página encontrada
                    break

            # Se nenhuma página de petição for encontrada, pegar as 30 primeiras
            if not paginas_peticao:
                paginas_peticao = list(range(0, min(30, len(doc))))

            # Localizar a última decisão, despacho ou ofício
            paginas_decisao = localizar_paginas(doc, [TERMO_DECISAO], [TERMO_JUIZ], ultima_ocorrencia=True)
            if not paginas_decisao:  # Se nenhuma decisão for encontrada, buscar despachos
                paginas_decisao = localizar_paginas(doc, [TERMO_DESPACHO], [TERMO_JUIZ], ultima_ocorrencia=True)
            if not paginas_decisao:  # Se nenhum despacho for encontrado, buscar ofícios
                paginas_decisao = localizar_paginas(doc, [TERMO_OFICIO], ultima_ocorrencia=True)

            # Extrair as páginas identificadas
            doc_peticao = fitz.open()
            doc_decisao = fitz.open()

            for pagina in paginas_peticao:
                doc_peticao.insert_pdf(doc, from_page=pagina, to_page=pagina)

            for pagina in paginas_decisao:
                doc_decisao.insert_pdf(doc, from_page=pagina, to_page=pagina)

            # Juntar os documentos extraídos
            doc_final = fitz.open()
            doc_final.insert_pdf(doc_peticao)
            doc_final.insert_pdf(doc_decisao)

            # Renomear e salvar o arquivo final
            nome_base = Path(arquivo).stem  # Remove a extensão .pdf
            caminho_saida = os.path.join(PASTA_SAIDA, f"{nome_base}.pdf")
            doc_final.save(caminho_saida)

            print(f"Processado e salvo: {caminho_saida}")

        finally:
            # Fechar todos os documentos para liberar o arquivo original
            doc.close()
            doc_peticao.close()
            doc_decisao.close()
            doc_final.close()

        # Enviar o arquivo original para a lixeira
        shutil.move(caminho_arquivo, Path.home() / ".Trash")



#Código 5 - numeração das folhas dos arquivos .PDF

def carregar_dados_txt(caminho_txt):
    """Lê a sigla e a matrícula de um arquivo .txt"""
    with open(caminho_txt, "r") as f:
        linhas = f.readlines()
    sigla = linhas[0].strip()  # Primeira linha: Sigla
    matricula = linhas[1].strip()  # Segunda linha: Matrícula
    return sigla, matricula


from reportlab.lib.colors import red  # Importar a cor desejada (vermelho)
from reportlab.lib.colors import black  # Importar a cor desejada (preto)


def ajustar_orientacao_e_numerar(input_path, output_path, sigla, matricula, inicio=2, fonte="Helvetica",
                                 tamanho_fonte=12):
    """Ajusta a orientação das páginas para retrato e adiciona numeração personalizada."""
    reader = PdfReader(input_path)
    writer = PdfWriter()

    for numero, pagina in enumerate(reader.pages, start=inicio):
        # Verifica e ajusta a rotação da página
        if pagina.get("/Rotate") in [90, 270]:  # Páginas na horizontal
            pagina.rotate(0)  # Ajusta para retrato

        # Obtém dimensões da página
        largura_pagina = float(pagina.mediabox[2])
        altura_pagina = float(pagina.mediabox[3])

        # Buffer para gerar a numeração com ReportLab
        buffer = BytesIO()
        c = canvas.Canvas(buffer, pagesize=(largura_pagina, altura_pagina))

        # Define a posição para o canto superior direito
        texto_numero = f"Folha {numero}"
        texto_sigla = f"{sigla}"
        texto_matricula = f"{matricula}"

        margem_direita = 120  # Distância da borda direita
        margem_topo = 30  # Distância do topo

        # Adiciona o texto
        c.setFont(fonte, tamanho_fonte)
        c.setFillColor(red)  # Define a cor como vermelho
        c.drawString(largura_pagina - margem_direita, altura_pagina - margem_topo, texto_numero)
        c.setFont(fonte, tamanho_fonte - 2)
        c.drawString(largura_pagina - margem_direita, altura_pagina - margem_topo - 15, texto_sigla)
        c.drawString(largura_pagina - margem_direita, altura_pagina - margem_topo - 30, texto_matricula)

        c.save()
        buffer.seek(0)

        num_page = PdfReader(buffer).pages[0]
        pagina.merge_page(num_page)
        writer.add_page(pagina)

    # Salva o arquivo ajustado com numeração
    with open(output_path, "wb") as f:
        writer.write(f)


def processar_pasta_docs(
        pasta_entrada="docs processados",
        pasta_saida="docs numerados",
        caminho_txt="sigla_para_a_numeracao.txt",
        inicio=2,
        fonte="Helvetica",
        tamanho_fonte=12
):
    """Processa todos os PDFs da pasta de entrada e os salva numerados na pasta de saída"""
    sigla, matricula = carregar_dados_txt(caminho_txt)

    # Garante que a pasta de saída exista
    os.makedirs(pasta_saida, exist_ok=True)

    for arquivo in os.listdir(pasta_entrada):
        if arquivo.endswith(".pdf"):
            input_path = os.path.join(pasta_entrada, arquivo)
            output_path = os.path.join(pasta_saida, arquivo)
            print(f"Numerando: {arquivo}")
            ajustar_orientacao_e_numerar(input_path, output_path, sigla, matricula, inicio, fonte, tamanho_fonte)


# Personalize os parâmetros aqui
processar_pasta_docs(
    pasta_entrada="docs processados",
    pasta_saida="docs numerados",
    caminho_txt="sigla_para_a_numeracao.txt",  # Caminho para o arquivo .txt com a sigla e matrícula
    inicio=2,  # Início da numeração
    fonte="Helvetica-Bold",  # Fonte para o texto
    tamanho_fonte=10  # Tamanho da fonte principal
)




#Código 6 autuar os processos no GIAP;


# Função para carregar credenciais de login
def carregar_credenciais(caminho_json):
    with open(caminho_json, 'r') as arquivo:
        return json.load(arquivo)


# Dicionário de siglas
dicionario_orgao = {
    "TJSP": "TRIBUNAL DE JUSTIÇA DO ESTADO DE SÃO PAULO",
    "DEFENS. PÚBLICA de SP": "DEFENSORIA PUBLICA DO ESTADO DE SAO PAULO",
    "TRT 2": "TRIBUNAL REGIONAL DO TRABALHO DA 2A REGIAO",
    # Adicione outras siglas conforme necessário
}

# Configurar WebDriver
def configurar_webdriver():
    options = webdriver.ChromeOptions()
    # Diretório onde o arquivo PDF será salvo
    caminho_para_salvar_pdf = r"C:\Users\wesley\PycharmProjects\Autuar-processos\docs numerados"

    # Configurações para salvar como PDF
    settings = {
        "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": ""
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2,
        "isHeaderFooterEnabled": False,
        "isCssBackgroundEnabled": True,
    }

    prefs = {
        "printing.print_preview_sticky_settings.appState": json.dumps(settings),
        "savefile.default_directory": caminho_para_salvar_pdf,
    }

    options.add_experimental_option("prefs", prefs)
    options.add_argument('--kiosk-printing')  # Imprime automaticamente sem abrir a janela de impressão

    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


def renomear_capa_pdf(diretorio, texto_busca):
    for arquivo in os.listdir(diretorio):
        if arquivo.endswith(".pdf"):
            caminho_pdf = os.path.join(diretorio, arquivo)
            with fitz.open(caminho_pdf) as pdf:
                texto = ""
                for pagina in pdf:
                    texto += pagina.get_text()

                # Procurar o identificador único no texto
                if texto_busca in texto:
                    inicio = texto.find(texto_busca) + len(texto_busca)
                    fim = texto.find("\n", inicio)  # Encontra o final da linha
                    numero_processo = texto[inicio:fim].strip()

                    # Substituir "/" por "-" no nome do arquivo
                    numero_processo = numero_processo.replace("/", "-")

                    # Criar o novo nome para o arquivo
                    novo_nome = f"capa_ref_pa_{numero_processo}.pdf"
                    caminho_novo = os.path.join(diretorio, novo_nome)

            #Esse time sleep é para dar tempo de renomear o arquivo de capa
            time.sleep(0.5)
            # Fechando o arquivo antes de renomeá-lo
            os.rename(caminho_pdf, caminho_novo)
            print(f"Arquivo renomeado para: {novo_nome}")
            return  # Parar após renomear o arquivo correto


# Função principal
def processar_planilha(caminho_excel, caminho_credenciais):
    # Carregar planilha
    workbook = openpyxl.load_workbook(caminho_excel)
    sheet = workbook.active  # Considera a primeira aba como ativa

    # Carregar credenciais
    credenciais = carregar_credenciais(caminho_credenciais)
    usuario = credenciais['usuario']
    senha = credenciais['senha']

    # Iniciar WebDriver
    driver = configurar_webdriver()

    try:
        # Realizar login no GIAP
        driver.get('https://carapicuiba.giap.com.br/apex/carapi/f?p=652:LOGIN')
        driver.find_element(By.ID, 'P101_USERNAME').clear()
        driver.find_element(By.ID, 'P101_USERNAME').send_keys(usuario)
        driver.find_element(By.ID, 'P101_PASSWORD').clear()
        driver.find_element(By.ID, 'P101_PASSWORD').send_keys(senha)
        driver.find_element(By.ID, 'wwvFlowForm').submit()

        time.sleep(1)
        driver.find_element(By.XPATH,
                            '//*[@id="report_R5001749296453489731"]/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/a').click()

        # Iterar pelas linhas da planilha (ignorando o cabeçalho)
        for row_index, row in enumerate(sheet.iter_rows(min_row=2), start=2):
            orgao = row[2].value  # Coluna "Orgão" (índice 2, base 0)
            informacoes = row[7].value  # Coluna "Informações completas para autuar o PA"

            if not orgao:  # Interrompe ao encontrar linha vazia
                break

                # Buscar o significado da sigla no dicionário
            significado = dicionario_orgao.get(orgao, "Sigla não encontrada")

            #BTN de clicar na guia de processo
            driver.find_element(By.XPATH,
                                '// *[ @ id = "wwvFlowForm"] / div[2] / div / table / tbody / tr / td / div[1] / div[3] / div[1] / img').click()
            #Aqui está no menu principal, dai esse botão é para clicar em "Abertura de Processo"
            driver.find_element(By.XPATH,
                                '//*[@id="R5002551580972218898"]/tbody/tr[2]/td/ol/li[1]/a').click()

            #Aqui está faltando o trecho até chegar ao campo de busca
            # Preencher dados no GIAP
            time.sleep(2)
            campo_de_busca_nome_orgao = driver.find_element(By.XPATH, '//*[@id="P52_DSP_RESPONSAVEL"]')
            campo_de_busca_nome_orgao.clear()
            campo_de_busca_nome_orgao.send_keys(significado)
            time.sleep(2)
            campo_de_busca_nome_orgao.submit()
            time.sleep(1)

            # Selecionar opções específicas no GIAP (exemplo)
            if significado == "TRIBUNAL DE JUSTIÇA DO ESTADO DE SÃO PAULO":
                    driver.find_element(By.XPATH,
                                        '//*[@id="report_P52_TIPO_EXPEDIENTE"]/tbody/tr[2]/td/table/tbody/tr[3]/td[10]/a').click()
            elif significado == "DEFENSORIA PUBLICA DO ESTADO DE SAO PAULO":
                    driver.find_element(By.XPATH,
                                        '//*[@id="report_P52_TIPO_EXPEDIENTE"]/tbody/tr[2]/td/table/tbody/tr[2]/td[10]/a').click()
            elif significado == "TRIBUNAL REGIONAL DO TRABALHO DA 2A REGIAO":
                    driver.find_element(By.XPATH,
                                        '//*[@id="report_P52_TIPO_EXPEDIENTE"]/tbody/tr[2]/td/table/tbody/tr[2]/td[10]/a').click()

            dropdown = driver.find_element(By.XPATH, '//*[@id="P51_TIPO_PROCESSO"]')
            select = Select(dropdown)
            select.select_by_visible_text("SEC. MUNICIPAL DE ASSUNTOS JURÍDICOS")
            time.sleep(1)

            # Preencher informações no campo de texto
            driver.find_element(By.XPATH, '//*[@id="P51_COD_ASSUNTO1"]').send_keys('37')
            driver.find_element(By.XPATH, '//*[@id="P51_PRCS_DES_PROCESSO"]').send_keys(informacoes)
            time.sleep(1)

            # Clicar no botão "gerar processo"
            driver.find_element(By.XPATH, '//*[@id="B5000650400896932596"]/span').click()

            # Copiar número do PA gerado
            copiar_num_pa = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="apex_layout_5000650194595932588"]/tbody/tr/td[2]'))
            )
            valor_giap_copiado = copiar_num_pa.text

            # Atualizar a célula correspondente na planilha
            sheet.cell(row=row_index, column=9, value=valor_giap_copiado)  # Coluna 9 é "Número PA gerado"

                # Salvar a planilha atualizada
            workbook.save(caminho_excel)

            # Daqui para baixo segue o meu código para  baixar a capa como PDF

            # Salvar a página como PDF
            driver.find_element(By.XPATH, '//*[@id="B4985322504071026678"]/span').click()
            driver.switch_to.window(driver.window_handles[1])
            print("Salvando a página como PDF...")
            driver.execute_script('window.print();')
            time.sleep(5)  # Tempo para salvar o arquivo
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            print("Capa do processo salva como PDF com sucesso!")

            # Renomear a capa do processo
            diretorio_pdfs = r"C:\Users\wesley\PycharmProjects\Autuar-processos\docs numerados"
            renomear_capa_pdf(diretorio_pdfs, "INFORMAÇÕES DO PROCESSO - ")

            #Clicar na engrenagem para aparecer o menu para voltar para a tela principal do GIAP
            driver.find_element(By.XPATH, '//*[@id="menu_app"]').click()

            #Depois de clicar na engranagem, vai clicar no botão  chamado de "Processo" para voltar para a tela principal do GIAP
            driver.find_element(By.XPATH, '// *[ @ id = "aparece_app"] / div[2] / table / tbody / tr / td / div[2] / div / div / div[1] / div[1] / img').click()

    finally:
        driver.quit()


# Caminhos dos arquivos
caminho_excel = "dados_para_autuar_processos.xlsx"
caminho_credenciais = "credenciais_login_GIAP.json"

# Executar
processar_planilha(caminho_excel, caminho_credenciais)


#Código 7 - Juntar capa com os documentos em si


# Função para verificar se um número segue o padrão especificado
def verificar_padrao_numero(numero):
    padrao = r"\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}"
    return bool(re.match(padrao, numero))


# Função para processar o diretório e os arquivos do Excel
def processar_documentos(diretorio_docs, arquivo_excel):
    # Carregar o arquivo Excel e a guia 'CAPTACOES'
    wb = load_workbook(arquivo_excel)
    ws = wb["CAPTACOES"]

    # Iterar sobre as linhas da planilha (ignorando o cabeçalho)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in row):
            break  # Parar se encontrar uma linha completamente vazia

        orgao = row[2]  # Coluna "Orgão"
        id_processo = row[0]  # Coluna "ID processo"
        numero_pa_gerado = row[8]  # Coluna "Numero PA gerado"

        # Verificar as condições da linha
        if orgao == "TJSP" and id_processo and verificar_padrao_numero(id_processo):
            # Tratar o valor de 'numero_pa_gerado'
            numero_processo = numero_pa_gerado.replace("/", "-")
            nome_capa = f"capa_ref_pa_{numero_processo}.pdf"
            caminho_capa = os.path.join(diretorio_docs, nome_capa)
            caminho_documento = os.path.join(diretorio_docs, f"{id_processo}.pdf")

            # Verificar se a capa e o documento existem
            if os.path.exists(caminho_capa) and os.path.exists(caminho_documento):
                # Criar o nome do arquivo combinado
                nome_saida = f"{id_processo}_completo.pdf"
                caminho_saida = os.path.join(diretorio_docs, nome_saida)

                # Combinar os PDFs
                merger = PdfMerger()
                merger.append(caminho_capa)
                merger.append(caminho_documento)
                merger.write(caminho_saida)
                merger.close()

                # Mover a capa processada para a lixeira
                send2trash(caminho_capa)

                # Mover o documento principal para a lixeira
                send2trash(caminho_documento)

                print(f"Processado: {nome_saida}")
            else:
                print(f"Arquivos ausentes para o ID {id_processo}.")
        elif id_processo and id_processo.startswith("/// "):
            numero_potencial = id_processo[4:].strip()
            if verificar_padrao_numero(numero_potencial):
                # Repetir o mesmo processo para as linhas com "/// "
                numero_processo = numero_pa_gerado.replace("/", "-")
                nome_capa = f"capa_ref_pa_{numero_processo}.pdf"
                caminho_capa = os.path.join(diretorio_docs, nome_capa)
                caminho_documento = os.path.join(diretorio_docs, f"{numero_potencial}.pdf")

                # Verificar se a capa e o documento existem
                if os.path.exists(caminho_capa) and os.path.exists(caminho_documento):
                    nome_saida = f"{numero_potencial}_completo.pdf"
                    caminho_saida = os.path.join(diretorio_docs, nome_saida)

                    # Combinar os PDFs
                    merger = PdfMerger()
                    merger.append(caminho_capa)
                    merger.append(caminho_documento)
                    merger.write(caminho_saida)
                    merger.close()

                    # Mover a capa processada para a lixeira
                    send2trash(caminho_capa)

                    # Mover o documento principal para a lixeira
                    send2trash(caminho_documento)

                    print(f"Processado: {nome_saida}")
                else:
                    print(f"Arquivos ausentes para o número potencial {numero_potencial}.")


# Definir os caminhos
diretorio_docs = "docs numerados"
arquivo_excel = "dados_para_autuar_processos.xlsx"

# Executar o processamento
processar_documentos(diretorio_docs, arquivo_excel)



#Código 8 Criar o modelo de registro para os cadernos SAJ de procurador

# Caminho para o arquivo Excel e o modelo do Word
excel_file = 'dados_para_autuar_processos.xlsx'
sheet_name = 'CAPTACOES'
word_template = 'modelos_pa_procuradoria.docx'

# Carregar o Excel com pandas
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Verificar o DataFrame carregado
print("Cabeçalhos do DataFrame:", df.columns)
print(df.head())  # Mostra as primeiras linhas para confirmar os dados

# Obter a data atual
current_day = datetime.today().day
current_month = datetime.today().month
months = [
    'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun',
    'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'
]
current_month_text = months[current_month - 1]
current_year = datetime.today().year
current_date = f"{current_day} {current_month_text} {current_year}"

# Filtrar as linhas onde a coluna "Procurador" não está vazia
df_filtered = df[df['Procurador'].notna()]

# Verificar o filtro
print("DataFrame filtrado:")
print(df_filtered)
print(f"Quantidade de linhas após filtro: {len(df_filtered)}")

# Preparar os dados para preencher a tabela no Word
table_data = []
for _, row in df_filtered.iterrows():
    numero_pa = row['Numero PA gerado']
    info_completa = row['Informacoes completas para autuar o PA']
    procurador = row['Procurador']
    data_procurador = f"{current_date}  {procurador}"

    table_data.append({
        'Data Atual': current_date,
        'Número PA Gerado': numero_pa,
        'Informações Completas para Autuar o PA': info_completa,
        'Data Atual + Procurador': data_procurador
    })

# Verificar os dados preparados
print("Dados preparados para o Word:")
print(table_data)

# Carregar o modelo do Word
doc = DocxTemplate(word_template)

# Adicionar os dados ao contexto do documento
context = {'table': table_data}

# Renderizar o documento com os dados
doc.render(context)

# Salvar o documento preenchido
output_file = r"docs numerados/modelo_registros_PA_procuradoria_preenchido.docx"
doc.save(output_file)

print(f"Documento gerado com sucesso: {output_file}")



# Código 8.1 - Imprimir arquivos de um diretório
# Escolher qual impressora usar
lista_impressoras = win32print.EnumPrinters(2)
impressora = lista_impressoras[5]

# Mostrar a lista das impressoras conectadas ao PC
#print(lista_impressoras)

win32print.SetDefaultPrinter(impressora[2])

# Mandar imprimir todos os arquivos de uma pasta
caminho = r"C:\Users\wesley\PycharmProjects\Autuar-processos\docs numerados"
lista_arquivos = os.listdir(caminho)

# https://docs.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shellexecutea
for arquivo in lista_arquivos:
    win32api.ShellExecute(0, "print", os.path.join(caminho, arquivo), None, caminho, 0)

print("Arquivos enviados para impressão. Aguardando a conclusão da impressão...")
# Espera para garantir que todos os arquivos sejam impressos
time.sleep(60)  # Ajuste o tempo conforme necessário para garantir que todos os arquivos sejam impressos





#Código 9 atualizar as planilhas de judicial e do expediente

# Função para autenticar usando uma conta de serviço
def autenticar_google_sheets():
    # Caminho para o arquivo JSON das credenciais
    credenciais_json = 'chave.json'

    # Escopos necessários para acessar Google Sheets e Drive
    escopos = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

    # Configuração das credenciais
    credenciais = ServiceAccountCredentials.from_json_keyfile_name(credenciais_json, escopos)
    cliente = gspread.authorize(credenciais)
    return cliente


# Função para encontrar a próxima linha vazia
def encontrar_proxima_linha_vazia(planilha):
    valores = planilha.col_values(2)  # Coluna "Processo Administrativo" como referência
    return len(valores) + 1


# Função para processar o Excel e atualizar Google Sheets
def processar_excel_para_google_planilhas():
    # Caminho do arquivo Excel
    caminho_excel = 'dados_para_autuar_processos.xlsx'

    # Abrindo o arquivo Excel
    workbook = openpyxl.load_workbook(caminho_excel)
    sheet = workbook.active

    # Autenticando e conectando ao Google Sheets
    cliente = autenticar_google_sheets()

    # Acessando a planilha Google para o Caso 1
    planilha_caso_1 = cliente.open_by_key('1B-zoQMhTcQM4mSSKeQdHJj6uidH_q_NfkvkJYrrTMoU').worksheet(
        'Processos Judiciais')

    # Processando as linhas do Excel
    for row in sheet.iter_rows(min_row=2, values_only=True):
        # Verificar se a linha está completamente vazia
        if all(cell is None for cell in row):
            break

        # Dados das colunas
        id_processo = row[0]  # Primeira coluna
        procurador = row[1]  # Segunda coluna
        livro_saj = row[5]  # Sexta coluna
        info_autuar_pa = row[7]  # Oitava coluna
        numero_pa_gerado = row[8]  # Nona coluna

        # Verificar se o ID do processo começa com '///'
        if isinstance(id_processo, str) and id_processo.startswith('///'):
            # Extrair o conteúdo entre as barras
            id_processo = id_processo.split('/')[3]

        # Caso 1: Se "Procurador" estiver preenchido
        if procurador:
            data_atual = datetime.now().strftime('%d/%m/%Y')
            linha_caso_1 = [
                numero_pa_gerado,  # Coluna "PA"
                id_processo,  # Coluna "NÚMERO DE PROCESSO JUDICIAL"
                info_autuar_pa,  # Coluna "ASSUNTO"
                f"{procurador} {data_atual}",  # Coluna "andamento"
                procurador,  # Coluna "PROCURADOR RESPONSÁVEL"
                livro_saj  # Coluna "LIVRO SAJ"
            ]
            planilha_caso_1.append_row(linha_caso_1)

        # Caso 2: Se "Procurador" estiver vazio, pula para a próxima linha
        else:
            continue

    print("Processos inseridos na planilha da procuradoria com sucesso!")


# Função para processar o Caso 2
def processar_caso_2():
    # Caminho do arquivo Excel
    caminho_excel = 'dados_para_autuar_processos.xlsx'

    # Abrindo o arquivo Excel
    workbook = openpyxl.load_workbook(caminho_excel)
    sheet = workbook.active

    # Autenticando e conectando ao Google Sheets
    cliente = autenticar_google_sheets()

    # Acessando a planilha Google do Caso 2
    planilha_caso_2 = cliente.open_by_key('16_zlC5bRdyGTqFcVFvRIBCYzP-fjoPN9i64tD5DGe5c').worksheet('Atual')

    # Processando as linhas do Excel
    for row in sheet.iter_rows(min_row=2, values_only=True):
        # Verificar se a linha está completamente vazia
        if all(cell is None for cell in row):
            break

        # Dados das colunas
        id_processo = row[0]  # Primeira coluna
        procurador = row[1]  # Segunda coluna
        numero_pa_gerado = row[8]  # Nona coluna
        orgao = row[2]  # Terceira coluna (para "Órgão de Destino")

        # Verificar se o ID do processo começa com '///'
        if isinstance(id_processo, str) and id_processo.startswith('///'):
            # Extrair o conteúdo entre as barras
            id_processo = id_processo.split('/')[3]

        # Pular a linha se "Procurador" estiver preenchido
        if procurador:
            continue

        # Encontrar a próxima linha vazia
        proxima_linha = encontrar_proxima_linha_vazia(planilha_caso_2)

        # Atualizar as colunas manualmente
        planilha_caso_2.update_cell(proxima_linha, 2, numero_pa_gerado)  # Coluna "Processo Administrativo"
        planilha_caso_2.update_cell(proxima_linha, 3, id_processo)  # Coluna "Referência"
        planilha_caso_2.update_cell(proxima_linha, 4, orgao)  # Coluna "Órgão de Destino"
        planilha_caso_2.update_cell(proxima_linha, 14, "Aguardando providências")
        # Atualizar a célula com a mensagem e a data
        planilha_caso_2.update_cell(proxima_linha, 15, f"PA autuado em: {datetime.now().strftime('%d/%m/%Y')}")

    print("Processos atualizados com sucesso na planilha de controle de Ofícios")


# Executar os processos
processar_excel_para_google_planilhas()
processar_caso_2()




#Código 10 - Limpar os itens da planilha excel e diretórios para a próxima execução
def limpar_planilha_captacoes():
    def contagem_regressiva():
        for t in range(30, 0, -1):
            print(f"\rAs informações da aba 'CAPTACOES' do arquivo Excel serão apagadas em {t} segundos...", end="")
            time.sleep(1)
        print("\nTempo esgotado. Limpando a aba automaticamente para futuros usos.")
        realizar_limpeza()

    def realizar_limpeza():
        # Caminho do arquivo Excel
        caminho_arquivo = "dados_para_autuar_processos.xlsx"

        # Abrindo o arquivo e selecionando a aba "CAPTACOES"
        workbook = load_workbook(caminho_arquivo)
        planilha = workbook["CAPTACOES"]

        # Iterando sobre as linhas, preservando o cabeçalho
        for row in planilha.iter_rows(min_row=2, min_col=1, max_col=9):  # Colunas A (1) até I (9)
            for cell in row:
                cell.value = None  # Apaga o conteúdo da célula

        # Salvando as modificações
        workbook.save(caminho_arquivo)
        print("As informações da aba 'CAPTACOES' foram limpas para a próxima execução.")

    # Inicia a contagem regressiva em uma nova thread
    thread_timer = threading.Thread(target=contagem_regressiva)
    thread_timer.start()


# Chama a função para execução
limpar_planilha_captacoes()





#Código 10.1 - mover  o conteudo dos diretórios para a lixeira

print("Movendo os arquivos das pastas de PDF's de PJ's para a lixeira.")

def mover_para_lixeira():
    # Lista das pastas que terão seus conteúdos movidos para a lixeira
    pastas = ["docs PJs", "docs numerados", "docs processados"]

    for pasta in pastas:
        # Verifica se a pasta existe
        if not os.path.exists(pasta):
            print(f"A pasta '{pasta}' não foi encontrada.")
            continue

        # Itera sobre os arquivos e subdiretórios na pasta
        for item in os.listdir(pasta):
            caminho_item = os.path.join(pasta, item)

            # Move para a lixeira apenas arquivos ou diretórios válidos
            if os.path.isfile(caminho_item) or os.path.isdir(caminho_item):
                send2trash(caminho_item)
                print(f"Movido para a lixeira: {caminho_item}")

    print("Todos os arquivos e subdiretórios das pastas especificadas foram movidos para a lixeira.")

# Chamar a função
mover_para_lixeira()






