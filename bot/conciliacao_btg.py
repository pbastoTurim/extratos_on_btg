import os
import shutil
import glob
import pandas as pd

import datetime
from openpyxl import Workbook, load_workbook

from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

from modules.boxes_input import input_token_btg

LOG_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'logs.xlsx')

# Configurações do Selenium
URL_SITE = r'https://access.btgpactualdigital.com/login/externo'
URL_EMPRESA = r"https://access.btgpactualdigital.com/login/multiprofile"
SCROLL_DOWN = -500
DIRETORIO_DESTINO_PDFS = r"F:\ExtratosBcosOff\Controle Data Analysis\code\extratos_onshore_btg\documents"
DIRETORIO_DESTINO_PDFS2 = r"F:\ExtratosBcosOff\Controle Data Analysis\code\extratos_onshore_btg\docs_informes"

def get_pasta_mes(mes_ano: str) -> str:
    """
    Retorna o caminho da pasta para o mês/ano específico
    :param mes_ano: mês e ano no formato "MM/YYYY"
    :return: caminho absoluto da pasta
    """
    mes, ano = map(int, mes_ano.split("/"))
    pasta_mes = os.path.join(DIRETORIO_DESTINO_PDFS, f"{ano}_{mes:02d}")
    if not os.path.exists(pasta_mes):
        os.makedirs(pasta_mes)
    return pasta_mes

def verificar_extratos_existentes(mes_ano: str) -> set:
    """
    Verifica quais clientes já têm extratos baixados na pasta do mês
    :param mes_ano: mês e ano no formato "MM/YYYY"
    :return: conjunto com os nomes dos clientes que já têm extratos
    """
    pasta_mes = get_pasta_mes(mes_ano)
    arquivos = glob.glob(os.path.join(pasta_mes, "*.pdf"))
    clientes_processados = set()
    
    for arquivo in arquivos:
        nome_arquivo = os.path.basename(arquivo)
        # O nome do arquivo segue o padrão: CLIENTE-BTG-AAAAMM.pdf
        if " - BTG - " in nome_arquivo:
            cliente = nome_arquivo.split(" - BTG - ")[0].strip()
            clientes_processados.add(cliente)
    
    return clientes_processados

# Caminho do arquivo de log

def log_download(cliente:str, status:str, path:str=None):
    """
    Atualiza o status do download para o cliente especificado no arquivo Excel.
    """
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    if os.path.exists(LOG_PATH):
        wb = load_workbook(LOG_PATH)
        ws = wb.active
        
        # Procura pela linha do cliente e atualiza
        found = False
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if row[1] == cliente:  # Cliente está na coluna B (índice 1)
                ws.cell(row=row_idx, column=1, value=now)         # Timestamp
                ws.cell(row=row_idx, column=3, value=status)      # Status
                ws.cell(row=row_idx, column=4, value=path or "")  # Caminho
                found = True
                break
        
        # Se não encontrar o cliente (situação improvável), adiciona uma nova linha
        if not found:
            ws.append([now, cliente, status, path or ""])
            
        wb.save(LOG_PATH)

def mudar_cliente(url: str, new_client: str) -> str:
    """
    Modifica a URL substituindo apenas o ID do cliente entre 'clients' e 'accounts'
    :param url: URL atual
    :param old_client: ID do cliente atual (não usado mais)
    :param new_client: Novo ID do cliente
    :return: Nova URL com o ID do cliente atualizado
    """

    new_client = "00" + new_client.strip()  # Remove espaços em branco desnecessários
    # Divide a URL em partes usando 'accounts/' como separador
    parte_inicial = url.split('accounts/')[0] + 'accounts/'
    
    # Pega o resto da URL após o ID do cliente
    parte_final = url.split('history/')[1]

    # Pega o número da conta atual
    account_number = parte_final.split('/')[0]
    
    # Monta a nova URL
    new_url = f"{parte_inicial}{new_client}/history/{parte_final}"
    
    return new_url

def close_overlays(driver):
    # Move o mouse para uma área “neutra”
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        ActionChains(driver).move_to_element_with_offset(body, 10, 10).perform()
        ActionChains(driver).move_to_element_with_offset(body, 100, 100).perform()
    except Exception:
        pass


def contas_clientes(path: str) -> dict:
    """
    Método que retorna todas as contas do banco BTG.
    :param path: caminho do arquivo Excel com as contas
    :return: dicionário com as contas, os clientes e a empresa
    :rtype: dict
    """
    try:
        df = pd.read_excel(path, sheet_name='BTG', dtype=str)

        if 'NOME ARQUIVO SE SALVO HOJE' not in df.columns:
            raise ValueError("A coluna 'NOME ARQUIVO SE SALVO HOJE' não foi encontrada no arquivo.")

        if 'FAMILIA' not in df.columns:
            raise ValueError("A coluna 'FAMILIA' não foi encontrada no arquivo.")

        if 'SIGLA' not in df.columns:
            raise ValueError("A coluna 'SIGLA' não foi encontrada no arquivo.")
        
        if 'EMPRESA' not in df.columns:
            raise ValueError("A coluna 'EMPRESA' não foi encontrada no arquivo.")
        
        df['CLIENTE'] = df['FAMILIA'].astype(str) + ' - ' + df['SIGLA'].astype(str)
        
        # Retorna um dicionário com a conta como chave e um tupla (cliente, empresa) como valor
        contas = {}
        for _, row in df.iterrows():
            conta = row['NOME ARQUIVO SE SALVO HOJE']
            cliente = row['CLIENTE']
            empresa = str(row['EMPRESA']).lower().strip()  # Normaliza para minúsculo
            contas[conta] = (cliente, empresa)
        
        return contas
    
    except FileNotFoundError:
        print(f"O arquivo '{path}' não foi encontrado.")
        return {}
    
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return {}

def data_referencia(mes_ano: str) -> tuple:
    mes_atual, ano_atual = map(int, mes_ano.split("/"))

    if mes_atual == 1:
        mes = 12
        ano = ano_atual - 1
    else:
        mes = mes_atual - 1
        ano = ano_atual
    
    mes_str = f"{mes:02d}"
    ano_str = str(ano)

    return mes_str, ano_str
    
def data_informe(mes_ano: str):
    mes_atual, ano_atual = map(int, mes_ano.split("/"))
    ano = ano_atual - 1    
    ano_str = str(ano)
    return ano_str

# Função para renomear o arquivo baixado

def selecionar_perfil(navegador, empresa: str) -> bool:
    """
    Seleciona o perfil correto (TURIM ou TORI) na página multiprofile
    :param navegador: instância do webdriver
    :param empresa: 'turim' ou 'tori'
    :return: True se conseguiu selecionar, False caso contrário
    """
    try:
        escritorio_texto = "TURIM 21 INVESTIMENTOS LTDA" if empresa == 'turim' else "TORI 21 CAPITAL LTDA"
        
        # Aguarda a página multiprofile carregar
        WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, "//span[@class='item-size' and contains(text(), 'TURIM 21 INVESTIMENTOS LTDA') or contains(text(), 'TORI 21 CAPITAL LTDA')]" ))
        )
        sleep(1)
        
        # Encontra o botão "Acessar" do perfil correto
        # Procura pelo div que contém o texto do escritório e depois encontra o botão dentro dele
        botao_perfil = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//div[@class='user-data'][.//span[@class='item-size' and contains(text(), '{escritorio_texto}')]]//button[@type='button']"))
        )
        sleep(0.5)
        botao_perfil.click()
        sleep(1)
        
        print(f"\nPerfil {empresa.upper()} selecionado com sucesso!")
        return True
    except Exception as e:
        print(f"\nErro ao selecionar perfil {empresa.upper()}: {e}")
        return False

def renomear_arquivo(diretorio: str, cliente: str, mes_ano: str) -> str | None:
    try:
        # Localizar o arquivo mais recente no diretório
        arquivos = glob.glob(os.path.join(diretorio, "*"))

        if not arquivos:
            print(f"Nenhum arquivo encontrado no diretório {diretorio}")
            return None

        arquivo_mais_recente = max(arquivos, key=os.path.getctime)

        # Obter a extensão do arquivo
        extensao = os.path.splitext(arquivo_mais_recente)[1]
        mes, ano = data_referencia(mes_ano)
        familia = cliente.split(" - ")[0].strip()

        # Primeiro, salva na pasta do mês para controle
        try:
            pasta_mes = get_pasta_mes(mes_ano)
            novo_nome_mes = f"{cliente} - BTG - {ano}{mes}{extensao}"
            arquivo_mes = os.path.join(pasta_mes, novo_nome_mes)
            shutil.copy2(arquivo_mais_recente, arquivo_mes)
        except Exception as e:
            print(f"Erro ao salvar na pasta de controle: {e}")
            # Continue mesmo se falhar o salvamento na pasta de controle

        # Depois move para a pasta final
        destino_base = r"F:\Extratos_Bcos\ONSHORE"
        pasta_destino = os.path.join(destino_base, ano, f"{ano} {mes}", familia)

        if not os.path.exists(pasta_destino):
            print(f"Diretório {pasta_destino} não existe. Tentando criar...")
            try:
                os.makedirs(pasta_destino, exist_ok=True)
            except Exception as e:
                print(f"Erro ao criar diretório: {e}")
                return None

        novo_nome = f"{cliente} - BTG - {ano}{mes}{extensao}"
        destino_arquivo = os.path.join(pasta_destino, novo_nome)

        try:
            shutil.move(arquivo_mais_recente, destino_arquivo)
            return destino_arquivo
        except Exception as e:
            print(f"Erro ao mover arquivo para destino final: {e}")
            # Se falhou ao mover, tenta pelo menos copiar
            try:
                shutil.copy2(arquivo_mais_recente, destino_arquivo)
                os.remove(arquivo_mais_recente)  # Tenta remover o original
                return destino_arquivo
            except:
                return None

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")
        return None

def renomear_informe(diretorio: str, cliente: str, mes_ano: str) -> str | None:
    try:
        # Localizar o arquivo mais recente no diretório
        arquivos = glob.glob(os.path.join(diretorio, "*"))

        if not arquivos:
            print(f"Nenhum arquivo encontrado no diretório {diretorio}")
            return None

        arquivo_mais_recente = max(arquivos, key=os.path.getctime)

        # Obter a extensão do arquivo
        extensao = os.path.splitext(arquivo_mais_recente)[1]
        ano = data_informe(mes_ano)
        familia = cliente.split(" - ")[0].strip()

        # Move para a pasta final
        destino_base = r"F:\Informes de Rendimentos\Casca BTG 2025" # ALTERAR AQUI
        pasta_destino = os.path.join(destino_base, familia)

        if not os.path.exists(pasta_destino):
            print(f"Diretório {pasta_destino} não existe. Tentando criar...")
            try:
                os.makedirs(pasta_destino, exist_ok=True)
            except Exception as e:
                print(f"Erro ao criar diretório: {e}")
                return None

        novo_nome = f"{cliente} - BTG Informe - {ano}{extensao}"
        destino_arquivo = os.path.join(pasta_destino, novo_nome)

        try:
            shutil.move(arquivo_mais_recente, destino_arquivo)
            return destino_arquivo
        except Exception as e:
            print(f"Erro ao mover arquivo para destino final: {e}")
            # Se falhou ao mover, tenta pelo menos copiar
            try:
                shutil.copy2(arquivo_mais_recente, destino_arquivo)
                os.remove(arquivo_mais_recente)  # Tenta remover o original
                return destino_arquivo
            except:
                return None

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")
        return None

def baixar_posicoes(login: str, senha: str, mes_ano: str) -> None:
    """
    Método que baixa os documentos de imposto de renda do BTG Pactual.
    :param login: login do cliente
    :param senha: senha do cliente
    :return: True se o download for realizado com sucesso, False caso contrário.
    """
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option("prefs", {
        "disable-popup-blocking": True,
        "download.default_directory": DIRETORIO_DESTINO_PDFS,
        "savefile.default_directory": DIRETORIO_DESTINO_PDFS,
    })    

    # Carrega a lista de clientes
    clients = contas_clientes(path=r"F:\ExtratosBcosOff\Controle Data Analysis\controle_extratos_ONSHORE.xlsm")
    clients_list = list(clients.keys())
    
    # Verifica quais clientes já têm extratos baixados
    clientes_processados = verificar_extratos_existentes(mes_ano)
    
    # Filtra a lista de clientes para processar apenas os que não têm extratos
    # e separa por empresa (primeiro turim, depois tori)
    clients_turim = []
    clients_tori = []
    
    for client_id in clients_list:
        cliente_nome, empresa = clients[client_id]
        if cliente_nome not in clientes_processados:
            if empresa == 'turim':
                clients_turim.append(client_id)
            elif empresa == 'tori':
                clients_tori.append(client_id)
            else:
                # Se não for turim nem tori, adiciona ao final (tori)
                clients_tori.append(client_id)
    
    # Concatena turim primeiro, depois tori
    clients_para_processar = clients_turim + clients_tori
    
    if len(clients_para_processar) == 0:
        print(f"Todos os extratos do mês {mes_ano} já foram baixados!")
        return
    
    print(f"Encontrados {len(clientes_processados)} extratos existentes.")
    print(f"Serão processados {len(clients_para_processar)} clientes pendentes.")
    print(f"  - Turim: {len(clients_turim)} clientes")
    print(f"  - Tori: {len(clients_tori)} clientes")
    
    # Cria ou atualiza o arquivo de log
    if not os.path.exists(LOG_PATH):
        wb = Workbook()
        ws = wb.active
        ws.append(["Timestamp", "Cliente", "Status", "Caminho"])
    else:
        wb = load_workbook(LOG_PATH)
        ws = wb.active
    
    # Atualiza o log apenas com os clientes que serão processados
    for client_id in clients_para_processar:
        cliente_nome, empresa = clients[client_id]
        ws.append(["", cliente_nome, "", ""])
    
    wb.save(LOG_PATH)

    navegador = None
    max_retries = 3
    retry_count = 0
    
    while retry_count < max_retries:
        try:
            navegador = webdriver.Chrome(options=chrome_options)
            navegador.set_window_position(1920, 0)  # Ajustar para a posição x e y do seu segundo monitor
                    
            # Navegar até a página
            navegador.get(URL_SITE)
            navegador.maximize_window()
            break  # Se chegou aqui, a inicialização foi bem sucedida
        except Exception as e:
            retry_count += 1
            print(f"\nTentativa {retry_count} de {max_retries} falhou ao inicializar o navegador: {e}")
            if navegador:
                try:
                    navegador.quit()
                except:
                    pass
            if retry_count == max_retries:
                print("Falha ao inicializar o navegador após todas as tentativas.")
                return
            sleep(5)  # Espera 5 segundos antes de tentar novamente
    
    total_clients = len(clients_para_processar)
    if total_clients == 0:
        if navegador:
            navegador.quit()
        return

    try:
        # Autenticação
        try:
            WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@name='login']"))).send_keys(login) # usuário
            WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@name='password']"))).send_keys(senha) # senha

            token = input_token_btg() # token
            
            WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@name='softToken']"))).send_keys(token)
            sleep(0.5)
            # clicar no botão de Entrar
            WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "//button[@name='entrar']"))).click()    
            sleep(2)
        except Exception as e:
            print(f"\nErro na autenticação: {e}")
            return

        # Processar cada empresa separadamente
        empresas_a_processar = []
        if len(clients_turim) > 0:
            empresas_a_processar.append(('turim', clients_turim))
        if len(clients_tori) > 0:
            empresas_a_processar.append(('tori', clients_tori))
        
        for empresa, lista_clientes in empresas_a_processar:
            print(f"\n\n{'='*60}")
            print(f"Iniciando processamento dos clientes da {empresa.upper()}")
            print(f"Total de clientes: {len(lista_clientes)}")
            print(f"{'='*60}\n")
            
            # Selecionar o perfil correto
            if not selecionar_perfil(navegador, empresa):
                print(f"Não foi possível selecionar o perfil da {empresa.upper()}. Pulando...")
                continue
            
            # Após selecionar o perfil, navegar para a área de trabalho
            try:
                # clicar em Operacionalização
                WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//img[@src='/assets/img/ico-operation.svg']"))).click()
                sleep(0.5)
                # clicar em Atendimento ao cliente
                WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//li/a[@class='menu-second-route' and contains(text(), 'Atendimento ao cliente')]"))).click()
                sleep(0.5)
            except Exception as e:
                print(f"\nErro ao navegar para Atendimento ao cliente ({empresa.upper()}): {e}")
                continue

            # Loop principal de processamento dos clientes da empresa atual
            for i in range(len(lista_clientes)):
                cliente_nome, empresa_cliente = clients[lista_clientes[i]]
                print(f"\r{i+1}/{len(lista_clientes)} clientes {empresa.upper()} em processamento: {cliente_nome}", end="", flush=True)
                try:
                    if i == 0:
                        # buscar pelo cpf/conta/cnpj
                        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='text']"))).send_keys(lista_clientes[i])
                        sleep(0.5)
                        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//span[@class='ng-binding']"))).click()
                        sleep(0.5)    
                        # Clicar em Histórico
                        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//a[@class='nav-link text-uppercase ng-binding' and contains(text(),'Hist')]"))).click()
                        sleep(0.5)
                        # Clicar em Documentos
                        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//a[@class='ng-binding' and contains(text(),'Documentos')]"))).click()
                        sleep(0.5)
                        # Clicar em Extarto de Investimento
                        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//a[contains(text(),'Extrato de Investimento')]"))).click()
                        sleep(0.5)
                        # Realizar 3 rolagens para baixo
                        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='text']"))).click()
                
                        for _ in range(2):
                            # Scroll para baixo para garantir que todos os elementos estejam visíveis
                            navegador.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                            WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='text']"))).click()
                            sleep(1)  # Aguarde um momento para o carregamento completo
                            
                            # Verificar se foi liberado o do mês corrente
                            botoes = WebDriverWait(navegador, 10).until(EC.presence_of_all_elements_located((By.XPATH, f"//tr[@class='ng-scope' and td[contains(text(),'{mes_ano}')]]//button[@type='button' and @title='Gerar extrato']")))

                            # Selecionar o primeiro botão da lista e clicar
                            if botoes:
                                try:
                                    sleep(0.5)
                                    botoes[0].click()
                                    sleep(4)
                                    try:
                                        file_path = renomear_arquivo(DIRETORIO_DESTINO_PDFS, cliente_nome, mes_ano)
                                        log_download(cliente_nome, "ok", file_path)
                                    except Exception as e:
                                        print(f"\nErro ao salvar arquivo para {cliente_nome}: {e}")
                                        log_download(cliente_nome, "erro ao salvar", None)
                                except Exception as e:
                                    print(f"\nErro ao clicar no botão de download para {cliente_nome}: {e}")
                                    log_download(cliente_nome, "erro no download", None)
                            else:
                                log_download(cliente_nome, "not ok - botão não encontrado", None)

                    else:
                        # Extrair a URL atual para iterar por todos os clientes ajustando a url
                        current_url = navegador.current_url
                        # Mudar o cliente na URL
                        new_client = lista_clientes[i]
                        new_url = mudar_cliente(current_url, new_client)
                        navegador.get(new_url)

                        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='text']"))).click()
                        # Realizar 3 rolagens para baixo
                        for _ in range(4):
                            sleep(1)
                            # Scroll para baixo para garantir que todos os elementos estejam visíveis
                            navegador.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                        
                        sleep(2)
                        # Obter a lista de botões correspondentes ao BAIXAR TUDO
                        botoes = WebDriverWait(navegador, 10).until(EC.presence_of_all_elements_located((By.XPATH, f"//tr[@class='ng-scope' and td[contains(text(),'{mes_ano}')]]//button[@type='button' and @title='Gerar extrato']")))
                        # Selecionar o primeiro botão da lista e clicar
                        if botoes:
                            try:
                                botoes[0].click()
                                sleep(4)
                                try:
                                    file_path = renomear_arquivo(DIRETORIO_DESTINO_PDFS, cliente_nome, mes_ano)
                                    log_download(cliente_nome, "ok", file_path)
                                except Exception as e:
                                    print(f"\nErro ao salvar arquivo para {cliente_nome}: {e}")
                                    log_download(cliente_nome, "erro ao salvar", None)
                            except Exception as e:
                                print(f"\nErro ao clicar no botão de download para {cliente_nome}: {e}")
                                log_download(cliente_nome, "erro no download", None)
                        else:
                            log_download(cliente_nome, "not ok - botão não encontrado", None)
                except Exception as e:
                    print(f"\nErro ao processar cliente {cliente_nome}: {e}")
                    log_download(cliente_nome, "erro no processamento", None)
                    continue  # Continua para o próximo cliente mesmo se houver erro
            
            print(f"\n\nProcessamento da {empresa.upper()} concluído! {len(lista_clientes)} clientes processados.")
            
            # Se ainda houver outra empresa para processar, volta para a página multiprofile
            if empresa == 'turim' and len(clients_tori) > 0:
                print("\nVoltando para a página de seleção de perfil...")
                navegador.get(URL_EMPRESA)
                sleep(2)
        
        print(f"\n\n{'='*60}")
        print(f"Processamento total concluído! {total_clients} clientes processados.")
        print(f"{'='*60}")
    except Exception as e:
        print(f"\nErro durante o processamento: {e}")
    finally:
        if navegador:
            try:
                navegador.quit()
            except:
                print("Erro ao fechar o navegador")
                pass

def baixar_informes(login: str, senha: str, mes_ano: str) -> None:
    """
    Método que baixa os informes de imposto de renda do BTG Pactual.
    :param login: login do cliente
    :param senha: senha do cliente
    :return: True se o download for realizado com sucesso, False caso contrário.
    """
    chrome_options = webdriver.ChromeOptions()

    chrome_options.add_experimental_option("prefs", {
        "disable-popup-blocking": True,
        "download.default_directory": DIRETORIO_DESTINO_PDFS2,
        "savefile.default_directory": DIRETORIO_DESTINO_PDFS2,
    })

    # Carrega a lista de clientes
    clients = contas_clientes(path=r"F:\ExtratosBcosOff\Controle Data Analysis\controle_extratos_ONSHORE.xlsm")
    clients_list = list(clients.keys())
     
    # Filtra a lista de clientes para processar apenas os que não têm extratos
    # e separa por empresa (primeiro turim, depois tori)
    clients_turim = []
    clients_tori = []
    
    for client_id in clients_list:
        cliente_nome, empresa = clients[client_id]
        if empresa == 'turim':
            clients_turim.append(client_id)
        elif empresa == 'tori':
            clients_tori.append(client_id)
        else:
            # Se não for turim nem tori, adiciona ao final (tori)
            clients_tori.append(client_id)
    
    # Concatena turim primeiro, depois tori
    clients_para_processar = clients_turim + clients_tori
    
    if len(clients_para_processar) == 0:
        print(f"Todos os extratos do mês {mes_ano} já foram baixados!")
        return
    
    print(f"Encontrados {len(clients_list)} clientes existentes.")
    print(f"Serão processados {len(clients_para_processar)} clientes pendentes.")
    print(f"  - Turim: {len(clients_turim)} clientes")
    print(f"  - Tori: {len(clients_tori)} clientes")
    
    # Cria ou atualiza o arquivo de log
    if not os.path.exists(LOG_PATH):
        wb = Workbook()
        ws = wb.active
        ws.append(["Timestamp", "Cliente", "Status", "Caminho"])
    else:
        wb = load_workbook(LOG_PATH)
        ws = wb.active
    
    # Atualiza o log apenas com os clientes que serão processados
    for client_id in clients_para_processar:
        cliente_nome, empresa = clients[client_id]
        ws.append(["", cliente_nome, "", ""])
    
    wb.save(LOG_PATH)

    navegador = None
    max_retries = 3
    retry_count = 0
    
    while retry_count < max_retries:
        try:
            navegador = webdriver.Chrome(options=chrome_options)
            navegador.set_window_position(1920, 0)  # Ajustar para a posição x e y do seu segundo monitor
                    
            # Navegar até a página
            navegador.get(URL_SITE)
            navegador.maximize_window()
            break  # Se chegou aqui, a inicialização foi bem sucedida
        except Exception as e:
            retry_count += 1
            print(f"\nTentativa {retry_count} de {max_retries} falhou ao inicializar o navegador: {e}")
            if navegador:
                try:
                    navegador.quit()
                except:
                    pass
            if retry_count == max_retries:
                print("Falha ao inicializar o navegador após todas as tentativas.")
                return
            sleep(5)  # Espera 5 segundos antes de tentar novamente
    
    total_clients = len(clients_para_processar)
    if total_clients == 0:
        if navegador:
            navegador.quit()
        return

    try:
        # Autenticação
        try:
            WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@name='login']"))).send_keys(login) # usuário
            WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@name='password']"))).send_keys(senha) # senha

            token = input_token_btg() # token
            
            WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@name='softToken']"))).send_keys(token)
            sleep(0.5)
            # clicar no botão de Entrar
            WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, "//button[@name='entrar']"))).click()    
            sleep(2)
        except Exception as e:
            print(f"\nErro na autenticação: {e}")
            return

        # Processar cada empresa separadamente
        empresas_a_processar = []
        if len(clients_turim) > 0:
            empresas_a_processar.append(('turim', clients_turim))
        if len(clients_tori) > 0:
            empresas_a_processar.append(('tori', clients_tori))
        
        for empresa, lista_clientes in empresas_a_processar:
            print(f"\n\n{'='*60}")
            print(f"Iniciando processamento dos clientes da {empresa.upper()}")
            print(f"Total de clientes: {len(lista_clientes)}")
            print(f"{'='*60}\n")
            
            # Selecionar o perfil correto
            if not selecionar_perfil(navegador, empresa):
                print(f"Não foi possível selecionar o perfil da {empresa.upper()}. Pulando...")
                continue
            
            # Após selecionar o perfil, navegar para a área de trabalho
            try:
                # clicar em Operacionalização
                WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//img[@src='/assets/img/ico-operation.svg']"))).click()
                sleep(0.5)
                # clicar em Atendimento ao cliente
                WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//li/a[@class='menu-second-route' and contains(text(), 'Atendimento ao cliente')]"))).click()
                sleep(0.5)
            except Exception as e:
                print(f"\nErro ao navegar para Atendimento ao cliente ({empresa.upper()}): {e}")
                continue

            # Loop principal de processamento dos clientes da empresa atual
            for i in range(len(lista_clientes)):
                cliente_nome, empresa_cliente = clients[lista_clientes[i]]
                print(f"\r{i+1}/{len(lista_clientes)} clientes {empresa.upper()} em processamento: {cliente_nome}", end="", flush=True)
                try:
                    if i == 0:
                        # buscar pelo cpf/conta/cnpj
                        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='text']"))).send_keys(lista_clientes[i])
                        sleep(0.5)
                        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//span[@class='ng-binding']"))).click()
                        sleep(0.5)    
                        # Clicar em Histórico
                        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//a[@class='nav-link text-uppercase ng-binding' and contains(text(),'Hist')]"))).click()
                        sleep(0.5)
                        # Clicar em Documentos
                        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//a[@class='ng-binding' and contains(text(),'Documentos')]"))).click()
                        sleep(0.5)
                        # Clicar em Informe de Rendimento
                        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//a[contains(text(),'Imposto de Renda')]"))).click()
                        sleep(0.5)
                        for _ in range(2):
                            # Scroll para baixo para garantir que todos os elementos estejam visíveis            
                            close_overlays(navegador)  
                            main_content = navegador.find_element(By.CSS_SELECTOR, "btgmenu .main-content")
                            navegador.execute_script("arguments[0].scrollTop = arguments[0].scrollTop + arguments[1];", main_content, 700)
                            sleep(0.8)
                            botao = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//button[@id='downloadGroupsConfig'][1]")))
                            sleep(1)  # Aguarde um momento para o carregamento completo
                            if botao:
                                try:
                                    sleep(0.5)
                                    botao.click()
                                    sleep(3)
                                    try:
                                        file_path = renomear_informe(DIRETORIO_DESTINO_PDFS2, cliente_nome, mes_ano)
                                        log_download(cliente_nome, "ok", file_path)
                                    except Exception as e:
                                        print(f"\nErro ao salvar arquivo para {cliente_nome}: {e}")
                                        log_download(cliente_nome, "erro ao salvar", None)
                                except Exception as e:
                                    print(f"\nErro ao clicar no botão de download para {cliente_nome}: {e}")
                                    log_download(cliente_nome, "erro no download", None)
                            else:
                                log_download(cliente_nome, "not ok - botão não encontrado", None)

                    else:
                        # Extrair a URL atual para iterar por todos os clientes ajustando a url
                        current_url = navegador.current_url
                        # Mudar o cliente na URL
                        new_client = lista_clientes[i]
                        new_url = mudar_cliente(current_url, new_client)
                        navegador.get(new_url)
                        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "btgmenu .main-content")))
                        main_content = navegador.find_element(By.CSS_SELECTOR, "btgmenu .main-content")
                        navegador.execute_script("arguments[0].scrollTop = arguments[0].scrollTop + arguments[1];", main_content, 700)
                        sleep(0.8)
                        botao = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH,"//button[@id='downloadGroupsConfig'][1]")))
                        if botao:
                            try:
                                botao.click()
                                sleep(3)
                                try:
                                    file_path = renomear_informe(DIRETORIO_DESTINO_PDFS2, cliente_nome, mes_ano)
                                    log_download(cliente_nome, "ok", file_path)
                                except Exception as e:
                                    print(f"\nErro ao salvar arquivo para {cliente_nome}: {e}")
                                    log_download(cliente_nome, "erro ao salvar", None)
                            except Exception as e:
                                print(f"\nErro ao clicar no botão de download para {cliente_nome}: {e}")
                                log_download(cliente_nome, "erro no download", None)
                        else:
                            log_download(cliente_nome, "not ok - botão não encontrado", None)
                except Exception as e:
                    print(f"\nErro ao processar cliente {cliente_nome}: {e}")
                    log_download(cliente_nome, "erro no processamento", None)
                    continue  # Continua para o próximo cliente mesmo se houver erro
            
            print(f"\n\nProcessamento da {empresa.upper()} concluído! {len(lista_clientes)} clientes processados.")
            
            # Se ainda houver outra empresa para processar, volta para a página multiprofile
            if empresa == 'turim' and len(clients_tori) > 0:
                print("\nVoltando para a página de seleção de perfil...")
                navegador.get(URL_EMPRESA)
                sleep(2)
        
        print(f"\n\n{'='*60}")
        print(f"Processamento total concluído! {total_clients} clientes processados.")
        print(f"{'='*60}")
    except Exception as e:
        print(f"\nErro durante o processamento: {e}")
    finally:
        if navegador:
            try:
                navegador.quit()
            except:
                print("Erro ao fechar o navegador")
                pass