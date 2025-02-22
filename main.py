import re
import PyPDF2
import tkinter as tk
from tkinter import filedialog
import calendar
from selenium import webdriver
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
from fpdf import FPDF  # Biblioteca para gerar PDFs
import platform
import sys


def extrair_texto_pdf(pdf_path):
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        texto = ""
        for page in reader.pages:
            texto += page.extract_text()
    return texto

def extrair_aih_codes(texto):
    return re.findall(r"\b\d{13}\b", texto)

def extrair_competencias(texto):
    return re.findall(r"Compet[eê]ncia:\s*(\d{1,2}/\d{4})", texto)

def get_date_ranges(competencia):
    month_str, year_str = competencia.split("/")
    month = int(month_str)
    year = int(year_str)
    ranges = []
    
    for i in range(2):  # 2 intervalos de 2 meses cada
        start_month = month - (2*i + 1)
        end_month = month - (2*i)
        start_year = year
        end_year = year
        
        while start_month <= 0:
            start_month += 12
            start_year -= 1
            
        while end_month <= 0:
            end_month += 12
            end_year -= 1
        
        start_date = f"01/{start_month:02d}/{start_year}"
        end_day = calendar.monthrange(end_year, end_month)[1]
        end_date = f"{end_day:02d}/{end_month:02d}/{end_year}"
        
        ranges.append((start_date, end_date))
    
    # Reverte a lista para que o período mais antigo venha primeiro.
    ranges.reverse()
    return ranges

def iniciar_navegador():
    # Define o caminho base: se estiver empacotado, usa sys._MEIPASS
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = "."
    
    arch = platform.architecture()[0]
    if arch == "64bit":
        driver_path = f"{base_path}/geckodriver/64x/geckodriver.exe"
    else:
        driver_path = f"{base_path}/geckodriver/32x/geckodriver.exe"
    
    service = FirefoxService(executable_path=driver_path)
    driver = webdriver.Firefox(service=service)
    return driver

def get_credentials():
    """Exibe uma interface gráfica para capturar login e senha do usuário."""
    cred_window = tk.Tk()
    cred_window.title("Credenciais de Acesso ao CARE")

    tk.Label(cred_window, text="Login:").grid(row=0, column=0, padx=5, pady=5)
    login_entry = tk.Entry(cred_window)
    login_entry.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(cred_window, text="Senha:").grid(row=1, column=0, padx=5, pady=5)
    senha_entry = tk.Entry(cred_window, show="*")
    senha_entry.grid(row=1, column=1, padx=5, pady=5)

    creds = {}

    def submit():
        creds["login"] = login_entry.get()
        creds["senha"] = senha_entry.get()
        cred_window.quit()  # Finaliza o loop do Tkinter
        cred_window.destroy()

    tk.Button(cred_window, text="OK", command=submit).grid(row=2, column=0, columnspan=2, pady=10)
    cred_window.update_idletasks()
    cred_window.mainloop()
    
    return creds.get("login", ""), creds.get("senha", "")

def gerar_pdf(report, comp, save_path):
    """Gera um PDF com o relatório de auditoria.
       O título inclui a competência e logo abaixo é exibido o total de AIH's autorizadas e não autorizadas."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Título
    pdf.set_font("Arial", 'B', 16)
    title = f"Auditoria AIH'S previamente autorizadas - competência {comp}"
    pdf.cell(0, 10, title, ln=True, align="C")
    pdf.ln(5)
    
    # Contagem dos status
    authorized = 0
    not_authorized = 0
    for linha in report:
        if "com status autorizada" in linha.lower():
            authorized += 1
        elif "sem status de autorização" in linha.lower():
            not_authorized += 1
        elif "não encontrada" in linha.lower():
            not_authorized += 1  # Considera as não encontradas como não autorizadas

    summary = f"Total autorizadas: {authorized}    Total não autorizadas: {not_authorized}"
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, summary, ln=True, align="C")
    pdf.ln(10)
    
    # Relatório detalhado
    pdf.set_font("Arial", size=12)
    for linha in report:
        pdf.multi_cell(0, 10, linha)
    
    pdf.output(save_path)
    print(f"Relatório salvo como: {save_path}")

def realizar_pesquisa(driver, date_ranges, aih_codes, login, senha):
    
    wait = WebDriverWait(driver, 10)
    report = []

    # === Parte 1: Login Duplo e Popup ===
    driver.get("https://www.saude.pr.gov.br/Pagina/Sistema-Estadual-de-Regulacao")
    
    print("25 segundos para a inclusão de senha de internet.")
    time.sleep(25)

    elemento_care = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//span[@class='collapsible-item-title-link-text' and contains(text(), 'CARE Paraná')]")
    ))
    elemento_care.click()
    print("Cliquei em CARE Paraná")
    time.sleep(2)

    elemento_img = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//img[@src='/sites/default/arquivos_restritos/files/imagem/2020-06/aih_care.png']")
    ))
    driver.execute_script("arguments[0].scrollIntoView(true);", elemento_img)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", elemento_img)

    original_tab = driver.current_window_handle
    wait.until(lambda d: len(d.window_handles) > 1)
    new_tab = [handle for handle in driver.window_handles if handle != original_tab][0]
    driver.switch_to.window(new_tab)
    print("Nova aba aberta. Título:", driver.title)
    time.sleep(2)

    # **Login na central (login duplo)**
    for attempt in range(2):
        btn_central = wait.until(EC.element_to_be_clickable((By.ID, "btnCentral")))
        btn_central.click()

        cpf_field = wait.until(EC.element_to_be_clickable((By.ID, "attribute_central")))
        cpf_field.clear()
        cpf_field.send_keys(login)
        print("Login inserido.")

        password_field = wait.until(EC.element_to_be_clickable((By.ID, "password")))
        password_field.clear()
        password_field.send_keys(senha)
        print("Senha inserida.")

        btn_entrar = wait.until(EC.element_to_be_clickable((By.ID, "btn-central-acessar")))
        btn_entrar.click()
        time.sleep(2)
    
    if len(driver.window_handles) > 1:
        popup_handle = driver.window_handles[-1]
        driver.switch_to.window(popup_handle)
    else:
        print("Popup detectado na mesma janela.")
    time.sleep(2)
    
    try:
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "frameId")))
        print("Trocado para o frame com id 'frameId'.")
    except Exception as e:
        print("Erro ao trocar para o frame:", e)
    
    try:
        driver.execute_script("oCMenu.onclck('_14283');")
        print("Menu 'Avaliar Laudo de AIH' foi selecionado automaticamente.")
        time.sleep(3)  # Aguarda a página carregar após a seleção
    except Exception as e:
        print("Erro ao tentar selecionar o menu automaticamente:", e)
        driver.quit()
        exit()
        
    # === Parte 3: Processar cada AIH com os intervalos de data ===
    for aih_code in aih_codes:
        aih_encontrada = False
        # Insere o número da AIH uma vez (permanece fixo)
        try:
            numero_aih_input = wait.until(EC.element_to_be_clickable((By.ID, "numeroAih")))
            driver.execute_script("arguments[0].scrollIntoView(true);", numero_aih_input)
            time.sleep(1)
            numero_aih_input.click()
            numero_aih_input.clear()
            numero_aih_input.send_keys(aih_code)
            print(f"Número AIH {aih_code} inserido.")
        except Exception as e:
            print("Erro ao inserir número AIH:", e)
            continue
        
        for start_date, end_date in date_ranges:
            print(f"\nPesquisando AIH {aih_code} para o período {start_date} a {end_date}")
            try:
                data_inicio_input = wait.until(EC.element_to_be_clickable((By.ID, "dataInicio")))
                data_inicio_input.clear()
                data_inicio_input.send_keys(start_date)
                print(f"Data de Início {start_date} inserida.")
            except Exception as e:
                print("Erro ao inserir Data de Início:", e)
                continue
            
            try:
                data_fim_input = wait.until(EC.element_to_be_clickable((By.ID, "dataFim")))
                data_fim_input.clear()
                data_fim_input.send_keys(end_date)
                print(f"Data de Fim {end_date} inserida.")
            except Exception as e:
                print("Erro ao inserir Data de Fim:", e)
                continue
            
            try:
                btn_pesquisar = wait.until(EC.element_to_be_clickable((By.ID, "btnPesquisar")))
                driver.execute_script("arguments[0].scrollIntoView(true);", btn_pesquisar)
                time.sleep(1)
                btn_pesquisar.click()
                print("Cliquei no botão Pesquisar.")
                time.sleep(2)
            except Exception as e:
                print("Erro ao clicar no botão Pesquisar:", e)
                continue
            
            # Verifica se "Nenhum registro encontrado" aparece
            try:
                driver.find_element(By.XPATH, "//td[contains(@class, 'msg_aviso') and contains(text(), 'Nenhum registro encontrado')]")
                print(f"Nenhum registro encontrado para AIH {aih_code} no período {start_date} a {end_date}.")
                continue  # Tenta o próximo período
            except:
                print(f"Registro encontrado para AIH {aih_code} no período {start_date} a {end_date}.")
                try:
                    status_element = driver.find_element(By.XPATH, "//div[contains(@id, 'aih.status.descricao')]")
                    status_text = status_element.text.strip()
                    if status_text.lower() == "autorizado":
                        report.append(f"AIH nº {aih_code} com status autorizada.")
                        print(f"AIH {aih_code} está autorizada.")
                    else:
                        report.append(f"AIH nº {aih_code} sem status de autorização (Status: {status_text}).")
                        print(f"AIH {aih_code} encontrada, mas status: {status_text}.")
                    aih_encontrada = True
                    break
                except Exception as e:
                    print("Erro ao verificar o status da AIH:", e)
                    report.append(f"AIH nº {aih_code} - erro ao verificar status.")
                    aih_encontrada = True
                    break
        
        if not aih_encontrada:
            report.append(f"AIH nº {aih_code} não encontrada em nenhum período.")
    
    print("\nRELATÓRIO FINAL:")
    for linha in report:
        print(linha)
    
    return report

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    pdf_path = filedialog.askopenfilename(
        title="Selecione o arquivo PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )

    if pdf_path:
        login, senha = get_credentials()  # Captura login e senha do usuário
        texto = extrair_texto_pdf(pdf_path)
        aih_codes = extrair_aih_codes(texto)
        competencias = extrair_competencias(texto)
        competencias = list(set(competencias))
        competencias.sort(key=lambda x: (int(x.split("/")[1]), int(x.split("/")[0])))

        print("Códigos de AIH encontrados:", aih_codes)
        print("Competências encontradas:", competencias)

        if aih_codes and competencias:
            comp = competencias[0]  # Usa a primeira competência
            date_ranges = get_date_ranges(comp)
            print(f"Meses da pesquisa da competência: {comp}:")
            for start_date, end_date in date_ranges:
                print(f"  De {start_date} a {end_date}")
            
            driver = iniciar_navegador()
            print("Iniciando a navegação para auditoria...")
            report = realizar_pesquisa(driver, date_ranges, aih_codes, login, senha)
            time.sleep(2)
            driver.quit()
            
            # Abre uma janela para o usuário escolher onde salvar o PDF
            save_path = filedialog.asksaveasfilename(
                title="Salvar Relatório PDF",
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                initialfile=f"Auditoria_AIHS_{comp.replace('/', '-')}.pdf"
            )
            if save_path:
                gerar_pdf(report, comp, save_path)
            else:
                print("Nenhum caminho de salvamento selecionado.")
        else:
            print("Não foram encontrados códigos de AIH ou competências.")
    else:
        print("Nenhum arquivo foi selecionado.")
