import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time

def coletar_dados_fiis(lista_fii):
    # Configuração do Selenium e navegador dentro da função de coleta de dados
    chrome_options = Options()
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_argument("--no-sandbox")

    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico, options=chrome_options)
    navegador.maximize_window()  # Maximiza a janela para garantir visibilidade

    lista_indicadores_fii = []

    for fii in lista_fii:
        url = f"https://statusinvest.com.br/fundos-imobiliarios/{fii}"

        try:
            navegador.get(url)
            
            # Aguarda os elementos estarem presentes na página
            WebDriverWait(navegador, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="fund-section"]/div/div/div[4]/div/div[1]/div/div/div/a/strong'))
            )
            WebDriverWait(navegador, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[1]/div/div[1]/strong'))
            )

            segmento = navegador.find_element(By.XPATH, '//*[@id="fund-section"]/div/div/div[4]/div/div[1]/div/div/div/a/strong').text
            tipo_anbima = navegador.find_element(By.XPATH, '//*[@id="fund-section"]/div/div/div[2]/div/div[5]/div/div/div/strong').text
            valor_atual = navegador.find_element(By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[1]/div/div[1]/strong').text
            dividend_yield = navegador.find_element(By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
            valor_dividendos_12m = navegador.find_element(By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[2]/div/span[2]').text

            try:
                pvp = navegador.find_element(By.XPATH, '//*[@id="main-2"]/div[5]/div/div[2]/div/div[1]/strong').text
            except:
                pvp = "N/A"

            # Coletando dados da tabela em todas as páginas
            dados_tabela = []
            for pagina in range(1, 6):  # Itera sobre as páginas de 1 a 5
                try:
                    tabela_xpath = '//*[@id="earning-section"]/div[7]/div/div[2]/table/tbody'
                    WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, tabela_xpath))
                    )
                    tabela = navegador.find_element(By.XPATH, tabela_xpath)
                    linhas = tabela.find_elements(By.TAG_NAME, "tr")

                    for linha in linhas:
                        colunas = linha.find_elements(By.TAG_NAME, "td")
                        if colunas:
                            dados_linha = [coluna.text for coluna in colunas]
                            dados_tabela.append(dados_linha)

                    # Navega para a próxima página se não for a última
                    if pagina < 5:
                        botao_proxima_xpath = '//*[@id="earning-section"]/div[7]/ul/li[@data-next="1"]/a'
                        try:
                            botao_proxima = WebDriverWait(navegador, 10).until(
                                EC.element_to_be_clickable((By.XPATH, botao_proxima_xpath))
                            )
                            botao_proxima.click()
                            time.sleep(2)  # Aguarda o carregamento da nova página
                        except Exception as e:
                            print(f"Erro ao clicar no botão 'próxima página' na página {pagina}: {e}")
                            break
                except Exception as e:
                    print(f"Erro ao processar página {pagina} para {fii}: {e}")
                    break

            valor_atual_float = float(valor_atual.replace("R$", "").replace(",", "."))
            valor_dividendos_12m_float = float(valor_dividendos_12m.replace("R$", "").replace(",", "."))

            rendimento_dividendos_percentual = (valor_dividendos_12m_float / valor_atual_float) * 100
            preco_ideal = valor_dividendos_12m_float / 0.08

            dicionario = {
                "fii": fii,
                "segmento": segmento,
                "tipo_anbima": tipo_anbima,
                "valor_atual": valor_atual,
                "dividend_yield": dividend_yield,
                "valor_dividendos_12m": valor_dividendos_12m,
                "rendimento_dividendos_percentual (%)": rendimento_dividendos_percentual,
                "preco_ideal": preco_ideal,
                "pvp": pvp,
                "dados_tabela": dados_tabela,
            }

            lista_indicadores_fii.append(dicionario)

        except Exception as e:
            print(f"Erro ao processar {fii}: {e}")

    navegador.quit()

    # Criando diretório no desktop
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    base_folder = os.path.join(desktop_path, "análise de FII")
    if not os.path.exists(base_folder):
        os.makedirs(base_folder)

    today = datetime.now().strftime("%Y-%m-%d")
    date_folder = os.path.join(base_folder, today)
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)

    now = datetime.now().strftime("%H-%M-%S")
    file_path = os.path.join(date_folder, f"fiis_{now}.xlsx")

    # Salvando os dados no Excel
    with pd.ExcelWriter(file_path) as writer:
        for fii in lista_indicadores_fii:
            df_info = pd.DataFrame([fii]).drop(columns=["dados_tabela"])
            df_info.to_excel(writer, sheet_name=fii["fii"] + "_info", index=False)

            if fii["dados_tabela"]:
                df_tabela = pd.DataFrame(fii["dados_tabela"], columns=["Período", "Rendimento", "Yield (%)", "Data de Pagamento"])
                df_tabela.to_excel(writer, sheet_name=fii["fii"] + "_tabela", index=False)

    return file_path

def iniciar_analise():
    fiis = entry_fiis.get().split(",")
    if fiis:
        label_status.config(text="Analisando FIIs, por favor aguarde...")
        root.update_idletasks()  # Atualiza a interface gráfica
        file_path = coletar_dados_fiis([fii.strip() for fii in fiis])
        label_status.config(text="Análise concluída!")
        messagebox.showinfo("Concluído", f"Análise concluída! Arquivo salvo em: {file_path}")
    else:
        messagebox.showwarning("Atenção", "Por favor, insira pelo menos um fundo imobiliário.")

def pressionar_enter(event):
    iniciar_analise()

root = tk.Tk()
root.title("Analisador de FIIs")
root.geometry("500x300")

style = ttk.Style(root)
style.theme_use('clam')

label = ttk.Label(root, text="Insira os códigos dos FIIs separados por vírgula:")
label.pack(pady=20)

entry_fiis = ttk.Entry(root, width=50)
entry_fiis.pack(pady=10)
entry_fiis.bind("<Return>", pressionar_enter)  # Associa a tecla Enter ao iniciar_analise

btn_analisar = ttk.Button(root, text="Iniciar Análise", command=iniciar_analise)
btn_analisar.pack(pady=20)

label_status = ttk.Label(root, text="")
label_status.pack(pady=10)

root.mainloop()
