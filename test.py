import time
import pandas as pd
import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# =========================
# ESCOLHA DE ENTRADA
# =========================

print("\nModo de consulta")
print("1 - Digitar ativos manualmente")
print("2 - Ler arquivo Excel (isins.xlsx)")

modo = input("Escolha uma opção: ")

valores = []

if modo == "1":

    print("\nDigite os ativos (digite 'sair' para finalizar)\n")

    while True:

        ativo = input("Digite o código do ativo: ")

        if ativo.lower() == "sair":
            break

        valores.append(ativo)

elif modo == "2":

    arquivo_excel = r"C:\Users\lukas\OneDrive\DankiCode\Curso de Python\Projeto-Consultor-Anbima-Data\isins.xlsx"

    df = pd.read_excel(arquivo_excel)

    valores = df.iloc[:,0].dropna().astype(str).tolist()

else:

    print("Opção inválida")
    exit()


print("\nAtivos que serão consultados:")
print(valores)


# =========================
# CONFIG NAVEGADOR
# =========================

options = Options()

options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")

driver = webdriver.Edge(options=options)

wait = WebDriverWait(driver, 30)

url = "https://data.anbima.com.br/"

driver.get(url)

time.sleep(5)

lista_resultados = []


# =========================
# LOOP DE CONSULTA
# =========================

for ativo in valores:

    print("\nPesquisando:", ativo)

    try:

        input_busca = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//input[@type='search']")
            )
        )

        input_busca.click()
        input_busca.clear()
        input_busca.send_keys(ativo)
        input_busca.send_keys(Keys.ENTER)

        time.sleep(5)

        botao_detalhes = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "(//span[text()='Ver detalhes'])[1]")
            )
        )

        driver.execute_script(
            "arguments[0].click();", botao_detalhes
        )

        time.sleep(5)

        # =========================
        # EXTRAÇÃO DOS DADOS
        # =========================

        dados = {"Ativo": ativo}

        campos = {
            "Remuneração": "output__container--remuneracao",
            "Data início rentabilidade": "output__container--dataInicioRentabilidade",
            "VNA": "output__container--vna",
            "VNE": "output__container--vne",
            "Volume emissão": "output__container--volumeSerieEmissao",
            "Quantidade emissão": "output__container--quantidadeSerieEmissao",
            "Data emissão": "output__container--dataEmissao",
            "Data vencimento": "output__container--dataVencimento",
            "Prazo remanescente": "output__container--prazoRemanescente",
            "ISIN": "output__container--isin",
            "Próximo evento": "output__container--dataProximoEventoAgenda"
        }

        for nome, campo_id in campos.items():

            try:

                valor = driver.find_element(By.ID, campo_id).text.split("\n")[-1]

                dados[nome] = valor

            except:

                dados[nome] = "N/A"

        print("\nResumo do ativo")

        for k,v in dados.items():
            print(f"{k}: {v}")

        lista_resultados.append(dados)

    except Exception as erro:

        print("Erro ao processar:", ativo)
        print(erro)

    driver.get(url)

    time.sleep(5)


# =========================
# SALVAR RESULTADOS
# =========================

salvar = input("\nDeseja salvar os resultados em Excel? (s/n): ")

if salvar.lower() == "s":

    nome_arquivo = r"C:\Users\lukas\OneDrive\DankiCode\Curso de Python\Projeto-Consultor-Anbima-Data\resumo_ativos.xlsx"

    # apagar arquivo antigo se existir
    if os.path.exists(nome_arquivo):
        os.remove(nome_arquivo)
        print("Arquivo antigo removido.")

    df_final = pd.DataFrame(lista_resultados)

    df_final.to_excel(nome_arquivo, index=False)

    print(f"\nArquivo salvo: {nome_arquivo}")

else:

    print("\nResultados não foram salvos.")


input("\nPressione ENTER para finalizar")

driver.quit()