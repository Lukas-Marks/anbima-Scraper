import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# =========================
# CONFIG
# =========================

arquivo_excel = r"C:\Users\lukas\OneDrive\DankiCode\Curso de Python\Projeto-Consultor-Anbima-Data\isins.xlsx"
url = "https://data.anbima.com.br/"

# =========================
# LER EXCEL
# =========================

df = pd.read_excel(arquivo_excel)
valores = df.iloc[:,0].dropna().astype(str).tolist()

print("Ativos encontrados no Excel:")
print(valores)

# =========================
# NAVEGADOR
# =========================

options = Options()
options.add_argument("--start-maximized")

driver = webdriver.Edge(options=options)
wait = WebDriverWait(driver, 30)

driver.get(url)

time.sleep(5)

lista_resultados = []

# =========================
# LOOP
# =========================

for valor in valores:

    print("\nPesquisando:", valor)

    try:

        input_busca = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//input[@type='search']")
            )
        )

        input_busca.click()
        input_busca.clear()
        input_busca.send_keys(valor)
        input_busca.send_keys(Keys.ENTER)

        time.sleep(5)

        botao_detalhes = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "(//span[text()='Ver detalhes'])[1]")
            )
        )

        driver.execute_script("arguments[0].click();", botao_detalhes)

        time.sleep(5)

        # =========================
        # EXTRAIR DADOS
        # =========================

        dados = {"Ativo": valor}

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
                valor_campo = driver.find_element(By.ID, campo_id).text.split("\n")[-1]
                dados[nome] = valor_campo
            except:
                dados[nome] = "N/A"

        # mostrar resumo
        print("\nResumo do ativo")

        for k,v in dados.items():
            print(f"{k}: {v}")

        lista_resultados.append(dados)

    except Exception as e:

        print("Erro no ativo:", valor)
        print(e)

    # voltar para home
    driver.get(url)
    time.sleep(5)


# =========================
# SALVAR RESULTADOS
# =========================

salvar = input("\nDeseja salvar os resultados em Excel? (s/n): ")

if salvar.lower() == "s":

    df_final = pd.DataFrame(lista_resultados)

    df_final.to_excel(r"C:\Users\lukas\OneDrive\DankiCode\Curso de Python\Projeto-Consultor-Anbima-Data\ativos_anbima.xlsx", index=False)

    print("\nArquivo salvo como: ativos_anbima.xlsx")

else:

    print("\nResultados não foram salvos.")


input("\nPressione ENTER para fechar")

driver.quit()