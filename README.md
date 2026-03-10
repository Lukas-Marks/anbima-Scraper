# 📊 ANBIMA Asset Scraper

Script em Python para automatizar consultas no site da **ANBIMA** e extrair informações resumidas de ativos (ex.: debêntures) a partir de uma lista em Excel.

O script:

* lê códigos de ativos da **coluna A de um arquivo Excel**
* realiza busca automática no site da ANBIMA
* acessa a página **"Ver detalhes"**
* extrai os principais dados do ativo
* mostra um resumo no terminal
* opcionalmente salva os dados em **Excel utilizando pandas**

---

# 🚀 Funcionalidades

✔ Automação de consulta no site da ANBIMA
✔ Execução em **modo headless (navegador invisível)**
✔ Leitura de múltiplos ativos via Excel
✔ Extração automática de dados relevantes
✔ Exportação opcional para Excel
✔ Estrutura simples e extensível

---

# 📦 Tecnologias utilizadas

* Python 3
* Selenium
* Pandas
* OpenPyXL
* Microsoft Edge WebDriver

---

# 📂 Estrutura do projeto

```
anbima-scraper
│
├── script.py
├── isins.xlsx
├── README.md
└── ativos_anbima.xlsx (gerado automaticamente)
```

---

# 📥 Instalação

Clone o repositório:

```
git clone https://github.com/seu-usuario/anbima-scraper.git
```

Entre na pasta:

```
cd anbima-scraper
```

Instale as dependências:

```
pip install selenium pandas openpyxl
```

---

# 📊 Estrutura do arquivo de entrada

Crie um arquivo chamado:

```
isins.xlsx
```

Com os códigos na **coluna A**:

| A      |
| ------ |
| ABFR12 |
| PETR15 |
| XYZB23 |

---

# ▶️ Como executar

Execute o script:

```
python script.py
```

O programa irá:

1. abrir o site da ANBIMA
2. pesquisar cada ativo
3. acessar a página de detalhes
4. extrair os dados principais
5. perguntar se deseja salvar em Excel

---

# 📈 Exemplo de saída no terminal

```
Pesquisando: ABFR12

Resumo do ativo

Remuneração: IPCA + 8,1869%
Data emissão: 15/12/2021
Data vencimento: 15/12/2027
VNA: R$ 610,796371
ISIN: BRABFRDBS016
Próximo evento: 15/06/2026
```

---

# 📄 Arquivo gerado

Caso escolha salvar os resultados, será criado:

```
ativos_anbima.xlsx
```

Exemplo:

| Ativo  | Remuneração  | VNA    | Data vencimento | ISIN         |
| ------ | ------------ | ------ | --------------- | ------------ |
| ABFR12 | IPCA + 8,18% | 610,79 | 15/12/2027      | BRABFRDBS016 |

---

# ⚠️ Observações

O site da ANBIMA utiliza **React e carregamento dinâmico**, portanto o script utiliza Selenium para simular a navegação de um usuário.

Dependendo da conexão, pode ser necessário ajustar os tempos de espera (`sleep`).

---

# 🔮 Melhorias futuras

* extração de mais campos do ativo
* paralelização das consultas
* geração de dashboard
* substituição do Selenium por consultas diretas à API

---

# 👨‍💻 Autor

Projeto criado para estudo de **automação de coleta de dados financeiros** e integração com Python.

---
