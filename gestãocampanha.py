from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time
import pandas as pd


def salvar_em_excel(dados, nome_arquivo, atendentes):
    print("Transformando os dados em DataFrame...")
    df = pd.DataFrame(dados, columns=['Número de telefone', 'Data e Hora do último acesso', 'Acessos ao link', 'Status interações'])
    
    print("Calculando distribuição de linhas por atendente...")
    num_linhas_por_atendente = len(df) // len(atendentes)
    resto = len(df) % len(atendentes)
    print(f"Cada atendente receberá aproximadamente {num_linhas_por_atendente} linhas.")
    if resto:
        print(f"Serão distribuídas {resto} linhas adicionais para os primeiros atendentes.")

    atendente_coluna = []
    for index, atendente in enumerate(atendentes):
        if index < resto:
            atendente_coluna += [atendente] * (num_linhas_por_atendente + 1)
        else:
            atendente_coluna += [atendente] * num_linhas_por_atendente
    
    if resto:
        atendente_coluna += atendentes[:resto]

    df['Atendente'] = atendente_coluna[:len(df)]  # Corta a lista para o tamanho do df
    print("Coluna de atendentes atribuída ao DataFrame.")

    print("Salvando os dados em arquivo Excel...")
    df.to_excel(nome_arquivo, index=False)
    print(f"Dados salvos em {nome_arquivo}")

def extrair_dados_da_pagina(browser):
    html = browser.page_source
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table', {'class': 'MuiTable-root'})
    rows = table.find_all('tr')[1:]  # Ignora o cabeçalho
    data = [[cell.get_text(strip=True) for cell in row.find_all('td')[:4]] for row in rows]
    return data

# Configuração do WebDriver para o Microsoft Edge
edge_options = webdriver.EdgeOptions()
edge_options.add_argument('--enable-chrome-browser-cloud-management')
edge_options.add_argument('--start-maximized')
browser = webdriver.Edge(options=edge_options)

# Abrindo a página desejada
url = '*************'
browser.get(url)

# Logando na página
elemento_login = WebDriverWait(browser, 30).until(
    EC.element_to_be_clickable((By.ID, "email"))
)
elemento_login.send_keys("********@gmail.com")

elemento_senha = WebDriverWait(browser, 30).until(
    EC.element_to_be_clickable((By.ID, "password"))
)
elemento_senha.send_keys("*******")

# Clicando no botão de login
botao_login = WebDriverWait(browser, 30).until(
    EC.element_to_be_clickable((By.XPATH, '//button[@type="submit"]'))
)
botao_login.click()

elemento_menu_lateral = WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sidebar-subitem-mon-de-env"]')))
elemento_menu_lateral.click()
x = input("campanha escolhida?")
# Aguardar a tabela carregar
WebDriverWait(browser, 30).until(
    EC.presence_of_element_located((By.CLASS_NAME, "MuiTable-root"))
)

# Lista para armazenar todos os dados
all_data = []

# Loop para percorrer todas as páginas
while True:
    # Extrair dados da página atual
    current_page_data = extrair_dados_da_pagina(browser)
    all_data.extend(current_page_data)
    print(f"Dados da página atual extraídos: {len(current_page_data)} registros encontrados.")
    
    # Tenta clicar no botão 'Próxima página'
    try:
        next_button = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="btn-prox-pagina"]'))
        )
        # Verifica se o botão está desativado (última página)
        if "alguma_classe_que_indica_desabilitado" in next_button.get_attribute('class'):
            print("Última página alcançada.")
            break
        next_button.click()
        print("Navegando para a próxima página...")
        # Espera a página carregar após clicar
        time.sleep(5)
    except Exception as e:
        print("Erro ao tentar navegar para a próxima página ou última página alcançada.", str(e))
        break

# Imprime todos os dados coletados
print(f"Total de dados coletados: {len(all_data)} registros.")
# Lista de atendentes
atendentes = [
    "A**** C*****", "P******* V*****", "K****** K****", "C**** L****",
    "S*** L*****", "H***** S******", "R***** S****", "M***** L****",
    "M***** B*****", "S***** M****", "G****** O*******", "R***** M***",
    "B**** O******", "V***** A******", "L**** G****", "T****** R****",
    "V***** C****", "E**** V****", "C*** N******", "T******* C****"
]


# Certifique-se de chamar esta função após a coleta de todos os dados
salvar_em_excel(all_data, 'dados_campanha.xlsx', atendentes)


# Fecha o navegador
browser.quit()