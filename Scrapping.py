from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import time
import pandas as pd
import os

load_dotenv()

SMTP_URL = os.getenv('SMTP_URL')
SMTP_USER = os.getenv('SMTP_USER')
SMTP_PASSWORD = os.getenv('SMTP_PASSWORD')

# Pegar o diretório atual onde o script está sendo executado
diretorio_atual = os.path.dirname(os.path.abspath(__file__))

# Caminho do WebDriver
CHROME_DRIVER_PATH = "chromedriver.exe"

# Credenciais
LOGIN_URL = SMTP_URL
USERNAME = SMTP_USER
PASSWORD = SMTP_PASSWORD

# Lista de Códigos de Lojas
codigos_lojas = [
    "007162", "008333", "011996", "012268", "012396", "012674", "013003", "013391", "013557", 
    "013793", "017868", "019386", "019530", "019964", "020843", "020844", "020845", "020846", 
    "021045", "021056", "021057", "021058", "021059", "021060", "021061", "021062", "021207",
    "021274", "021372", "022631", "022744", "023252", "023263", "023264", "023269", "023666", 
    "023667", "911557"
]

# Criar um dicionário onde a chave é o código original e o valor é o código sem zeros à esquerda
dicionario_lojas = {codigo: str(int(codigo)) for codigo in codigos_lojas}

# Iniciar o navegador
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--window-size=1920,1080")
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 60)

# Abrir a página de login
driver.get(LOGIN_URL)
time.sleep(2)

# Localizar e preencher o campo de usuário
campo_usuario = driver.find_element(By.ID, "signInName")
campo_usuario.send_keys(USERNAME)

# Localizar e preencher o campo de senha
campo_senha = driver.find_element(By.NAME, "Password")
campo_senha.send_keys(PASSWORD)
campo_senha.send_keys(Keys.ENTER)

# Aguarde o carregamento do menu
wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[1]/div[1]/div/div/div/div/div/div/div/button')))

lojas_comentarios = [] # Dicionário para armazenar lojas e comentários

for i, valor_loja_desejada in enumerate(codigos_lojas): # Iterar sobre os códigos de loja
    try:
        valor_loja_atual = codigos_lojas[i] # Definir a loja atual como a loja que acabou de ser processada
        
        # Abrir o menu de filtros
        filtros_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[1]/div[1]/div/div/div/div/div/div/div/button')))
        filtros_button.click()
        wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="pfe_unitid-control"]')))
        
        # Clicar no botão "Redefinir" para limpar as seleções
        redefinir_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[1]/div[1]/div/div/div/div/form/div/div[2]/div/button[1]')))
        redefinir_button.click()
        
        # Abrir novamente o menu de filtros
        filtros_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[1]/div[1]/div/div/div/div/div/div/div/button')))
        filtros_button.click()
        wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="pfe_unitid-control"]')))

        # Expandir a lista de lojas
        expandir_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pfe_unitid-control"]')))
        expandir_button.click()
        time.sleep(2)  

        # Seleciona a loja desejada
        botao_loja = wait.until(EC.element_to_be_clickable((By.XPATH, f'//button[@data-label="{valor_loja_desejada}"]')))
        driver.execute_script("arguments[0].scrollIntoView();", botao_loja)
        botao_loja.click()

        # Seleciona o período desejado
        botao_tempo = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="timeperiod-control"]')))
        botao_tempo.click()
        selecao_tempo = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="340"]/div')))
        selecao_tempo.click()

        # Aplicar o filtro após selecionar a loja
        aplicar_filtro_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main-content-for-skip-link"]/div[1]/div[1]/div/div/div/div/form/div/div[2]/div/button[3]')))
        aplicar_filtro_button.click()
        time.sleep(3)  # Espera a aplicação do filtro

        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);") # Rolar a página para carregar todos os comentários
        time.sleep(3)  # Tempo para carregar os novos dados

        # Encontrar a div que contém os comentários, notas e datas
        comentarios_elements = driver.find_elements(By.XPATH, '//div[@data-test-name="comment-stream-comment"]')

        if comentarios_elements: # Verificar se comentários foram encontrados
            for comentario_element in comentarios_elements: # Extrair o texto, nota e data para cada comentário encontrado
                try:
                    comentario = comentario_element.find_element(By.XPATH, './/div[@data-test-name="expanded-comment"]/div/div[2]').text.strip() # Extrair o comentário
                    nota = comentario_element.find_element(By.XPATH, './/div[@data-test-name="comment-score"]/span').text.strip() # Extrair a nota
                    data = comentario_element.find_element(By.XPATH, './/span[@data-test-name="comment-date"]').text.strip() # Extrair a data
                    loja_formatada = dicionario_lojas[valor_loja_atual] # Armazenar a loja sem os zeros à esquerda
                    lojas_comentarios.append({"Loja": loja_formatada, "Comentario": comentario, "Nota": nota, "Data": data})# Adicionar os dados à lista ou processá-los conforme necessário
                    print(f"Loja: {loja_formatada}\nComentário: {comentario}\nNota: {nota}\nData: {data}\n")
                except Exception as e:
                    print(f"Erro ao processar o comentário: {e}")
        else:
            print(f"Não foram encontrados comentários para a loja {valor_loja_desejada}.")

        print(f"Dados carregados para a loja {valor_loja_desejada}")

    except Exception as e:
        print(f"Erro ao processar a loja {valor_loja_desejada}: {e}")

df_comentarios = pd.DataFrame(lojas_comentarios) # Criar um DataFrame com os dados extraídos

# Salvar os dados em um arquivo Excel
output_file = os.path.join(diretorio_atual, "comentarios_lojas.xlsx")
df_comentarios.to_excel(output_file, index=False)

print(f"Arquivo gerado: {output_file}")

input("Pressione Enter para fechar...") # Manter a janela aberta para ver os resultados

driver.quit() # Fecha o driver após apertar Enter