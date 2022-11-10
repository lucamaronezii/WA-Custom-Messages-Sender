from IPython.display import display
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
import urllib
from urllib import parse
from selenium.webdriver.common.by import By

# leitura do Excel que servirá de base para disparar o envio de mensagens
contatos_df = pd.read_excel("QualquerPlanilha.xlsx")

# definição do navegador utilizado. Logo depois, entrando no WhatsApp através da URL
navegador = webdriver.Chrome()
navegador.get("https://web.whatsapp.com/")

# enquanto o usuário não escanear o QRCode, o código fica estagnado
while len(navegador.find_elements(By.XPATH, '//div[@id="side"]')) < 1:
    sleep(1)

# seleção das colunas para enviar as mensagens
for i, mensagem in enumerate(contatos_df['Mensagem']):
    pessoa = contatos_df.loc[i, "Nome"]  # importando nome personalizado da pessoa na mensagem
    numero = contatos_df.loc[i, "Telefone"]  # importando respectivo número de telefone
    if numero == "-" or numero == "nan" or numero == "55-" or numero == "550000000":
        pass
    else:
        # importando mensagens personalizadas nas respectivas colunas
        mensagem2 = contatos_df.loc[i, "Mensagem2"]

        # formatando mensagem para que fique em conformidade com a URL
        texto = urllib.parse.quote(f'Olá, {pessoa}\n\n{mensagem}\n\n{mensagem2}')
        link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
        navegador.get(link)
        sleep(10)

        # se o número não existir, um pop-up aparecerá na tela. Nesses casos, o código quebrava.
        # Agora, o operador WHILE remove o pop-up, caso apareça, e continua para o próximo contato
        # da lista
        while len(navegador.find_elements(By.XPATH, '//div[@data-testid="popup-contents"]')) == 1:
            navegador.find_element_by_xpath('//div[@role="button"]').send_keys(Keys.ENTER)
            break
        else:
            # se o número existir, a mensagem é disparada com os respectivos delays definidos nas
            # funções SLEEP.
            sleep(10)  # 12

            # aguarda o aparecimento do elemento único definido pelo XPATH
            while len(navegador.find_elements(By.XPATH, '//div[@id="side"]')) < 1:
                sleep(1)

            # clica no botão de envio
            navegador.find_element_by_xpath(
                '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div').send_keys(Keys.ENTER)

            sleep(8)    # quanto mais tempo, melhor. Menos chance terá do WhatsApp bloquear o número
