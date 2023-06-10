from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


number_of_pages = 3

# inicialize um driver do selenium para controlar um navegador
driver = webdriver.Chrome()
driver.implicitly_wait(30)

# encontre o campo de email e preencha com o seu email
wait = WebDriverWait(driver, 30)

page = 1

df = pd.DataFrame(columns=['Coluna 1', 'Href'])

while page <= number_of_pages:
    # pegue o HTML da página atual
    driver.get(f'https://dashboard.onesignal.com/apps?page={page}')

    if page ==1:

        search_input = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.ID, 'field_email')
            )
        )

        email_field = driver.find_element(By.ID,'field_email')
        email_field.send_keys('desenvolvimento@scaffoldeducation.com.br')

        # encontre o campo de senha e preencha com a sua senha
        password_field = driver.find_element(By.ID,'field_password')
        password_field.send_keys('E2?u7/V9_L,9,NS')

        # encontre o botão de login e clique nele
        login_button = driver.find_element(By.XPATH, '//*[@id="root"]/div[2]/div/div/form/fieldset/div/button')
        login_button.click()
    
    search_input = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located(
        (By.CLASS_NAME, 'sc-diYIgC hgzULd')
        )
    )      

    html = driver.page_source

    # use o BeautifulSoup para analisar o HTML
    soup = BeautifulSoup(html, 'html.parser')

    # encontre a tabela na página
    table = soup.find('table', {'class': 'sc-fBwhgt iBKoez'})

    rows = table.find_all('tr')

    # percorra as linhas da tabela e imprima a primeira coluna de cada linha
    for linha in rows:
        first_column_tag = linha.find('td')
        if first_column_tag is not None:
            first_column = first_column_tag.text.strip()
            tag_a = first_column_tag.find('a')
            if tag_a is not None:
                href = tag_a['href']
            else:
                href = None
            df = df.append({'Coluna 1': first_column, 'Href': href}, ignore_index=True)

    page += 1

# escreve o DataFrame em uma planilha Excel
writer = pd.ExcelWriter('dados.xlsx')
df.to_excel(writer, index=False)
writer.save()