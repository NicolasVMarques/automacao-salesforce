import pyautogui
import time
import pandas as pd

tabela = pd.read_excel('Cópia de modelo para robo indicação.xlsx')

campo_primeiro_nome_x, campo_primeiro_nome_y = 541, 506
campo_sobrenome_x, campo_sobrenome_y = 1532, 563
campo_email_x, campo_email_y = 518, 677
campo_telefone_x, campo_telefone_y = 1179, 674
campo_razao_social_x, campo_razao_social_y = 500, 753
campo_cnpj_x, campo_cnpj_y = 1557, 741
botao_confirmar_x, botao_confirmar_y = 600, 700
email_padrao = 'empresario@gmail.com'

print("Iniciando automação em 5 segundos...")
print("Posicione a janela do formulário e não mexa no mouse!")
print("Para pausar a automação: Mova o mouse para o canto superior esquerdo.")
time.sleep(5)


for index, row in tabela.iterrows():
    
    contato = str(row['Contato'])
    telefone = str(row['Telefone'])
    cnpj = str(row['CNPJ'])
    email = str(row['EMAIL'])


    if pd.isna(email) or email.strip() == '':
        email_final = email_padrao
    else:
        email_final = email.strip() 

    print(f"\n--- Processando Linha {index + 1} de {len(tabela)} ---")
    print(f"Contato: {contato}")
    print(f"Telefone: {telefone}")
    print(f"CNPJ: {cnpj}")
    print(f"Email: {email_final}")

    pyautogui.click(campo_primeiro_nome_x, campo_primeiro_nome_y)
    time.sleep(0.3)
    pyautogui.write(contato, interval=0.05)

    pyautogui.click(campo_sobrenome_x, campo_sobrenome_y)
    time.sleep(0.3)
    pyautogui.write(contato, interval=0.05)
    
    pyautogui.click(campo_razao_social_x, campo_razao_social_y)
    time.sleep(0.3)
    pyautogui.write(contato, interval=0.05)

    pyautogui.click(campo_telefone_x, campo_telefone_y)
    time.sleep(0.3)
    pyautogui.write(telefone, interval=0.05)

    pyautogui.click(campo_cnpj_x, campo_cnpj_y)
    time.sleep(0.3)
    pyautogui.write(cnpj, interval=0.05)

    pyautogui.click(campo_email_x, campo_email_y)
    time.sleep(0.3)
    pyautogui.write(email_final, interval=0.05)

    print("Clicando em 'Salvar'...")
    pyautogui.click(botao_confirmar_x, botao_confirmar_y)

    print("Aguardando 3 segundos para o próximo registro...")
    time.sleep(3)
