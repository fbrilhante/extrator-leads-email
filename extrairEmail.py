import re
import imaplib 
import email
from email.header import decode_header 
import os
from dotenv import load_dotenv
import pandas as pd
import openpyxl

load_dotenv() 


EMAIL_USER = "franciscobraga99@gmail.com"
EMAIL_PASS = os.getenv("EMAIL_PASS")
IMAP_SERVER = "imap.gmail.com"


def buscar_email_lead():
    #conecta ao servidor do GMAIL usando SSL
    mail= imaplib.IMAP4_SSL(IMAP_SERVER)

    #login
    print("Conectando ao Gmail...")
    mail.login(EMAIL_USER, EMAIL_PASS)

    #seleciona a pasta 
    mail.select("inbox")

    #busca.
    status, messages = mail.search(None,'UNSEEN SUBJECT' , '"Lead - Cobertura Concept"')

    email_ids = messages[0].split()

    #se nao houver emails
    if not email_ids:
        print("Nenhum email de lead encontrado.")
        return None
    
    #pega o primeiro item da pilha (o email mais recente)
    ultimo_id = email_ids[-1]

    print(f"Processando email ID: {ultimo_id.decode()}...")

    #Fetch (baixar o conteúdo))
    #O padrão RFC822 recupera o email completo
    status, msg_data = mail.fetch(ultimo_id, "(RFC822)")


    #Decodifica assunto 
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            #transforma bytes em um objeto de email manipulavel
            msg = email.message_from_bytes(response_part[1])

            #decodifica o assunto (às vezes vem com caractere estranho)
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding if encoding else "utf-8")

    #Extrai  o corpo
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
        
            if content_type == "text/plain" and "attachment" not in content_disposition:
                body = part.get_payload(decode=True).decode()
                break
        
        else:
            body = msg.get_payload(decode=True).decode()
        
        return body 
    
    mail.close()
    mail.logout()



import pandas as pd


def extrair_dados_do_lead(corpo_do_email):
    """
    Extrai informações de um lead a partir do corpo do texto de um email.
    """
    if not corpo_do_email:
        return None

    lead = {}
    texto_linear = corpo_do_email.replace("\n", " ")

    # Regex para capturar infos
    match_nome = re.search(r"Name:\s*(.*?)\s*Telefone:", texto_linear)
    if match_nome:
        lead['Nome'] = match_nome.group(1).strip()
    
    match_data=re.search(r"Time:\s*(\d{2}/\d{2}/\d{4})",texto_linear)
    if match_data:
        lead['Data Entrada'] = match_data.group(1)

    match_hora = re.search(r"Time:\s*(\d{2}:\d{2})",texto_linear)
    if match_hora:
        lead['Hora Entrada'] = match_hora.group(1)

    match_email = re.search(r"Email:\s*([\w\.-]+@[\w\.-]+)", texto_linear)
    if match_email:
        lead['Email'] = match_email.group(1)

    match_public = re.search(r"utm_medium=([^&]+)", texto_linear)
    if match_public:
        lead['Publico'] = match_public.group(1).replace('+', ' ')

    match_content = re.search(r"utm_content=([^&]+)", texto_linear)
    if match_content:
        lead['Anuncio'] = match_content.group(1).replace('+', ' ')

    return lead


def salvar_em_excel(dados_do_lead):
    """
    Salva os dados extraídos de um lead em um arquivo Excel, adicionando ou atualizando.
    """
    if not dados_do_lead:
        print("Nenhum dado de lead para salvar.")
        return

    nome_arquivo = "leads_extraidos.xlsx"
    
    # Adiciona a nova coluna "Etapa Lead"
    dados_do_lead['Etapa Lead'] = 'Lead'
    
    # Converte o dicionário para um DataFrame do pandas
    df_novo_lead = pd.DataFrame([dados_do_lead])

    try:
        if os.path.exists(nome_arquivo):
            # Se o arquivo já existe, lê o conteúdo e adiciona o novo lead
            df_existente = pd.read_excel(nome_arquivo)
            df_final = pd.concat([df_existente, df_novo_lead], ignore_index=True)
        else:
            # Se o arquivo não existe, o DataFrame final é o novo lead
            df_final = df_novo_lead

        # Salva o DataFrame no arquivo Excel
        df_final.to_excel(nome_arquivo, index=False)
        
        nome_lead = dados_do_lead.get('Nome', 'Lead Desconhecido')
        print(f"Sucesso! Dados de '{nome_lead}' salvos em {nome_arquivo}")

    except Exception as e:
        print(f"Ocorreu um erro ao salvar o arquivo Excel: {e}")


def main():
    print("Iniciando processo de busca e extração de leads...")
    corpo_do_email = buscar_email_lead()

    if corpo_do_email:
        print("Email encontrado. Extraindo dados...")
        dados_lead = extrair_dados_do_lead(corpo_do_email)
        
        if dados_lead:
            print(f"Dados extraídos: {dados_lead}")
            salvar_em_excel(dados_lead)
        else:
            print("Não foi possível extrair dados do lead do corpo do email.")
    else:
        print("Nenhum corpo de email foi retornado para processamento.")

    print("Processo finalizado.")

if __name__ == "__main__":
    main()
