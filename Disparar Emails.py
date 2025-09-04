import os
import pandas as pd
import win32com.client as win32
 
# Caminho do arquivo Excel
excel_path = r"C:\\Users\\4006704\\Light\\BI ACR ACL CE - General\\Contratos 2025\\Base Contratos 2025 1.xlsx"
 
# Caminho base para os arquivos
base_path = r"C:\\Users\\4006704\\Light\\BI ACR ACL CE - General\\Contratos 2025"
 
# Caminho do arquivo fixo "FICHA CADASTRAL 2025.pdf"
ficha_cadastral_path = r"C:\\Users\\4006704\\Light\\BI ACR ACL CE - General\\Contratos 2025\\FICHA CADASTRAL 2025.pdf"
 
# Leitura da planilha
print("Lendo o arquivo Excel...")
try:
    df = pd.read_excel(excel_path)
    print(f"Arquivo lido com sucesso. Total de linhas: {len(df)}")
except Exception as e:
    print(f"Erro ao ler o arquivo Excel: {e}")
    exit()
# Normalizar os nomes das colunas (remover espaços extras)
df.columns = df.columns.str.strip()
# Verificar se todas as colunas necessárias estão presentes
required_columns = ['STATUS', 'LEILÃO', 'CONSÓRCIO / EMPRESA', 'CNPJ', 'EMPREENDIMENTO', 'CCEE', 'CCEAR', 'E-MAIL', 'E-MAIL 2', 'E-MAIL 3', 'E-MAIL 4', 'E-MAIL 5', 'E-MAIL 6', 'E-MAIL 7']
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    print(f"Colunas faltando no arquivo Excel: {missing_columns}")
    exit()
# Filtrar as linhas onde o STATUS é 'NÃO RECEBIDO'
print("Filtrando dados com STATUS 'NÃO RECEBIDO'...")
filtered_df = df[df['STATUS'].str.upper() == 'NÃO RECEBIDO']
print(f"Linhas com STATUS 'NÃO RECEBIDO': {len(filtered_df)}")
# Agrupar os dados por CONSÓRCIO / EMPRESA
print("Agrupando dados por CONSÓRCIO / EMPRESA...")
grouped = filtered_df.groupby('CONSÓRCIO / EMPRESA')
print(f"Total de grupos encontrados: {len(grouped)}")
# Função para formatar o CNPJ
def formatar_cnpj(cnpj):
    cnpj = str(cnpj).zfill(14)  # Garantir que tenha 14 dígitos
    return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:14]}"
# Função para enviar o e-mail pelo Outlook
def enviar_email_outlook(destinatario, assunto, corpo, anexos):
    try:
        print(f"Preparando e-mail para: {destinatario}")
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        # Configurar o e-mail
        mail.To = destinatario
        mail.CC = "GR_EMC_CONTRATOS@light.com.br"
        mail.Subject = assunto
        mail.HTMLBody = corpo
        # Anexar arquivos
        for anexo in anexos:
            if os.path.exists(anexo):
                mail.Attachments.Add(anexo)
                print(f"Anexo adicionado: {anexo}")
            else:
                print(f"Arquivo não encontrado: {anexo}")
        # Enviar o e-mail
        print(f"Enviando e-mail para {destinatario}...")
        mail.Send()
        print(f"E-mail enviado com sucesso para {destinatario}")
    except Exception as e:
        print(f"Erro ao enviar e-mail para {destinatario}: {e}")
# Loop para enviar e-mails para cada grupo
for consorcio, dados in grouped:
    print(f"Processando o grupo: {consorcio}")
    # Combinar todos os campos de e-mail em uma única lista
    email_columns = ['E-MAIL', 'E-MAIL 2', 'E-MAIL 3', 'E-MAIL 4', 'E-MAIL 5', 'E-MAIL 6', 'E-MAIL 7']
    destinatarios = dados[email_columns].apply(lambda x: x.dropna().unique(), axis=1).explode().dropna().unique()
    if not destinatarios.size:
        print(f"Nenhum destinatário encontrado para o grupo {consorcio}. Pulando...")
        continue
    # Juntar os e-mails em uma string separada por ponto e vírgula
    destinatario_principal = '; '.join(destinatarios)
    # Preparar o corpo do e-mail
    corpo_email = f"""

<div style="background-color: #fff3cd; color: #856404; padding: 10px; border: 1px solid #ffeeba; border-radius: 5px; font-family: Arial, sans-serif; margin-bottom: 20px;">
<h4 style="margin: 0; font-weight: bold;">ERRATA</h4>
<p style="margin: 0;">Favor considerar este e-mail para o envio das informações cadastrais do(s) contrato(s) 2025.</p>
</div>
<p>Prezados,</p>
<p>Boa tarde;</p>
<p>Tendo em vista o início de suprimento no ano de 2025 (referente ao {dados['LEILÃO'].iloc[0]}), peço o envio de seus respectivos dados cadastrais e de faturamento à Light.</p>
<p><b>Pedimos por gentileza que sigam este padrão para retorno das informações pertinentes ao cadastro. É importante não alterar a formatação, bem como os requisitos em si. Preservamos a padronização dos dados para manter a qualidade do processo.</b></p>
<hr>
<h3 style="color: #009a93;">Informações do Consórcio/Empresa: {consorcio}</h3>
<table border="1" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
<thead>
<tr style="background-color: #009a93; color: white; font-weight: bold;">
<th style="padding: 8px; text-align: center;">Leilão</th>
<th style="padding: 8px; text-align: center;">CONSÓRCIO / EMPRESA</th>
<th style="padding: 8px; text-align: center;">CNPJ</th>
<th style="padding: 8px; text-align: center;">Empreendimento</th>
<th style="padding: 8px; text-align: center;">CCEE</th>
<th style="padding: 8px; text-align: center;">CCEAR</th>
</tr>
</thead>
<tbody>
    """
    anexos = [ficha_cadastral_path]  # Lista de anexos, começando com o anexo fixo "FICHA CADASTRAL 2025.pdf"
    for _, row in dados.iterrows():
        cnpj_formatado = formatar_cnpj(row['CNPJ'])
        corpo_email += f"""
<tr>
<td style="padding: 8px; text-align: center;">{row['LEILÃO']}</td>
<td style="padding: 8px; text-align: center;">{row['CONSÓRCIO / EMPRESA']}</td>
<td style="padding: 8px; text-align: center;">{cnpj_formatado}</td>
<td style="padding: 8px; text-align: center;">{row['EMPREENDIMENTO']}</td>
<td style="padding: 8px; text-align: center;">{row['CCEE']}</td>
<td style="padding: 8px; text-align: center;">{row['CCEAR']}</td>
</tr>
        """
        # Adicionar o caminho do arquivo para anexo
        leilao = row['LEILÃO']
        ccear = str(row['CCEAR']).split('.')[0]  # Remove o ".0" se existir
        caminho_arquivo = os.path.join(base_path, leilao, f"CCEAR_{ccear}.pdf")
        anexos.append(caminho_arquivo)
    corpo_email += "</tbody></table><hr>"
    corpo_email += """
<p>Abaixo seguem os principais e-mails de acordo com a demanda.</p>
<ul>
<li><b>Resposta para este e-mail:</b> adriana.sobrinho@light.com.br; GR_EMC_CONTRATOS@light.com.br </li>
<li><b>Faturamento:</b> fatura.ccear@light.com.br</li>
<li><b>Assuntos relacionados ao faturamento:</b> heloisa.cavalcanti@light.com.br; renato.oliveira@light.com.br</li>
</ul>
<p>Estamos à disposição, obrigada.</p>
<p><i>Esta é uma mensagem automática.</i></p>
<p><b>Coordenação de BackOffice e ACR</b></p>
    """
    # Enviar o e-mail com anexos
    enviar_email_outlook(destinatario_principal, f" ERRATA Informações do CONSÓRCIO / EMPRESA {consorcio}", corpo_email, anexos)
print("Processo concluído!")