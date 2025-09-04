import os
import pandas as pd
import fitz  # PyMuPDF
import re
 
# Caminho da planilha e pasta de destino
planilha_path = r"C:\\Users\\4006704\\Light\\BI ACR ACL CE - General\\Contratos 2025\\Outra Base Teste.xlsx"
diretorio_base = r"C:\\Users\\4006704\\Light\\BI ACR ACL CE - General\\Contratos 2025\\34º LEILÃO DE ENERGIA NOVA A-4"
 
# Carregar a planilha
df = pd.read_excel(planilha_path)
 
# Verificar se a coluna "LEILÃO" existe
if "LEILÃO" not in df.columns:
    print("A coluna 'LEILÃO' não foi encontrada na planilha.")
    exit()
 
# Filtrar as linhas que possuem o leilão "30º LEILÃO DE ENERGIA NOVA A-6"
df_filtrado = df[df["LEILÃO"] == "34º LEILÃO DE ENERGIA NOVA A-4"].copy()
 
# Função para extrair dados do PDF
def extrair_dados_pdf(caminho_pdf):
    try:
        # Abrir o arquivo PDF
        doc = fitz.open(caminho_pdf)
        texto = ""
 
        # Extrair texto de todas as páginas
        for pagina in doc:
            texto += pagina.get_text()
 
        # Usar regex para capturar os três primeiros contatos, telefones e e-mails
        contatos = re.findall(r"A\/C: ([^\n]+)", texto)
        telefones = re.findall(r"Tel\.: ([^\n]+)", texto)
        emails = re.findall(r"E-mail: ([^\n]+)", texto)
 
        # Preencher os campos com até 3 valores ou "" caso não exista ou seja indesejado
        dados = {
            "CONTATO": contatos[0] if len(contatos) > 0 and contatos[0] not in ["CARLOS DURVAL DE MORAES", "CARLOS DURVAL DE MORAIS", "IZABELLA REBOREDO VERAS"] else "",
            "TELEFONE": telefones[0] if len(telefones) > 0 else "",
            "E-MAIL": emails[0] if len(emails) > 0 else "",
            "CONTATO 2": contatos[1] if len(contatos) > 1 and contatos[1] != contatos[0] and contatos[1] not in ["CARLOS DURVAL DE MORAES", "CARLOS DURVAL DE MORAIS", "IZABELLA REBOREDO VERAS"] else "",
            "TELEFONE 2": telefones[1] if len(telefones) > 1 and telefones[1] != telefones[0] and telefones[1] != "(21) 2211-2540" and telefones[1] != "(21) 2211-2914" else "",
            "E-MAIL 2": emails[1] if len(emails) > 1 and emails[1] != emails[0] and emails[1] != "carlos.moraes@light.com.br" and emails[1] != "fatura.ccear@light.com.br" else "",
            "CONTATO 3": contatos[2] if len(contatos) > 2 and contatos[2] != contatos[0] and contatos[2] != contatos[1] and contatos[2] not in ["CARLOS DURVAL DE MORAES", "CARLOS DURVAL DE MORAIS", "IZABELLA REBOREDO VERAS"] else "",
            "TELEFONE 3": telefones[2] if len(telefones) > 2 and telefones[2] != telefones[0] and telefones[2] != telefones[1] else "",
            "E-MAIL 3": emails[2] if len(emails) > 2 and emails[2] != emails[0] and emails[2] != emails[1] else ""
        }
 
        return dados
    except Exception as e:
        print(f"Erro ao processar o PDF {caminho_pdf}: {e}")
        return {
            "CONTATO": "Erro", "TELEFONE": "Erro", "E-MAIL": "Erro",
            "CONTATO 2": "Erro", "TELEFONE 2": "Erro", "E-MAIL 2": "Erro",
            "CONTATO 3": "Erro", "TELEFONE 3": "Erro", "E-MAIL 3": "Erro"
        }
 
# Processar os PDFs e atualizar o DataFrame
def processar_pdfs():
    for idx, row in df_filtrado.iterrows():
        contrato = row.get("CCEAR", "Desconhecido")
 
        if pd.notna(contrato):
            print(f"Processando contrato {contrato}...")
            contrato_str = f"CCEAR_{int(contrato)}"  # Adiciona o prefixo "CCEAR_" e remove o ".0"
            arquivos_pdf = [f for f in os.listdir(diretorio_base) if contrato_str in f]
 
            if arquivos_pdf:
                caminho_pdf = os.path.join(diretorio_base, arquivos_pdf[0])
                dados_pdf = extrair_dados_pdf(caminho_pdf)
 
                # Atualizar apenas as colunas específicas no DataFrame filtrado
                for chave in ["CONTATO", "TELEFONE", "E-MAIL", "CONTATO 2", "TELEFONE 2", "E-MAIL 2", "CONTATO 3", "TELEFONE 3", "E-MAIL 3"]:
                    if chave in dados_pdf:
                        df_filtrado.at[idx, chave] = dados_pdf[chave]
            else:
                print(f"Arquivo PDF para o contrato {contrato_str} não encontrado.")
   
    # Salvar as atualizações na planilha
    df.update(df_filtrado)  # Atualiza apenas as linhas filtradas
    df.to_excel(planilha_path, index=False)
    print("Processo concluído! Planilha atualizada com os dados extraídos.")
 
# Executar o processamento
processar_pdfs()