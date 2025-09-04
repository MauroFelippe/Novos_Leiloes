# ⚡ Automação de Contratos - Light

Scripts em **Python** para automatizar processos relacionados aos contratos de energia da Light.  
O projeto cobre desde a **extração de contatos de PDFs** até o **disparo automatizado de e-mails via Outlook**.

---

## 🚀 Funcionalidades

### 🔎 Buscar Contatos (Leilões 23, 30 e 34)
- Lê planilhas Excel de contratos.
- Filtra contratos por leilão específico.
- Localiza arquivos PDF correspondentes (CCEAR).
- Extrai até **3 contatos, telefones e e-mails** de cada contrato.
- Atualiza automaticamente a planilha original com as informações.

### 📧 Disparar Emails
- Lê a base de contratos (Excel).
- Filtra registros com `STATUS = "NÃO RECEBIDO"`.
- Agrupa os contratos por **Consórcio/Empresa**.
- Gera automaticamente um **e-mail em HTML** com os dados formatados em tabela.
- Anexa a **Ficha Cadastral** e os contratos PDF.
- Envia via **Outlook**, com cópia para os setores responsáveis.

---

## 🛠️ Tecnologias Utilizadas
- [Python 3.x](https://www.python.org/)
- [pandas](https://pandas.pydata.org/) → manipulação de dados em Excel
- [openpyxl](https://openpyxl.readthedocs.io/) → leitura/escrita em planilhas
- [PyMuPDF (fitz)](https://pymupdf.readthedocs.io/) → leitura e extração de texto em PDFs
- [re (Regex)](https://docs.python.org/3/library/re.html) → tratamento de padrões em textos
- [pywin32](https://pypi.org/project/pywin32/) → integração com Microsoft Outlook para envio de e-mails

---

## 📂 Estrutura dos Arquivos

📁 Automacao_Contratos
┣ 📜 Buscar Contatos 23 LEN.py
┣ 📜 Buscar Contatos 30 LEN.py
┣ 📜 Buscar Contatos 34 LEN.py
┣ 📜 Disparar Emails.py
┣ 📜 README.md ← (este arquivo)


---

## ▶️ Como Usar

### 1. Preparar Ambiente
Instale as dependências:
```bash
pip install pandas openpyxl PyMuPDF pywin32

2. Ajustar Caminhos

Nos scripts .py, atualize os caminhos de acordo com a sua máquina/pasta:

planilha_path = r"C:\\Users\\XXXX\\Light\\...\\Arquivo.xlsx"
diretorio_base = r"C:\\Users\\XXXX\\Light\\...\\Pasta dos PDFs"

3. Executar Scripts

Para extrair contatos de um leilão específico:

python "Buscar Contatos 23 LEN.py"
python "Buscar Contatos 30 LEN.py"
python "Buscar Contatos 34 LEN.py"

Para disparar os e-mails automáticos

python "Disparar Emails.py"

⚠️ Observações Importantes

O Outlook precisa estar instalado e configurado para que o envio de e-mails funcione.

Os PDFs devem estar organizados nas pastas correspondentes a cada leilão (23, 30, 34 etc).

O script já aplica regras de negócio específicas, como ignorar contatos/e-mails indesejados.

Recomenda-se rodar os scripts em ambiente controlado antes de uso em produção.

👨‍💻 Autor

Mauro Felippe Telles Junior
