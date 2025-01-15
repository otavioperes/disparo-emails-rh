import pandas as pd
import smtplib
from email.message import EmailMessage
import streamlit as st

# Configurações fixas do e-mail
EMAIL_ADDRESS = "otavioperes1233@gmail.com"  # Endereço de e-mail usado para enviar os contracheques
EMAIL_PASSWORD = "thlusanrdffmofth"          # Senha do e-mail gerada pelo Gmail

# Função para enviar e-mails
def enviar_email(destinatario, nome_arquivo, pdf_data):
    try:
        # Extrair o nome da pessoa do arquivo PDF
        nome_pessoa = nome_arquivo.replace(".pdf", "").replace("_", " ")

        # Configuração do e-mail
        msg = EmailMessage()
        msg["Subject"] = "Teste"
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = destinatario

        # Mensagem personalizada
        msg.set_content(f"""
        Olá {nome_pessoa},

        Este é um teste de envio de e-mails com anexos personalizados.

        Atenciosamente,
        Otavio
        """)

        # Anexar o arquivo PDF
        msg.add_attachment(pdf_data, maintype="application", subtype="pdf", filename=nome_arquivo)

        # Enviar o e-mail
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            smtp.send_message(msg)

        return f"E-mail enviado para {destinatario} com sucesso!"
    except Exception as e:
        return f"Erro ao enviar e-mail para {destinatario}: {e}"

# Interface com Streamlit
st.title("Envio de E-mails em Massa – RH")

# Passo 1: Upload dos PDFs
st.header("1. Upload dos Contracheques (PDFs)")
uploaded_pdfs = st.file_uploader(
    "Envie os arquivos PDF dos contracheques",
    accept_multiple_files=True,
    type=["pdf"]
)

pdf_files = {}
if uploaded_pdfs:
    for pdf in uploaded_pdfs:
        pdf_files[pdf.name] = pdf.read()
    st.success(f"{len(pdf_files)} arquivos PDF carregados com sucesso.")

# Passo 2: Seleção da Planilha Excel
st.header("2. Seleção da Planilha de E-mails")
planilha_file = st.file_uploader("Faça o upload da planilha Excel", type=["xlsx"])
email_map = {}

if planilha_file:
    planilha = pd.read_excel(planilha_file)
    st.write("Pré-visualização da planilha:")
    st.dataframe(planilha.head())

    # Criar dicionário de e-mails
    email_map = dict(zip(planilha["Nome Completo"], planilha["E-mail"]))

# Passo 3: Visualização da Associação
st.header("3. Associação entre PDFs e Planilha")
if uploaded_pdfs and planilha_file:
    associacao = []
    for pdf_name in pdf_files.keys():
        nome_completo = pdf_name.replace(".pdf", "")
        email = email_map.get(nome_completo, "E-mail não encontrado")
        associacao.append({"Arquivo PDF": pdf_name, "Nome": nome_completo, "E-mail": email})
    
    df_associacao = pd.DataFrame(associacao)
    st.write("Associação entre arquivos PDF e e-mails:")
    st.dataframe(df_associacao)

# Passo 4: Envio de E-mails
st.header("4. Envio de E-mails")

# Inicializar os logs de envio no session_state
if "log_envio" not in st.session_state:
    st.session_state.log_envio = {"Enviado com Sucesso": [], "Erro no Envio": []}

if st.button("Iniciar Envio"):
    if not uploaded_pdfs or not planilha_file:
        st.error("Por favor, envie os PDFs e a planilha antes de iniciar o envio.")
    else:
        # Resetar os logs para uma nova execução
        st.session_state.log_envio = {"Enviado com Sucesso": [], "Erro no Envio": []}

        # Processar os envios
        for _, row in df_associacao.iterrows():
            if row["E-mail"] != "E-mail não encontrado":
                resultado = enviar_email(row["E-mail"], row["Arquivo PDF"], pdf_files[row["Arquivo PDF"]])
                st.session_state.log_envio["Enviado com Sucesso"].append(resultado)
            else:
                st.session_state.log_envio["Erro no Envio"].append(
                    f"Erro: E-mail não enviado para {row['Nome']} ({row['Arquivo PDF']}). A associação não foi encontrada."
                )

# Resultados do Envio
if st.session_state.log_envio["Enviado com Sucesso"] or st.session_state.log_envio["Erro no Envio"]:
    st.header("4.1 Resultados do Envio")
    filtro = st.radio("Filtrar por:", options=["Todos", "Enviado com Sucesso", "Erro no Envio"])

    if filtro == "Todos":
        st.subheader("Enviados com Sucesso:")
        for log in st.session_state.log_envio["Enviado com Sucesso"]:
            st.write(log)
        st.subheader("Erros no Envio:")
        for log in st.session_state.log_envio["Erro no Envio"]:
            st.write(log)
    elif filtro == "Enviado com Sucesso":
        st.subheader("Enviados com Sucesso:")
        for log in st.session_state.log_envio["Enviado com Sucesso"]:
            st.write(log)
    elif filtro == "Erro no Envio":
        st.subheader("Erros no Envio:")
        for log in st.session_state.log_envio["Erro no Envio"]:
            st.write(log)
