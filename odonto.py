import os
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor

st.set_page_config(
    page_title="Telessaúde UEA - Triagem Odontológica",
    page_icon="🦷",
    layout="wide"
)

def validar_cpf(cpf):
    cpf = ''.join(filter(str.isdigit, str(cpf)))
    if len(cpf) != 11 or cpf == cpf[0] * 11:
        return False
    soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
    digito1 = (soma * 10) % 11
    if digito1 == 10:
        digito1 = 0
    if digito1 != int(cpf[9]):
        return False
    soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
    digito2 = (soma * 10) % 11
    if digito2 == 10:
        digito2 = 0
    if digito2 != int(cpf[10]):
        return False
    return True

st.title("🦷 Telessaúde UEA · Teleatendimento Odontológico")
st.markdown("Preencha os dados abaixo para gerar a solicitação oficial de teleatendimento odontológico.")

# Abas organizadas
tab1, tab2, tab3, tab4 = st.tabs([
    "1. Serviço e Solicitante", "2. Paciente", "3. Caso Clínico", "4. Urgência e Anexos"
])

with tab1:
    st.subheader("Tipo de Serviço")
    servico = st.selectbox("Qual serviço você deseja? *", [
        "Teleconsultoria (via texto)", 
        "Teleinterconsulta (via vídeo, em tempo real com participação do paciente) - a ser agendado"
    ])
    
    st.subheader("Informações do Solicitante / Cirurgião-Dentista (CD)")
    prof_nome = st.text_input("Nome do solicitante *")
    prof_cpf = st.text_input("CPF *")
    prof_cro = st.text_input("CRO do solicitante *")
    prof_unidade = st.text_input("Nome da unidade de saúde *")
    prof_email = st.text_input("E-mail do solicitante *")
    prof_whatsapp = st.text_input("WhatsApp do solicitante *")
    prof_municipio = st.text_input("Qual seu município/Ponto de Telessaúde UEA *")

with tab2:
    st.subheader("Identificação do Paciente")
    pac_nome = st.text_input("Nome Completo do Paciente *")
    pac_idade = st.text_input("Idade do paciente *")
    
    pac_menor = st.selectbox("Paciente menor de idade? *", ["Não", "Sim"])
    
    pac_email = st.text_input("E-mail do paciente")
    pac_whatsapp = st.text_input("WhatsApp do paciente")
    pac_cpf = st.text_input("Nº do CPF do paciente *")
    
    col1, col2 = st.columns(2)
    with col1:
        pac_peso = st.text_input("Peso")
        pac_altura = st.text_input("Altura")
    with col2:
        pac_sexo = st.selectbox("Sexo *", ["", "Masculino", "Feminino", "Outro"])
        if pac_sexo == "Outro":
            pac_sexo_outro = st.text_input("Especifique o sexo:")
        else:
            pac_sexo_outro = ""

    pac_cidade = st.text_input("Cidade *")
    pac_territorio = st.selectbox("Localidade do Paciente (território): *", [
        "Não faz parte de nenhum dos grupos", 
        "Quilombola", 
        "Área Ribeirinha", 
        "DSEI"
    ])

with tab3:
    st.subheader("Descrição do Caso Clínico")
    caso_desc = st.text_area("Descrição do caso Clínico * (Queixa principal, evolução, hipótese diagnóstica)")
    caso_duvida = st.text_area("Qual a sua dúvida referente a esse caso? * (Adicione aqui sua dúvida)")
    especialidade_desejada = st.text_input("Tem uma especialidade desejada? qual? *")

with tab4:
    st.subheader("Urgência e Envio de Exames / Fotos")
    urgencia = st.selectbox("Configura urgência? *", ["Não", "Sim", "Outro"])
    if urgencia == "Outro":
        urgencia_outro = st.text_input("Especifique a urgência:")
    else:
        urgencia_outro = ""

    st.info("💡 Insira as fotos em um documento Word e em seguida faça o upload abaixo.")
    imagens_exames = st.file_uploader("Faça upload do arquivo (Máx. 10 MB)", type=["docx", "doc", "png", "jpg", "jpeg"], accept_multiple_files=True)
    
    # --- SEÇÃO DE 5 FOTOS DE EXEMPLO (RESPONSIVO) ---
    st.markdown("---")
    st.markdown("### 📸 Exemplos de como tirar as fotos:")
    st.markdown("Veja abaixo o padrão ideal para o envio das imagens clínicas:")

    # Cria 5 colunas (no PC ficam lado a lado, no celular empilham automaticamente)
    c1, c2, c3, c4, c5 = st.columns(5)
    
    with c1:
        if os.path.exists("./fotos/foto1.jpeg"):
            st.image("./fotos/foto1.jpeg", caption="Exemplo 1", width=150)
        else:
            st.info("foto1.jpeg ausente")

    with c2:
        if os.path.exists("./fotos/foto2.jpeg"):
            st.image("./fotos/foto2.jpeg", caption="Exemplo 2", width=150)
        else:
            st.info("foto2.jpeg ausente")

    with c3:
        if os.path.exists("./fotos/foto3.jpeg"):
            st.image("./fotos/foto3.jpeg", caption="Exemplo 3", width=150)
        else:
            st.info("foto3.jpeg ausente")

    with c4:
        if os.path.exists("./fotos/foto4.jpeg"):
            st.image("./fotos/foto4.jpeg", caption="Exemplo 4", width=150)
        else:
            st.info("foto4.jpeg ausente")

    with c5:
        if os.path.exists("./fotos/foto 5.jpeg"):
            st.image("./fotos/foto 5.jpeg", caption="Exemplo 5", width=150)
        else:
            st.info("foto5.jpeg ausente")
    # -----------------------------------------------

    st.markdown("---")
    if st.button("📄 Gerar Documento Word (.docx)", type="primary"):
        if not prof_nome.strip() or not pac_nome.strip() or not caso_duvida.strip() or not prof_cpf.strip() or not pac_cpf.strip():
            st.error("❌ Erro: Preencha todos os campos obrigatórios marcados com (*), incluindo os CPFs.")
        elif prof_cpf and not validar_cpf(prof_cpf):
            st.error("❌ Erro: O CPF do solicitante é inválido.")
        elif pac_cpf and not validar_cpf(pac_cpf):
            st.error("❌ Erro: O CPF do paciente é inválido.")
        else:
            try:
                doc = Document()
                for section in doc.sections:
                    section.top_margin = Inches(1)
                    section.bottom_margin = Inches(1)
                    section.left_margin = Inches(1)
                    section.right_margin = Inches(1)

                if os.path.exists("logo.png"):
                    try:
                        doc.add_picture("logo.png", width=Inches(1.8))
                    except Exception:
                        pass

                p_title = doc.add_paragraph()
                r_title = p_title.add_run("SOLICITAÇÃO DE TELEATENDIMENTO — ODONTOLOGIA")
                r_title.bold = True
                r_title.font.size = Pt(16)
                r_title.font.color.rgb = RGBColor(23, 56, 50)

                p_sub = doc.add_paragraph()
                r_sub = p_sub.add_run("Telessaúde UEA")
                r_sub.italic = True
                r_sub.font.size = Pt(11)
                r_sub.font.color.rgb = RGBColor(44, 110, 99)

                doc.add_paragraph().paragraph_format.space_after = Pt(10)

                sexo_final = f"Outro ({pac_sexo_outro})" if pac_sexo == "Outro" else pac_sexo
                urgencia_final = f"Outro ({urgencia_outro})" if urgencia == "Outro" else urgencia

                dados_gerais = [
                    ("DADOS DO SERVIÇO", [
                        ("Tipo de serviço desejado", servico)
                    ]),
                    ("PROFISSIONAL SOLICITANTE (CD)", [
                        ("Nome do solicitante", prof_nome), ("CPF", prof_cpf), ("CRO", prof_cro),
                        ("Unidade de saúde", prof_unidade), ("E-mail", prof_email), 
                        ("WhatsApp", prof_whatsapp), ("Município / Ponto UEA", prof_municipio)
                    ]),
                    ("IDENTIFICAÇÃO DO PACIENTE", [
                        ("Nome do paciente", pac_nome), ("Idade", pac_idade), ("Menor de idade", pac_menor),
                        ("E-mail do paciente", pac_email), ("WhatsApp do paciente", pac_whatsapp),
                        ("CPF do paciente", pac_cpf), ("Peso", pac_peso), ("Altura", pac_altura),
                        ("Sexo", sexo_final), ("Cidade", pac_cidade), ("Território / Localidade", pac_territorio)
                    ]),
                    ("CASO CLÍNICO E SOLICITAÇÃO", [
                        ("Descrição do caso clínico", caso_desc), ("Dúvida clínica", caso_duvida),
                        ("Especialidade desejada", especialidade_desejada), ("Configura urgência", urgencia_final)
                    ])
                ]

                for sec_title, campos in dados_gerais:
                    p_sec = doc.add_paragraph()
                    r_sec = p_sec.add_run(sec_title)
                    r_sec.bold = True
                    r_sec.font.size = Pt(13)
                    r_sec.font.color.rgb = RGBColor(23, 56, 50)
                    p_sec.paragraph_format.space_before = Pt(14)
                    p_sec.paragraph_format.space_after = Pt(4)

                    for label, val in campos:
                        if val and str(val).strip() != "":
                            p_field = doc.add_paragraph()
                            p_field.paragraph_format.space_after = Pt(3)
                            r_label = p_field.add_run(f"{label}: ")
                            r_label.bold = True
                            r_label.font.size = Pt(11)
                            r_val = p_field.add_run(str(val))
                            r_val.font.size = Pt(11)

                if imagens_exames:
                    for idx, img_file in enumerate(imagens_exames):
                        doc.add_page_break()
                        p_desc = doc.add_paragraph()
                        r_d = p_desc.add_run(f"DOCUMENTO / IMAGEM ANEXADA {idx + 1}")
                        r_d.bold = True
                        r_d.font.size = Pt(13)
                        r_d.font.color.rgb = RGBColor(23, 56, 50)
                        p_desc.paragraph_format.space_after = Pt(14)
                        try:
                            if img_file.name.lower().endswith(('png', 'jpg', 'jpeg')):
                                doc.add_picture(img_file, width=Inches(5.5))
                            else:
                                doc.add_paragraph(f"[Arquivo anexado: {img_file.name}]")
                        except Exception:
                            pass

                file_name = f"Odontologia_{pac_nome.replace(' ', '_')}.docx"
                doc.save(file_name)
                
                st.success("✅ Documento Word de Odontologia gerado com sucesso!")
                with open(file_name, "rb") as f:
                    st.download_button(
                        label="📥 Baixar Documento Word (.docx)",
                        data=f,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary"
                    )
            except Exception as e:
                st.error(f"❌ Erro ao gerar documento: {str(e)}")
