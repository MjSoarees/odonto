import os
import base64
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml

st.set_page_config(
    page_title="Telessaúde UEA - Teleatendimento Odontológico Especializado",
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

def img_to_base64(path):
    if os.path.exists(path):
        with open(path, "rb") as f:
            data = f.read()
        return base64.b64encode(data).decode()
    return ""

def adicionar_cabecalho_e_marca_dagua(doc, logo_marca_dagua, logos_topo):
    section = doc.sections[0]
    header = section.header
    header_para = header.paragraphs[0]
    header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 1. Adiciona as 3 logos centralizadas no topo do cabeçalho
    for logofile in logos_topo:
        if os.path.exists(logofile):
            try:
                run = header_para.add_run()
                run.add_picture(logofile, width=Inches(0.9))
                header_para.add_run("   ") 
            except Exception:
                pass

    # 2. Adiciona a marca d'água grande de fundo centralizada
    if logo_marca_dagua and os.path.exists(logo_marca_dagua):
        try:
            watermark_p = header.add_paragraph()
            watermark_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_wm = watermark_p.add_run()
            pic = run_wm.add_picture(logo_marca_dagua, width=Inches(5.5))
            
            inline = pic._inline
            
            # XML robusto para posicionar a imagem em absoluto ao centro e atrás do texto
            watermark_xml = parse_xml(f'''
                <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" 
                           simplePos="0" relativeHeight="251658240" behindDoc="1" locked="1" layoutInCell="1" allowOverlap="1">
                    <wp:simplePos x="0" y="0"/>
                    <wp:positionH relativeFrom="page">
                        <wp:align>center</wp:align>
                    </wp:positionH>
                    <wp:positionV relativeFrom="page">
                        <wp:align>center</wp:align>
                    </wp:positionV>
                    <wp:extent cx="5000000" cy="5000000"/>
                    <wp:effectExtent l="0" t="0" r="0" b="0"/>
                    <wp:wrapNone/>
                    <wp:docPr id="999" name="Watermark"/>
                    <wp:cNvGraphicFramePr/>
                    <graphic xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
                        <graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:nvPr><pic:cNvPr id="0" name="Watermark"/><pic:cNvPicPr/></pic:nvPr>
                                <pic:blipFill>
                                    <a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="{inline.graphic.graphicData.uri}"/>
                                    <a:stretch><a:fillRect/></a:stretch>
                                </pic:blipFill>
                                <pic:spPr>
                                    <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                        <a:off x="0" y="0"/>
                                        <a:ext cx="5000000" cy="5000000"/>
                                    </a:xfrm>
                                    <a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect"><a:avLst/></a:prstGeom>
                                </pic:spPr>
                            </pic:pic>
                        </graphicData>
                    </graphic>
                </wp:anchor>
            ''')
            inline.getparent().replace(inline, watermark_xml)
        except Exception:
            pass

# Carrega as logos da pasta "fotos" em base64 para a interface web
img1_b64 = img_to_base64("./fotos/logo3.png") or img_to_base64("logo3.png")
img2_b64 = img_to_base64("./fotos/logo2.png") or img_to_base64("logo2.png")
img3_b64 = img_to_base64("./fotos/logo1.png") or img_to_base64("logo1.png")

# --- EXIBIÇÃO DAS 3 LOGOS CENTRALIZADAS E SEMPRE LADO A LADO (PC E MOBILE) ---
logos_html = '<div style="display: flex; justify-content: center; align-items: center; gap: 15px; margin-bottom: 15px;">'
if img1_b64:
    logos_html += f'<img src="data:image/png;base64,{img1_b64}" style="height: 50px; width: auto;" />'
if img2_b64:
    logos_html += f'<img src="data:image/png;base64,{img2_b64}" style="height: 50px; width: auto;" />'
if img3_b64:
    logos_html += f'<img src="data:image/png;base64,{img3_b64}" style="height: 50px; width: auto;" />'
logos_html += '</div>'

st.markdown(logos_html, unsafe_allow_html=True)
# ----------------------------------------------------------------------------

st.markdown("<h3 style='text-align: center; font-size: 1.25rem;'>🦷 Telessaúde UEA · Teleatendimento Odontológico Especializado</h3>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Preencha os dados abaixo para gerar a solicitação oficial de teleatendimento odontológico.</p>", unsafe_allow_html=True)

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

    c1, c2, c3, c4, c5 = st.columns(5)
    
    with c1:
        if os.path.exists("./fotos/foto1.jpeg"):
            st.image("./fotos/foto1.jpeg", caption="Exemplo 1", width=150)
        elif os.path.exists("foto1.jpeg"):
            st.image("foto1.jpeg", caption="Exemplo 1", width=150)
        else:
            st.info("foto1.jpeg ausente")

    with c2:
        if os.path.exists("./fotos/foto2.jpeg"):
            st.image("./fotos/foto2.jpeg", caption="Exemplo 2", width=150)
        elif os.path.exists("foto2.jpeg"):
            st.image("foto2.jpeg", caption="Exemplo 2", width=150)
        else:
            st.info("foto2.jpeg ausente")

    with c3:
        if os.path.exists("./fotos/foto3.jpeg"):
            st.image("./fotos/foto3.jpeg", caption="Exemplo 3", width=150)
        elif os.path.exists("foto3.jpeg"):
            st.image("foto3.jpeg", caption="Exemplo 3", width=150)
        else:
            st.info("foto3.jpeg ausente")

    with c4:
        if os.path.exists("./fotos/foto4.jpeg"):
            st.image("./fotos/foto4.jpeg", caption="Exemplo 4", width=150)
        elif os.path.exists("foto4.jpeg"):
            st.image("foto4.jpeg", caption="Exemplo 4", width=150)
        else:
            st.info("foto4.jpeg ausente")

    with c5:
        if os.path.exists("./fotos/foto 5.jpeg"):
            st.image("./fotos/foto 5.jpeg", caption="Exemplo 5", width=150)
        elif os.path.exists("foto5.jpeg"):
            st.image("foto5.jpeg", caption="Exemplo 5", width=150)
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

                logos_para_word = ["./fotos/logo1.png", "./fotos/logo2.png", "./fotos/logo3.png", "logo1.png", "logo2.png", "logo3.png"]
                logo_principal = "./fotos/logo1.png" if os.path.exists("./fotos/logo1.png") else "logo1.png"
                
                # Aplica o cabeçalho e a marca d'água robusta de fundo
                adicionar_cabecalho_e_marca_dagua(doc, logo_principal, logos_para_word)

                p_title = doc.add_paragraph()
                # Adiciona espaçamento antes do título para dar respiro entre as logos do topo e o texto
                p_title.paragraph_format.space_before = Pt(24)
                r_title = p_title.add_run("SOLICITAÇÃO DE TELEATENDIMENTO — ODONTOLOGIA")
                r_title.bold = True
                r_title.font.size = Pt(14)
                r_title.font.color.rgb = RGBColor(23, 56, 50)

                p_sub = doc.add_paragraph()
                r_sub = p_sub.add_run("Telessaúde UEA")
                r_sub.italic = True
                r_sub.font.size = Pt(10)
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
                    r_sec.font.size = Pt(12)
                    r_sec.font.color.rgb = RGBColor(23, 56, 50)
                    p_sec.paragraph_format.space_before = Pt(12)
                    p_sec.paragraph_format.space_after = Pt(4)

                    for label, val in campos:
                        if val and str(val).strip() != "":
                            p_field = doc.add_paragraph()
                            p_field.paragraph_format.space_after = Pt(3)
                            r_label = p_field.add_run(f"{label}: ")
                            r_label.bold = True
                            r_label.font.size = Pt(10)
                            r_val = p_field.add_run(str(val))
                            r_val.font.size = Pt(10)

                if imagens_exames:
                    for idx, img_file in enumerate(imagens_exames):
                        doc.add_page_break()
                        p_desc = doc.add_paragraph()
                        r_d = p_desc.add_run(f"DOCUMENTO / IMAGEM ANEXADA {idx + 1}")
                        r_d.bold = True
                        r_d.font.size = Pt(12)
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
