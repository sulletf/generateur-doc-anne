import re
import docx
from docx.oxml import parse_xml
from docx.shared import Inches, Pt
import easyocr
import numpy as np
from PIL import Image
import pdfplumber
import streamlit as st


@st.cache_resource
def load_ocr_reader():
  return easyocr.Reader(['fr', 'en'])


# --- CLASSE DE GÉNÉRATION D'ACCESSIBILITÉ ---
class AccessibilityDocBuilder:

  def __init__(self):
    self.doc = docx.Document()
    self._setup_margins_and_styles()

  def _setup_margins_and_styles(self):
    for section in self.doc.sections:
      section.top_margin = Inches(0.8)
      section.bottom_margin = Inches(0.8)
      section.left_margin = Inches(0.8)
      section.right_margin = Inches(0.8)

    normal_style = self.doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(12)

  def add_heading(self, title):
    p = self.doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.add_run(title)

  def _insert_omml(self, paragraph, xml_content):
    omml = f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{xml_content}</m:oMath>'
    paragraph._p.append(parse_xml(omml))

  def add_math(self, paragraph, text):
    """Nettoie et formate automatiquement les formules mathématiques."""
    clean_text = text.replace(' = ', '=').replace(' =', '=').replace('= ', '=')
    units = ['cm', 'mm', 'm', 'km', 'L', 'dL', 'g', 'kg', '°', 'pieds']
    for unit in units:
      clean_text = clean_text.replace(f' {unit}', f'\u00a0{unit}')
    self._insert_omml(paragraph, f'<m:r><m:t>{clean_text}</m:t></m:r>')

  def add_smart_paragraph(self, paragraph_text):
    """Analyse automatiquement le texte pour isoler les portions mathématiques

    (segments entre crochets, valeurs avec unités, équations, etc.).
    """
    p = self.doc.add_paragraph()

    # Expression régulière pour détecter :
    # 1. Les segments géométriques (ex: [DN], [AB])
    # 2. Les valeurs numériques suivies d'unités (ex: 4 m, 4,20 cm, 20 pieds)
    # 3. Les expressions de type RS=9,7 cm ou des variables isolées si besoin
    # Vous pouvez enrichir ce pattern selon vos besoins
    pattern = re.compile(
        r'(\[[A-Z]{2}\]|\d+(?:[,\.]\d+)?\s*(?:cm|mm|m|km|pieds|°))',
        re.IGNORECASE,
    )

    last_end = 0
    for match in pattern.finditer(paragraph_text):
      start, end = match.span()
      # Texte normal avant la formule
      if start > last_end:
        p.add_run(paragraph_text[last_end:start])

      # Élément mathématique détecté automatiquement
      math_expr = match.group(0)
      self.add_math(p, math_expr)
      last_end = end

    # Reste du texte de la fin de ligne
    if last_end < len(paragraph_text):
      p.add_run(paragraph_text[last_end:])

    return p

  def save(self, filename):
    self.doc.save(filename)


# --- FONCTIONS D'EXTRACTION DE TEXTE ---
def extract_text_from_pdf(pdf_file):
  text = ''
  with pdfplumber.open(pdf_file) as pdf:
    for page in pdf.pages:
      extracted = page.extract_text()
      if extracted:
        text += extracted + '\n'
  return text


def extract_text_from_image(image_file):
  reader = load_ocr_reader()
  img = Image.open(image_file)
  img_np = np.array(img)
  results = reader.readtext(img_np, detail=0)
  return '\n'.join(results)


def extract_text_from_docx(docx_file):
  doc = docx.Document(docx_file)
  return '\n'.join([p.text for p in doc.paragraphs if p.text.strip()])


# --- INTERFACE STREAMLIT ---
st.set_page_config(
    page_title="Générateur Accessible Word", page_icon="📄", layout="centered"
)

st.title("📄 Générateur Automatique de Documents - Pour Anne")
st.markdown("""
Glissez simplement **un fichier ou une capture d'écran** de votre manuel ci-dessous. 
L'application extrait le texte, détecte automatiquement les mathématiques/unités et génère le fichier `.docx` prêt pour Anne, sans intervention manuelle !
""")

uploaded_file = st.file_uploader(
    "Importez votre document (PNG, JPG, PDF, TXT, DOCX) :",
    type=["png", "jpg", "jpeg", "pdf", "txt", "docx"],
)

if uploaded_file is not None:
  file_ext = uploaded_file.name.split('.')[-1].lower()

  with st.spinner(
      'Analyse du document et extraction automatique en cours...'
  ):
    if file_ext in ['png', 'jpg', 'jpeg']:
      texte_extrait = extract_text_from_image(uploaded_file)
      st.image(
          uploaded_file, caption='Image analysée', use_container_width=True
      )
    elif file_ext == 'pdf':
      texte_extrait = extract_text_from_pdf(uploaded_file)
    elif file_ext == 'docx':
      texte_extrait = extract_text_from_docx(uploaded_file)
    elif file_ext == 'txt':
      texte_extrait = uploaded_file.read().decode('utf-8')

  st.success('Extraction et traitement automatiques réussis !')

  # Aperçu du texte extrait brut pour information
  with st.expander("Voir le texte brut extrait du fichier"):
    st.text(texte_extrait)

  titre_exo = st.text_input(
      'Titre du document Word généré :', 'Exercice adaptatif'
  )

  if st.button('Télécharger le fichier Word pour Anne', type='primary'):
    builder = AccessibilityDocBuilder()
    builder.add_heading(titre_exo)

    lignes = texte_extrait.split('\n')
    for ligne in lignes:
      if ligne.strip():
        builder.add_smart_paragraph(ligne.strip())

    output_filename = 'Exercice_Anne_Automatique.docx'
    builder.save(output_filename)

    with open(output_filename, 'rb') as f:
      st.download_button(
          label='📥 Télécharger le fichier .docx final',
          data=f,
          file_name=output_filename,
          mime=(
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
          ),
      )
    st.success('Fichiers générés et prêts à être téléchargés !')
