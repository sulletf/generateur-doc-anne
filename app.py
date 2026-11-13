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

  def _insert_omml(self, paragraph, xml_content):
    omml = f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{xml_content}</m:oMath>'
    paragraph._p.append(parse_xml(omml))

  def add_math(self, paragraph, text):
    """Convertit une expression mathématique en objet OMML propre."""
    clean_text = text.replace(' = ', '=').replace(' =', '=').replace('= ', '=')
    
    # Gestion des puissances au carré (ex: ER² ou ER^2)
    clean_text = clean_text.replace('^2', '²')
    if '²' in clean_text:
      parts_sup = clean_text.split('²')
      base = parts_sup[0]
      rest = parts_sup[1] if len(parts_sup) > 1 else ''
      xml_sup = f'<m:sSup><m:e><m:r><m:t>{base}</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>{rest}'
      self._insert_omml(paragraph, xml_sup)
    else:
      self._insert_omml(paragraph, f'<m:r><m:t>{clean_text}</m:t></m:r>')

  def add_smart_paragraph(self, paragraph_text):
    """Analyse le texte en se basant sur les délimitors $...$ du PDF source

    ainsi que sur une détection automatique des segments ou unités.
    """
    p = self.doc.add_paragraph()

    # Si le texte contient des $ (provenant du PDF), on découpe par paires de $
    if '$' in paragraph_text:
      parts = paragraph_text.split('$')
      for i, part in enumerate(parts):
        if not part:
          continue
        if i % 2 == 1:
          # Impair = à l'intérieur des $, c'est une formule mathématique
          self.add_math(p, part)
        else:
          # Pair = texte normal
          p.add_run(part)
    else:
      # Fallback : si pas de $, on applique une détection automatique par regex
      pattern = re.compile(
          r'(\[[A-Z\.]{2,}\]|\b[A-Za-z]+(?:\s*²|\^2)?\s*=\s*[^,\.\s]+|\d+(?:[,\.]\d+)?\s*(?:cm|mm|m|km|pieds|°))',
          re.IGNORECASE,
      )
      last_end = 0
      for match in pattern.finditer(paragraph_text):
        start, end = match.span()
        if start > last_end:
          p.add_run(paragraph_text[last_end:start])
        self.add_math(p, match.group(0))
        last_end = end
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
L'application extrait le texte, convertit les mathématiques en objets natifs et génère le fichier `.docx` propre !
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

  with st.expander("Voir le texte brut extrait du fichier"):
    st.text(texte_extrait)

  if st.button('Télécharger le fichier Word pour Anne', type='primary'):
    builder = AccessibilityDocBuilder()

    lignes = texte_extrait.split('\n')
    for ligne in lignes:
      if ligne.strip():
        builder.add_smart_paragraph(ligne.strip())

    output_filename = 'Document_Anne_Accessible.docx'
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
    st.success('Fichier généré et prêt à être téléchargé !')
