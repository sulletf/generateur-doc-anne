import docx
from docx.oxml import parse_xml
from docx.shared import Inches, Pt
import easyocr
import numpy as np
from PIL import Image
import pdfplumber
import streamlit as st

# Charger le moteur OCR une seule fois en cache pour booster la rapidité
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
    clean_text = text.replace(' = ', '=').replace(' =', '=').replace('= ', '=')
    units = ['cm', 'mm', 'm', 'km', 'L', 'dL', 'g', 'kg', '°', 'pieds']
    for unit in units:
      clean_text = clean_text.replace(f' {unit}', f'\u00a0{unit}')
    self._insert_omml(paragraph, f'<m:r><m:t>{clean_text}</m:t></m:r>')

  def add_paragraph_with_mixed_content(self, elements):
    p = self.doc.add_paragraph()
    for el_type, *args in elements:
      if el_type == 'text':
        p.add_run(args[0])
      elif el_type == 'math':
        self.add_math(p, args[0])
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

st.title("📄 Générateur de documents adaptatifs - Pour Anne")
st.markdown("""
Cette application permet de charger **n'importe quel type de fichier** (PNG, JPG, PDF, TXT, DOCX), d'en extraire le texte et de générer un fichier Word (.docx) accessible pour la plage tactile d'Anne.
""")

st.subheader("1. Charger votre fichier (Image, PDF, Word, Texte)")

uploaded_file = st.file_uploader(
    "Glissez-déposez votre document ici :",
    type=["png", "jpg", "jpeg", "pdf", "txt", "docx"],
)

texte_extrait = ""

if uploaded_file is not None:
  file_ext = uploaded_file.name.split(".")[-1].lower()

  with st.spinner("Extraction du texte en cours..."):
    if file_ext in ["png", "jpg", "jpeg"]:
      texte_extrait = extract_text_from_image(uploaded_file)
      st.image(
          uploaded_file, caption="Image importée", use_container_width=True
      )
    elif file_ext == "pdf":
      texte_extrait = extract_text_from_pdf(uploaded_file)
    elif file_ext == "docx":
      texte_extrait = extract_text_from_docx(uploaded_file)
    elif file_ext == "txt":
      texte_extrait = uploaded_file.read().decode("utf-8")

  st.success("Texte extrait avec succès !")

st.subheader("2. Révision et balisage du texte")
st.info(
    "Vous pouvez réviser ci-dessous le texte extrait. Entourez les expressions"
    " ou grandeurs mathématiques avec `[M]` et `[/M]` si besoin (ex :"
    " `[M]4 m[/M]`)."
)

texte_a_traiter = st.text_area(
    "Texte brut extrait (modifiable) :",
    value=texte_extrait,
    height=250,
    placeholder=(
        "Le texte extrait s'affichera automatiquement ici. Vous pouvez aussi le"
        " taper à la main."
    ),
)

titre_exo = st.text_input("Titre du document Word :", "Exercice adaptatif")

if st.button("Générer le document Word", type="primary"):
  if texte_a_traiter.strip():
    builder = AccessibilityDocBuilder()
    builder.add_heading(titre_exo)

    lignes = texte_a_traiter.split("\n")
    for ligne in lignes:
      if not ligne.strip():
        continue
      elements = []
      parts = ligne.split("[M]")
      for part in parts:
        if "[/M]" in part:
          math_content, rest_text = part.split("[/M]", 1)
          elements.append(("math", math_content))
          if rest_text:
            elements.append(("text", rest_text))
        else:
          if part:
            elements.append(("text", part))

      builder.add_paragraph_with_mixed_content(elements)

    output_filename = "Exercice_Anne_Accessible.docx"
    builder.save(output_filename)

    with open(output_filename, "rb") as f:
      st.download_button(
          label="📥 Télécharger le fichier .docx pour Anne",
          data=f,
          file_name=output_filename,
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )
    st.success("Document généré avec succès !")
  else:
    st.error("Aucun texte à traiter n'a été trouvé.")
