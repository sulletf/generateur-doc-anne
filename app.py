import os
import docx
from docx.oxml import parse_xml
from docx.shared import Inches, Pt
from google import genai
from google.genai import types
import pdfplumber
from PIL import Image
import streamlit as st

# --- CONFIGURATION DE L'API GEMINI ---
# Pour sécuriser votre clé, vous pouvez la stocker dans st.secrets["GEMINI_API_KEY"]
# ou laisser un champ de saisie dans l'interface pour les utilisateurs.
st.set_page_config(
    page_title="Générateur Accessible Anne", page_icon="📄", layout="centered"
)

st.title("📄 Générateur de Documents Adaptés - Pour Anne")
st.markdown("""
Cette application permet aux professeurs et aux parents de transformer n'importe quel support (PDF, image, texte) en un **document Word accessible** (avec mode équation OMML et règles d'accessibilité pour Anne).
""")

# Saisie de la clé API (recommandé pour le partage avec d'autres profs/parents)
api_key_input = st.text_input(
    "Entrez votre clé API Google Gemini (ou configurez-la dans Streamlit Secrets)"
    " :",
    type="password",
)

# Utilisation de la clé fournie ou des secrets Streamlit
api_key = api_key_input or st.secrets.get("GEMINI_API_KEY", "")


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
    """Convertit une expression mathématique en objet OMML propre, sans espaces autour de =."""
    clean_text = text.replace(' = ', '=').replace(' =', '=').replace('= ', '=')

    # Gestion des puissances au carré
    if '²' in clean_text or '^2' in clean_text:
      clean_text = clean_text.replace('^2', '²')
      parts = clean_text.split('²')
      base = parts[0]
      rest = parts[1] if len(parts) > 1 else ''
      xml_sup = f'<m:sSup><m:e><m:r><m:t>{base}</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>{rest}'
      self._insert_omml(paragraph, xml_sup)
    else:
      self._insert_omml(paragraph, f'<m:r><m:t>{clean_text}</m:t></m:r>')

  def add_paragraph_with_mixed_content(self, parts):
    """Ajoute un paragraphe contenant un mix de texte normal et de balises mathématiques."""
    p = self.doc.add_paragraph()
    for p_type, content in parts:
      if p_type == 'text':
        p.add_run(content)
      elif p_type == 'math':
        self.add_math(p, content)
    return p

  def save(self, filename):
    self.doc.save(filename)


# --- FONCTION D'INTERPRETATION PAR GEMINI ---
def parse_document_with_gemini(uploaded_file, file_bytes, mime_type, api_key):
  client = genai.Client(api_key=api_key)

  prompt = """
    Analyse ce document pédagogique de mathématiques destiné à une élève aveugle (Anne).
    Extrait tout le texte utile de manière propre. S'il y a des doublons (ex: deux fiches identiques sur la page), ne garde qu'un seul exemplaire propre.
    
    Tu dois structurer ta réponse sous forme de paragraphes. Pour chaque paragraphe, sépare clairement le texte brut des expressions mathématiques, des segments géométriques (ex: [LO]), des formules (ex: BO=37, ER²=...) et des égalités.
    
    Renvoie le résultat au format JSON strict (une liste de listes de tuples [type, contenu]), où le type est soit "text", soit "math".
    Exemple de structure attendue pour une phrase :
    [
       [ ["text", "Le triangle "], ["math", "BOL"], ["text", " est rectangle en "], ["math", "L"], ["text", "."] ]
    ]
    Assure-toi de bien isoler en "math" toutes les variables, lettres, segments entre crochets, valeurs avec unités et équations.
    """

  # Appel de Gemini selon le type de fichier
  if mime_type.startswith('image/'):
    image = Image.open(uploaded_file)
    response = client.models.generate_content(
        model='gemini-2.5-flash', contents=[image, prompt]
    )
  elif mime_type == 'application/pdf':
    # Pour un PDF, on peut l'envoyer directement ou extraire les pages
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=[
            types.Part.from_bytes(data=file_bytes, mime_type='application/pdf'),
            prompt,
        ],
    )
  else:
    text_content = file_bytes.decode('utf-8')
    response = client.models.generate_content(
        model='gemini-2.5-flash', contents=[text_content, prompt]
    )

  return response.text


# --- INTERFACE UTILISATEUR ---
uploaded_file = st.file_uploader(
    "Glissez-déposez le document de cours ou d'exercice (PDF, Image, TXT) :",
    type=["png", "jpg", "jpeg", "pdf", "txt"],
)

if uploaded_file is not None and api_key:
  file_bytes = uploaded_file.read()
  file_ext = uploaded_file.name.split(".")[-1].lower()
  mime_types = {
      "png": "image/png",
      "jpg": "image/jpeg",
      "jpeg": "image/jpeg",
      "pdf": "application/pdf",
      "txt": "text/plain",
  }
  mime_type = mime_types.get(file_ext, "text/plain")

  if st.button(
      "Générer le document Word accessible pour Anne", type="primary"
  ):
    with st.spinner("Gemini analyse le document avec les règles d'Anne..."):
      try:
        # Appel à l'intelligence de Gemini
        raw_result = parse_document_with_gemini(
            uploaded_file, file_bytes, mime_type, api_key
        )

        # Nettoyage et conversion de la réponse de Gemini en structure Python
        import json

        # Extraction propre du bloc JSON si Gemini a mis du texte autour
        json_str = raw_result
        if "```json" in raw_result:
          json_str = raw_result.split("```json")[1].split("```")[0].strip()
        elif "```" in raw_result:
          json_str = raw_result.split("```")[1].split("```")[0].strip()

        paragraphs_data = json.loads(json_str)

        # Construction du fichier Word
        builder = AccessibilityDocBuilder()
        for p_elements in paragraphs_data:
          builder.add_paragraph_with_mixed_content(p_elements)

        output_filename = "Document_Anne_Accessible.docx"
        builder.save(output_filename)

        with open(output_filename, "rb") as f:
          st.download_button(
              label="📥 Télécharger le fichier Word (.docx) pour Anne",
              data=f,
              file_name=output_filename,
              mime=(
                  "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              ),
          )
        st.success("Document généré avec succès par Gemini !")

      except Exception as e:
        st.error(
            f"Une erreur est survenue lors du traitement avec Gemini : {e}"
        )
        st.text("Réponse brute de l'IA :")
        st.code(raw_result if "raw_result" in locals() else str(e))
elif uploaded_file is not None and not api_key:
  st.warning(
      "Veuillez entrer votre clé API Gemini ci-dessus pour lancer la"
      " conversion intelligente."
  )
