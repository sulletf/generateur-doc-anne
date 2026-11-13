import json
import os
import docx
from docx.oxml import parse_xml
from docx.shared import Inches, Pt
from google import genai
from google.genai import types
import pdfplumber
from PIL import Image
import streamlit as st

# --- CONFIGURATION DE L'APPLICATION ---
st.set_page_config(
    page_title="Générateur Accessible Anne", page_icon="📄", layout="centered"
)

st.title("📄 Générateur de Documents Adaptés - Pour Anne")
st.markdown("""
Glissez-déposez le document de cours ou d'exercice (PDF, Image, TXT) ci-dessous. 
L'application le transforme automatiquement en un **document Word accessible** pour Anne.
""")

# Récupération automatique de la clé secrète configurée sur Streamlit Cloud
api_key = st.secrets.get("GEMINI_API_KEY", "")


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

    normal_style = self.doc.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(12)

  def _insert_omml(self, paragraph, xml_content):
    omml = f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{xml_content}</m:oMath>'
    paragraph._p.append(parse_xml(omml))

  def add_math(self, paragraph, text):
    """Convertit une expression mathématique en objet OMML propre, sans espaces autour de =."""
    clean_text = text.replace(" = ", "=").replace(" =", "=").replace("= ", "=")

    # Gestion des puissances au carré
    if "²" in clean_text or "^2" in clean_text:
      clean_text = clean_text.replace("^2", "²")
      parts = clean_text.split("²")
      base = parts[0]
      rest = parts[1] if len(parts) > 1 else ""
      xml_sup = f"<m:sSup><m:e><m:r><m:t>{base}</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>{rest}"
      self._insert_omml(paragraph, xml_sup)
    else:
      self._insert_omml(paragraph, f"<m:r><m:t>{clean_text}</m:t></m:r>")

  def add_paragraph_with_mixed_content(self, parts):
    """Ajoute un paragraphe contenant un mix de texte normal et de balises mathématiques."""
    p = self.doc.add_paragraph()
    for p_type, content in parts:
      if p_type == "text":
        p.add_run(content)
      elif p_type == "math":
        self.add_math(p, content)
    return p

  def save(self, filename):
    self.doc.save(filename)


# --- FONCTION D'INTERPRETATION PAR GEMINI (AVEC TOUTES LES RÈGLES D'ANNE) ---
def parse_document_with_gemini(uploaded_file, file_bytes, mime_type, api_key):
  client = genai.Client(api_key=api_key)

  prompt = """
    Tu es un assistant spécialisé dans l'adaptation de documents pédagogiques de mathématiques pour Anne, une élève aveugle qui utilise une plage braille et une synthèse vocale.
    
    Règles absolues de mise en forme et d'accessibilité :
    1. AUCUNE mise en forme visuelle : pas de gras, pas d'italique, pas de souligné, pas de couleurs. Du texte brut standard.
    2. Nettoyage des doublons et artefacts : si le document contient plusieurs exemplaires identiques de la même fiche ou des exercices en double sur la même page, n'en garde qu'un seul exemplaire propre et complet. Supprime les artefacts de texte corrompu ou d'impression.
    3. Mathématiques et géométrie : tu dois identifier avec une extrême précision tout ce qui relève des mathématiques (noms de triangles comme BOL, lettres de sommets, segments entre crochets comme [LO], formules, équations, valeurs numériques avec unités et calculs).
    4. Structure de sortie : Renvoie le résultat au format JSON strict (une liste de paragraphes, où chaque paragraphe est une liste de tuples [type, contenu] avec le type "text" ou "math").
    
    Exemple de structure JSON attendue :
    [
       [ ["text", "Le triangle "], ["math", "BOL"], ["text", " est rectangle en "], ["math", "L"], ["text", "."] ],
       [ ["math", "BO=37"], ["text", " et "], ["math", "BL=12"], ["text", "."] ]
    ]
    
    Analyse le document fourni en suivant scrupuleusement ces consignes et renvoie uniquement le JSON valide, sans texte additionnel autour.
    """

  if mime_type.startswith("image/"):
    image = Image.open(uploaded_file)
    response = client.models.generate_content(
        model="gemini-2.0-flash", contents=[image, prompt]
    )
  elif mime_type == "application/pdf":
    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=[
            types.Part.from_bytes(data=file_bytes, mime_type="application/pdf"),
            prompt,
        ],
    )
  else:
    text_content = file_bytes.decode("utf-8")
    response = client.models.generate_content(
        model="gemini-2.0-flash", contents=[text_content, prompt]
    )

  return response.text


# --- INTERFACE UTILISATEUR ---
uploaded_file = st.file_uploader(
    "Importez votre document de cours ou d'exercice (PDF, Image, TXT) :",
    type=["png", "jpg", "jpeg", "pdf", "txt"],
)

if uploaded_file is not None:
  if not api_key:
