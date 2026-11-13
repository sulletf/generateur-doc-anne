import json
import time
import re
import docx
from docx.oxml import parse_xml
from docx.shared import Inches, Pt
from google import genai
from google.genai import types
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
        clean_text = text.replace(" = ", "=").replace(" =", "=").replace("= ", "=")

        if "\\sqrt" in clean_text:
            match = re.search(r'\\sqrt\{([^}]+)\}', clean_text)
            if match:
                radicand = match.group(1)
                xml_rad = f'<m:rad><m:deg/><m:e><m:r><m:t>{radicand}</m:t></m:r></m:e></m:rad>'
                self._insert_omml(paragraph, xml_rad)
                return

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
        p = self.doc.add_paragraph()
        for p_type, content in parts:
            if p_type == "text":
                p.add_run(content)
            elif p_type == "math":
                self.add_math(p, content)
        return p

    def save(self, filename):
        self.doc.save(filename)


# --- FONCTION D'INTERPRETATION PAR GEMINI ---
def parse_document_with_gemini(uploaded_file, file_bytes, mime_type, api_key):
    client = genai.Client(api_key=api_key)

    prompt = """
    Tu es un assistant spécialisé dans l'adaptation de documents pédagogiques pour Anne, une élève aveugle qui utilise une plage braille et une synthèse vocale.
    
    Règles absolues de mise en forme et d'accessibilité :
    1. AUCUNE mise en forme visuelle : pas de gras, pas d'italique, pas de souligné, pas de couleurs. Du texte brut standard.
    2. Nettoyage des doublons et artefacts : si le document contient plusieurs exemplaires identiques, n'en garde qu'un seul.
    3. Mathématiques : identifie tout ce qui relève des mathématiques (formules, équations). Tu peux utiliser la syntaxe LaTeX comme \\sqrt{x} ou x².
    4. Structure de sortie : Renvoie le résultat au format JSON strict (une liste de listes de tuples [type, contenu] avec le type "text" ou "math").
    
    Exemple de structure JSON attendue :
    [
       [ ["text", "Calculer "], ["math", "\\sqrt{25}"], ["text", "."] ]
    ]
    
    Renvoie uniquement le JSON valide, sans texte additionnel.
    """

    if mime_type == "text/plain":
        text_content = file_bytes.decode("utf-8", errors="ignore")
        contents_payload = [text_content, prompt]
    elif mime_type.startswith("image/"):
        image = Image.open(uploaded_file)
        contents_payload = [image, prompt]
    else:
        contents_payload = [
            types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
            prompt,
        ]

    # Utilisation rigoureuse et directe du modèle exigé par l'API
    response = client.models.generate_content(
        model="gemini-3.8-flash", contents=contents_payload
    )
    return response.text


# --- INTERFACE UTILISATEUR ---
uploaded_file = st.file_uploader(
    "Importez votre document de cours ou d'exercice (PDF, Image, TXT) :",
    type=["png", "jpg", "jpeg", "pdf", "txt"],
)

if uploaded_file is not None:
    if not api_key:
        st.error(
            "Erreur de configuration : La clé API Gemini n'est pas configurée dans les secrets Streamlit."
        )
    else:
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

        if st.button("Générer le document Word accessible pour Anne", type="primary"):
            with st.spinner("Gemini analyse le document..."):
                try:
                    raw_result = parse_document_with_gemini(
                        uploaded_file, file_bytes, mime_type, api_key
                    )

                    json_str = raw_result
                    if "```json" in raw_result:
                        json_str = raw_result.split("```json")[1].split("```")[0].strip()
                    elif "```" in raw_result:
                        json_str = raw_result.split("```")[1].split("```")[0].strip()

                    paragraphs_data = json.loads(json_str)

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
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                    st.success("Document généré avec succès !")

                except Exception as e:
                    st.error(f"Une erreur est survenue : {e}")
