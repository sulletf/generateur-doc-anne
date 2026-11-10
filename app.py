import docx
from docx.oxml import parse_xml
from docx.shared import Inches, Pt
import streamlit as st

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


# --- INTERFACE STREAMLIT ---
st.set_page_config(
    page_title="Générateur Accessible Word",
    page_icon="📄",
    layout="centered",
)

st.title("📄 Générateur de documents adaptatifs - Pour Anne")
st.markdown("""
Cette application transforme du texte d'exercices en **fichier Word (.docx) structuré et accessible** pour la plage tactile / synthèse vocale d'Anne.

* **Formatage épuré :** Pas de couleur, pas de gras, pas d'italique.
* **Mathématiques propres :** Objets OMML sans espaces superflues autour des `=` et avec espaces insécables avant les unités.
""")

st.subheader("Saisir un exercice")
titre_exo = st.text_input("Titre de l'exercice :", "Exercice 1")

syntaxe_guide = st.expander("💡 Guide de saisie des formules mathématiques")
syntaxe_guide.write("""
Pour séparer le texte normal des formules mathématiques :
* Encadrez les expressions mathématiques ou grandeurs avec le tag `[M]` et `[/M]`.
* Exemples :
  * `Le triangle [M]PIE[/M] rectangle en [M]I[/M] est tel que [M]IP=7 cm[/M].`
  * `Une longueur de [M]12 m[/M] et une masse de [M]30 kg[/M].`
""")

texte_brut = st.text_area(
    "Texte des paragraphes :",
    height=200,
    placeholder=(
        "Entrez votre texte ici. Ex:\nSur le schéma, le segment [M][DN][/M]"
        " mesure [M]4 m[/M] de longueur."
    ),
)

if st.button("Générer le document Word", type="primary"):
  if texte_brut.strip():
    builder = AccessibilityDocBuilder()
    builder.add_heading(titre_exo)

    # Découpage ligne par ligne et parsing du tag [M]...[/M]
    lignes = texte_brut.split('\n')
    for ligne in lignes:
      if not ligne.strip():
        continue
      elements = []
      parts = ligne.split('[M]')
      for part in parts:
        if '[/M]' in part:
          math_content, rest_text = part.split('[/M]', 1)
          elements.append(('math', math_content))
          if rest_text:
            elements.append(('text', rest_text))
        else:
          if part:
            elements.append(('text', part))

      builder.add_paragraph_with_mixed_content(elements)

    output_filename = "Exercice_Anne_Accessible.docx"
    builder.save(output_filename)

    with open(output_filename, 'rb') as f:
      st.download_button(
          label="📥 Télécharger le fichier .docx pour Anne",
          data=f,
          file_name=output_filename,
          mime=(
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
          ),
      )
    st.success("Document généré avec succès !")
  else:
    st.error("Veuillez saisir du texte avant de lancer la génération.")