import os
import json
import base64
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import resend
import docx
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import qn

STATE_FILE = "seen_jobs.json"
RECIPIENT_EMAIL = "fredericshearn@gmail.com"

# --- 1. GESTION DE L'ÉTAT (DÉDUPLICATION) ---

def load_seen_jobs():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            try:
                return set(json.load(f))
            except json.JSONDecodeError:
                return set()
    return set()

def save_seen_jobs(seen_set):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(sorted(list(seen_set)), f, ensure_ascii=False, indent=2)

# --- 2. EXTRACTION DES OFFRES (SCRAPING) ---

def fetch_vitijob_offers():
    url = "https://www.vitijob.com/emploi/region/1"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    offers = []
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "html.parser")
            for a_tag in soup.find_all("a", href=True):
                href = a_tag["href"]
                title = a_tag.get_text(strip=True)
                if "/emploi/" in href and len(title) > 12:
                    full_url = href if href.startswith("http") else f"https://www.vitijob.com{href}"
                    offers.append({"id": full_url, "title": title, "url": full_url, "source": "Vitijob"})
    except Exception as e:
        print(f"Erreur lors du scraping : {e}")
        
    return offers

# --- 3. INGÉNIERIE DOCUMENTAIRE WORD ---

def setup_multilevel_numbering(doc):
    numbering = doc.part.numbering_part.numbering_definitions._numbering
    abstract_num_xml = """
    <w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="100">
        <w:multiLevelType w:val="multilevel"/>
        <w:lvl w:ilvl="0">
            <w:start w:val="1"/>
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1."/>
            <w:lvlJc w:val="left"/>
            <w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
                <w:b/>
                <w:color w:val="0F2240"/>
            </w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="1">
            <w:start w:val="1"/>
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1.%2."/>
            <w:lvlJc w:val="left"/>
            <w:pPr><w:ind w:left="576" w:hanging="432"/></w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
                <w:b/>
                <w:color w:val="780000"/>
            </w:rPr>
        </w:lvl>
    </w:abstractNum>
    """
    num_xml = """
    <w:num xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:numId="100">
        <w:abstractNumId w:val="100"/>
    </w:num>
    """
    numbering.append(parse_xml(abstract_num_xml))
    numbering.append(parse_xml(num_xml))

def add_numbered_heading(doc, text, level):
    p = doc.add_paragraph(style=f"Heading {level}")
    pPr = p._p.get_or_add_pPr()
    numPr = parse_xml(f"""
        <w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:ilvl w:val="{level-1}"/>
            <w:numId w:val="100"/>
        </w:numPr>
    """)
    pPr.append(numPr)
    p.add_run(text)
    return p

def add_page_number(run):
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

def generate_docx(new_offers, filename):
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        
    setup_multilevel_numbering(doc)
    
    styles = doc.styles
    h1 = styles['Heading 1']
    h1.font.name = 'Calibri'
    h1.font.size = Pt(16)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor(15, 34, 64)
    
    h2 = styles['Heading 2']
    h2.font.name = 'Calibri'
    h2.font.size = Pt(13)
    h2.font.bold = True
    h2.font.color.rgb = RGBColor(120, 0, 0)

    # Pied de page
    footer = doc.sections[0].footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    f_run = footer.add_run("Veille Quotidienne Encadrement Viticole | Page ")
    f_run.font.size = Pt(9)
    f_run.font.color.rgb = RGBColor(120, 120, 120)
    add_page_number(footer.add_run())

    # Titre principal
    t_p = doc.add_paragraph()
    t_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_run = t_p.add_run("VEILLE QUOTIDIENNE — OFFRES D'ENCADREMENT VITICOLE")
    t_run.bold = True
    t_run.font.size = Pt(18)
    t_run.font.color.rgb = RGBColor(15, 34, 64)
    
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    s_run = sub.add_run(f"Bassin Vallée du Rhône — Rapport du {datetime.now().strftime('%d/%m/%Y')}")
    s_run.font.size = Pt(12)
    s_run.font.italic = True
    s_run.font.color.rgb = RGBColor(120, 0, 0)

    doc.add_paragraph()

    add_numbered_heading(doc, "Nouvelles Opportunités Identifiées", level=1)
    
    for item in new_offers:
        add_numbered_heading(doc, item["title"], level=2)
        p = doc.add_paragraph()
        p.add_run("• Source : ").bold = True
        p.add_run(f"{item['source']}\n")
        p.add_run("• Lien direct : ").bold = True
        p.add_run(item["url"])
        
    doc.save(filename)

# --- 4. EXPÉDITION VIA RESEND ---

def send_email_via_resend(file_path, count):
    api_key = os.environ.get("RESEND_API_KEY")
    if not api_key:
        raise ValueError("La clé API RESEND_API_KEY est manquante.")

    resend.api_key = api_key

    # Encodage du fichier Word en Base64 pour l'API Resend
    with open(file_path, "rb") as f:
        file_bytes = f.read()
        encoded_file = base64.b64encode(file_bytes).decode("utf-8")

    params = {
        "from": "Agent IA Veille <onboarding@resend.dev>",
        "to": [RECIPIENT_EMAIL],
        "subject": f"[Veille Viticole] {count} nouvelle(s) offre(s) — {datetime.now().strftime('%d/%m/%Y')}",
        "html": f"""
            <p>Bonjour Frédéric,</p>
            <p><strong>{count}</strong> nouvelle(s) offre(s) d'encadrement viticole ont été détectées aujourd'hui sur le bassin Vallée du Rhône.</p>
            <p>Le document Word récapitulatif est joint à ce message.</p>
            <br>
            <p><em>Agent IA de Veille Automatisée</em></p>
        """,
        "attachments": [
            {
                "filename": os.path.basename(file_path),
                "content": encoded_file
            }
        ]
    }

    response = resend.Emails.send(params)
    print(f"E-mail envoyé via Resend. ID : {response.get('id')}")

# --- 5. EXÉCUTION ---

def main():
    seen_jobs = load_seen_jobs()
    current_offers = fetch_vitijob_offers()
    
    new_offers = [job for job in current_offers if job["id"] not in seen_jobs]
    
    if not new_offers:
        print("Aucune nouvelle offre détectée aujourd'hui.")
        return

    print(f"{len(new_offers)} nouvelle(s) offre(s) trouvée(s). Génération du livrable...")
    
    filename = f"Offres_Viticoles_{datetime.now().strftime('%Y%m%d')}.docx"
    generate_docx(new_offers, filename)
    
    send_email_via_resend(filename, len(new_offers))
    
    for job in new_offers:
        seen_jobs.add(job["id"])
    save_seen_jobs(seen_jobs)
    
    print("État mis à jour avec succès.")

if __name__ == "__main__":
    main()