import os
import json
import re
import base64
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import resend
import docx
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import qn, nsdecls

STATE_FILE = "seen_jobs.json"
RECIPIENT_EMAIL = "fredericshearn@gmail.com"

# --- 1. FILTRES D'ENCADREMENT & MOTS-CLÉS ---

EXCLUDE_KEYWORDS = [
    "ouvrier", "ouvrière", "ménage", "saisonnier", "vendangeur", "vendangeuse",
    "stagiaire", "stage", "apprentissage", "alternance", "manœuvre", "tractoriste"
]

INCLUDE_KEYWORDS = [
    "directeur", "directrice", "direction", "responsable", "manager", "head of",
    "commercial", "commerciale", "export", "adv", "daf", "raf", "gestion",
    "chef de secteur", "compte", "encadrement", "ingénieur", "chef de cave"
]

def is_management_offer(title):
    title_lower = title.lower()
    if any(ex in title_lower for ex in EXCLUDE_KEYWORDS):
        return False
    return any(inc in title_lower for inc in INCLUDE_KEYWORDS)

# --- 2. GESTION DE L'ÉTAT ---

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

# --- 3. SCRAPING AVANCÉ DES OFFRES ---

def fetch_job_details(url, headers):
    """Scrape le détail d'une offre d'emploi ciblée."""
    details = {"structure": "Non spécifiée", "location": "Vallée du Rhône", "missions": "Consulter l'annonce pour le détail des missions."}
    try:
        res = requests.get(url, headers=headers, timeout=10)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, "html.parser")
            
            # Recherche de la structure et localisation
            info_block = soup.find("div", class_="info-job") or soup.find("section", class_="content")
            if info_block:
                text = info_block.get_text(" ", strip=True)
                lines = [l.strip() for l in text.split("\n") if len(l.strip()) > 2]
                if len(lines) > 0:
                    details["structure"] = lines[0]
            
            # Extraction des détails de mission
            desc_tag = soup.find("div", class_="description") or soup.find("article")
            if desc_tag:
                clean_desc = desc_tag.get_text(" ", strip=True)
                if len(clean_desc) > 50:
                    details["missions"] = clean_desc[:250] + "..."
    except Exception:
        pass
    return details

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
                
                # Filtrage strict : URL d'annonce réelle uniquement (/emploi/DIGITS/...)
                if re.search(r'/emploi/\d+/', href) and len(title) > 8:
                    if is_management_offer(title):
                        full_url = href if href.startswith("http") else f"https://www.vitijob.com{href}"
                        offers.append({
                            "id": full_url,
                            "title": title,
                            "url": full_url,
                            "source": "Vitijob"
                        })
    except Exception as e:
        print(f"Erreur lors du scraping : {e}")
        
    # Enrichissement des annonces filtrées
    final_offers = []
    for job in offers[:10]: # Limite de sécurité
        details = fetch_job_details(job["url"], headers)
        job.update(details)
        final_offers.append(job)
        
    return final_offers

# --- 4. INGÉNIERIE DOCUMENTAIRE WORD ---

def setup_multilevel_numbering(doc):
    numbering = doc.part.numbering_part.numbering_definitions._numbering
    abstract_num_xml = """
    <w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="100">
        <w:multiLevelType w:val="multilevel"/>
        <w:lvl w:ilvl="0">
            <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>
            <w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr>
            <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="0F2240"/></w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="1">
            <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/>
            <w:pPr><w:ind w:left="576" w:hanging="432"/></w:pPr>
            <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="780000"/></w:rPr>
        </w:lvl>
    </w:abstractNum>
    """
    num_xml = """<w:num xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:numId="100"><w:abstractNumId w:val="100"/></w:num>"""
    numbering.append(parse_xml(abstract_num_xml))
    numbering.append(parse_xml(num_xml))

def add_numbered_heading(doc, text, level):
    p = doc.add_paragraph(style=f"Heading {level}")
    pPr = p._p.get_or_add_pPr()
    numPr = parse_xml(f'<w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:ilvl w:val="{level-1}"/><w:numId w:val="100"/></w:numPr>')
    pPr.append(numPr)
    p.add_run(text)
    return p

def set_cell_background(cell, fill_hex):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def generate_docx(new_offers, filename):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.5)
        
    setup_multilevel_numbering(doc)
    
    styles = doc.styles
    styles['Heading 1'].font.name = 'Calibri'
    styles['Heading 1'].font.size = Pt(16)
    styles['Heading 1'].font.bold = True
    styles['Heading 1'].font.color.rgb = RGBColor(15, 34, 64)
    
    styles['Heading 2'].font.name = 'Calibri'
    styles['Heading 2'].font.size = Pt(13)
    styles['Heading 2'].font.bold = True
    styles['Heading 2'].font.color.rgb = RGBColor(120, 0, 0)

    # Titre
    t_p = doc.add_paragraph()
    t_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_run = t_p.add_run("SYNTHÈSE DES OPPORTUNITÉS D'ENCADREMENT & DIRECTION")
    t_run.bold = True
    t_run.font.size = Pt(18)
    t_run.font.color.rgb = RGBColor(15, 34, 64)
    
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    s_run = sub.add_run(f"Secteur Viticole & Négoce — Bassin Vallée du Rhône\nRapport Automatisé du {datetime.now().strftime('%d/%m/%Y')}")
    s_run.font.size = Pt(12)
    s_run.font.italic = True
    s_run.font.color.rgb = RGBColor(120, 0, 0)

    doc.add_paragraph()

    # Section 1 : Matrice récapitulative
    add_numbered_heading(doc, "Matrice Récapitulative des Offres Identifiées", level=1)
    
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    headers = ["Intitulé du Poste", "Structure", "Localisation", "Source"]
    widths = [Cm(6.0), Cm(4.0), Cm(3.5), Cm(2.5)]
    
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "0F2240")
        p = hdr_cells[i].paragraphs[0]
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9.5)

    for row_idx, item in enumerate(new_offers):
        row_cells = table.add_row().cells
        bg_color = "F7FAFC" if row_idx % 2 == 1 else "FFFFFF"
        data = [item["title"], item.get("structure", "N/C"), item.get("location", "Vallée du Rhône"), item["source"]]
        for i, val in enumerate(data):
            row_cells[i].text = val
            set_cell_background(row_cells[i], bg_color)
            p = row_cells[i].paragraphs[0]
            for run in p.runs:
                run.font.size = Pt(9)

    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

    doc.add_paragraph()

    # Section 2 : Fiches détaillées
    add_numbered_heading(doc, "Détail des Postes d'Encadrement", level=1)
    
    for item in new_offers:
        add_numbered_heading(doc, item["title"], level=2)
        p = doc.add_paragraph()
        p.add_run("• Structure : ").bold = True
        p.add_run(f"{item.get('structure', 'N/C')}\n")
        p.add_run("• Missions & Profil : ").bold = True
        p.add_run(f"{item.get('missions', 'Consulter l\'annonce.')}\n")
        p.add_run("• Lien direct : ").bold = True
        p.add_run(item["url"])

    doc.save(filename)

# --- 5. EXPÉDITION VIA RESEND ---

def send_email_via_resend(file_path, count):
    api_key = os.environ.get("RESEND_API_KEY")
    if not api_key:
        raise ValueError("Clé RESEND_API_KEY manquante.")

    resend.api_key = api_key

    with open(file_path, "rb") as f:
        encoded_file = base64.b64encode(f.read()).decode("utf-8")

    params = {
        "from": "Agent IA Veille <onboarding@resend.dev>",
        "to": [RECIPIENT_EMAIL],
        "subject": f"[Veille Viticole Cadres] {count} nouvelle(s) offre(s) qualifiée(s) — {datetime.now().strftime('%d/%m/%Y')}",
        "html": f"""
            <p>Bonjour Frédéric,</p>
            <p><strong>{count}</strong> nouvelle(s) offre(s) d'encadrement/direction viticole qualifiée(s) ont été identifiées aujourd'hui sur le bassin Vallée du Rhône.</p>
            <p>Le rapport Word structuré mis à jour est disponible en pièce jointe.</p>
            <br>
            <p><em>Agent IA de Veille Automatisée</em></p>
        """,
        "attachments": [{"filename": os.path.basename(file_path), "content": encoded_file}]
    }
    resend.Emails.send(params)

# --- 6. EXÉCUTION ---

def main():
    seen_jobs = load_seen_jobs()
    current_offers = fetch_vitijob_offers()
    
    new_offers = [job for job in current_offers if job["id"] not in seen_jobs]
    
    if not new_offers:
        print("Aucune nouvelle offre d'encadrement qualifiée aujourd'hui.")
        return

    print(f"{len(new_offers)} offre(s) qualifiée(s) trouvée(s). Génération du rapport Word...")
    filename = f"Offres_Viticoles_{datetime.now().strftime('%Y%m%d')}.docx"
    generate_docx(new_offers, filename)
    send_email_via_resend(filename, len(new_offers))
    
    for job in new_offers:
        seen_jobs.add(job["id"])
    save_seen_jobs(seen_jobs)
    print("Envoi et mise à jour de l'état terminés.")

if __name__ == "__main__":
    main()