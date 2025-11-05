import streamlit as st
from io import BytesIO
import datetime
import uuid
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,
    Image, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PIL import Image as PILImage
from supabase import create_client, Client

# ============================
# CONFIG
# ============================
st.set_page_config(page_title="Cotation & Contrat - Caution Leadway", page_icon="briefcase", layout="wide")
LOGO_PATH = "leadway logo all formats big-02.png"
SIGNATURE_PATH = "signature.png"
BAS_DE_PAGE_PATH = "bas_de_page.png"

# Locale française
def set_french_locale():
    try:
        import locale
        locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')
    except:
        try:
            locale.setlocale(locale.LC_TIME, 'French_France.1252')
        except:
            st.warning("Locale française indisponible – dates en anglais.")
set_french_locale()

# ============================
# SUPABASE CONNECTION
# ============================
@st.cache_resource
def init_supabase_client():
    url = st.secrets["https://xabklcnwvlteldfukwgh.supabase.co"]
    key = st.secrets["eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhhYmtsY253dmx0ZWxkZnVrd2doIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjIzMzg2NjYsImV4cCI6MjA3NzkxNDY2Nn0.anB7Dhj8ucPIG1arHp_LJCE7Eq6g0-kR9CmhiFOdESU"]
    return create_client(url, key)

supabase: Client = init_supabase_client()

# ============================
# UTILITAIRES
# ============================
def fmt_money(val):
    try:
        val = float(val)
        return f"{int(round(val)):,}".replace(",", " ") + " F CFA"
    except:
        return "0 F CFA"

def format_date_fr(date_obj):
    mois = ["janvier","février","mars","avril","mai","juin",
            "juillet","août","septembre","octobre","novembre","décembre"]
    return f"{date_obj.day} {mois[date_obj.month-1]} {date_obj.year}"

def number_to_words(num):
    if num == 0:
        return "zéro"
    units = ["","un","deux","trois","quatre","cinq","six","sept","huit","neuf"]
    teens = ["dix","onze","douze","treize","quatorze","quinze","seize",
             "dix-sept","dix-huit","dix-neuf"]
    tens = ["","","vingt","trente","quarante","cinquante","soixante",
            "soixante-dix","quatre-vingt","quatre-vingt-dix"]
    thousands = ["","mille","million","milliard"]

    def below_1000(n):
        if n < 10:
            return units[n]
        if n < 20:
            return teens[n-10]
        if n < 100:
            t = n // 10
            u = n % 10
            if u == 0:
                return tens[t]
            if t in (7,9):
                return tens[t-1] + "-et-" + units[u] if u == 1 else tens[t-1] + "-" + units[u]
            return tens[t] + ("-" if u else "") + units[u]
        h = n // 100
        r = n % 100
        if r == 0:
            return units[h] + " cent"
        return units[h] + " cent " + below_1000(r)

    res, i = [], 0
    while num:
        chunk = num % 1000
        if chunk:
            word = below_1000(chunk)
            if i:
                # "mille" ne prend jamais de "s"
                # "million" et "milliard" s'accordent au pluriel
                if i == 1:  # mille - pas de pluriel
                    word += " " + thousands[i]
                elif i >= 2:  # million ou milliard - avec pluriel
                    word += " " + thousands[i] + ("s" if chunk > 1 else "")
            res.append(word)
        num //= 1000
        i += 1
    return " ".join(reversed(res)).capitalize()

# ============================
# SUPABASE FUNCTIONS (NEW SCHEMA)
# ============================

def save_cotation_to_supabase(data, lots_data, detail_agrement):
    """
    Étape 1 : Enregistre la cotation et ses lots.
    Retourne l'ID de la nouvelle cotation.
    """
    try:
        # 1. Préparer l'enregistrement de la cotation
        cotation_record = {
            "assure": data.get("assure"),
            "souscripteur": data.get("souscripteur"),
            "beneficiaire": data.get("beneficiaire"),
            "adresse_beneficiaire": data.get("adresse_beneficiaire"),
            "adresse": data.get("adresse"),
            "situation_geo": data.get("situation_geo"),
            "num_marche": data.get("num_marche"),
            "autorite": data.get("autorite"),
            "date_depot": data.get("date_depot"),
            "objet": data.get("objet"),
            "couverture": data.get("couverture"),
            "detail_agrement": detail_agrement if data.get("couverture") == "Caution d'agrément" else None,
            "montant_marche": data.get("montant_marche"),
            "duree": data.get("duree"),
            "montant_caution": data.get("montant_caution"),
            "prime_nette": data.get("prime_nette"),
            "frais_analyse": data.get("frais_analyse"),
            "accessoires": data.get("accessoires"),
            "taxes": data.get("taxes"),
            "prime_ttc": data.get("prime_ttc"),
            "date_cotation": data.get("date_cotation"),
            "suretes_text": data.get("suretes_text"),
            "statut": "Générée"
        }
        
        # Insérer la cotation et récupérer son ID
        response_cotation = supabase.table('cotations').insert(cotation_record).execute()
        
        if response_cotation.data:
            new_cotation_id = response_cotation.data[0]['id']
            
            # 2. Insérer les lots (s'il y en a)
            if lots_data:
                lots_to_insert = []
                for lot in lots_data:
                    lots_to_insert.append({
                        "cotation_id": new_cotation_id, # Clé de liaison
                        "lot_num": lot.get("Lot"),
                        "montant": lot.get("Montant"),
                        "designation": lot.get("Désignation")
                    })
                
                if lots_to_insert:
                    response_lots = supabase.table('lots').insert(lots_to_insert).execute()
                    if not response_lots.data:
                        raise Exception("Erreur lors de l'insertion des lots.")
            
            return new_cotation_id, "Cotation et lots enregistrés."
        else:
            raise Exception("Erreur lors de l'insertion de la cotation.")
            
    except Exception as e:
        st.error(f"Erreur Supabase (save_cotation): {e}")
        return None, str(e)

def save_police_to_supabase(cotation_db_id, contrat_data):
    """
    Étape 2 : Crée la police liée à la cotation et met à jour le statut de la cotation.
    """
    try:
        # 1. Préparer l'enregistrement de la police
        police_record = {
            "cotation_id": cotation_db_id,
            "police_num": contrat_data.get("police_num"),
            "date_emission": contrat_data.get("date_emission"),
            "date_effet": contrat_data.get("date_effet"),
            "date_echeance": contrat_data.get("date_echeance"),
            "duree_police": contrat_data.get("duree_police")
        }
        
        # Insérer la police
        response_police = supabase.table('polices').insert(police_record).execute()
        
        if not response_police.data:
            raise Exception("Erreur lors de l'insertion de la police.")

        # 2. Mettre à jour le statut de la cotation
        response_update = supabase.table('cotations') \
                                  .update({"statut": "Contractualisée"}) \
                                  .eq('id', cotation_db_id) \
                                  .execute()
        
        if not response_update.data:
             raise Exception("Erreur lors de la mise à jour du statut de la cotation.")

        return True, "Police enregistrée et cotation mise à jour."
        
    except Exception as e:
        # Gérer les doublons de police_num ou cotation_id (contrainte UNIQUE)
        if "duplicate key" in str(e):
            st.error(f"Erreur Supabase (save_police): Une police existe déjà pour cette cotation ou ce numéro de police.")
        else:
            st.error(f"Erreur Supabase (save_police): {e}")
        return False, str(e)


# ============================
# PDF COTATION
# ============================
def generate_caution_pdf(data, lots_data):
    buffer = BytesIO()
    
    # Fonction pour ajouter le bas de page
    def add_footer(canvas, doc):
        canvas.saveState()
        try:
            # Bas de page
            if os.path.exists(BAS_DE_PAGE_PATH):
                pil_img = PILImage.open(BAS_DE_PAGE_PATH)
                w, h = pil_img.size
                ratio = h / w
                footer_width = A4[0]
                footer_height = footer_width * ratio
                canvas.drawImage(BAS_DE_PAGE_PATH, 0, 0, 
                               width=footer_width, height=footer_height,
                               preserveAspectRatio=True, mask='auto')
        except Exception as e:
            pass
        canvas.restoreState()
    
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=20, leftMargin=20,
                            topMargin=20, bottomMargin=60)
    styles = getSampleStyleSheet()
    elements = []

    style_normal = ParagraphStyle('Normal', parent=styles['Normal'], fontSize=9, leading=12, alignment=4)
    style_bold = ParagraphStyle('Bold', parent=styles['Normal'], fontSize=9,
                                leading=12, fontName="Helvetica-Bold", alignment=4)

    # Logo
    try:
        pil_img = PILImage.open(LOGO_PATH)
        w, h = pil_img.size
        ratio = h / w
        target_w = 40 * mm
        target_h = target_w * ratio
        logo = Image(LOGO_PATH, width=target_w, height=target_h)
        logo.hAlign = 'RIGHT'
        elements.append(logo)
        elements.append(Spacer(1, 6))
    except Exception as e:
        st.warning(f"Logo introuvable : {e}")

    # Bandeau titre
    titre = data["couverture"].upper()
    bandeau = Table(
        [[Paragraph(f"<b>OFFRE D'ASSURANCE CAUTION DE {titre}</b>",
                    ParagraphStyle('Bandeau', textColor=colors.white,
                                   alignment=1, fontSize=12, leading=14))]],
        colWidths=[A4[0] - 40]
    )
    bandeau.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.black),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
    ]))
    elements.append(bandeau)
    elements.append(Spacer(1, 6))

    # Texte intro
    elements.append(Paragraph(
        f"Comme suite à votre demande de cotation du {data['date_cotation']}, "
        "nous vous présentons ci-dessous les conditions de garanties et de primes "
        "pour la couverture Caution sollicitée.",
        style_normal))
    elements.append(Spacer(1, 10))

    # Tableau infos
    table_data = [
        [Paragraph("<b>Assuré</b>", style_normal), data.get("assure", "N/A")],
        [Paragraph("<b>Adresse</b>", style_normal), data.get("adresse", "N/A")],
        [Paragraph("<b>Situation géographique du marché</b>", style_normal), data.get("situation_geo", "N/A")],
        [Paragraph("<b>Numéro du marché</b>", style_normal), data.get("num_marche", "N/A")],
        [Paragraph("<b>Autorité contractante</b>", style_normal), data.get("autorite", "N/A")],
        [Paragraph("<b>Date de dépôt du dossier</b>", style_normal),
         "Selon contrat" if data.get("couverture") != "Soumission" else data.get("date_depot", "N/A")],
        [Paragraph("<b>Objet du marché</b>", style_normal), data.get("objet", "N/A")],
        [Paragraph("<b>Couverture</b>", style_normal), data.get("couverture", "N/A")],
        [Paragraph("<b>Montant du marché</b>", style_normal), fmt_money(data.get('montant_marche', 0))],
        [Paragraph("<b>Durée de la garantie</b>", style_normal), str(data.get("duree", "N/A"))],
        [Paragraph("<b>Montant à cautionner</b>", style_normal), fmt_money(data.get('montant_caution', 0))],
        [Paragraph("<b>Limites & Franchises</b>", style_normal), "Néant"],
    ]
    table_infos = Table(table_data, colWidths=[260, 280])
    table_infos.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.6, colors.black),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 5),
    ]))
    elements.append(table_infos)
    elements.append(Spacer(1, 12))

    # DÉCOMPTE DE PRIME
    elements.append(Paragraph("<b>DÉCOMPTE DE PRIME :</b>", style_bold))

    style_cell = ParagraphStyle(name='Cell', fontName='Helvetica', fontSize=10,
                                alignment=1, leading=12, spaceAfter=0, spaceBefore=0)
    style_bold_cell = ParagraphStyle(name='BoldCell', parent=style_cell, fontName='Helvetica-Bold')

    prime_data = [
        [Paragraph("Prime HT", style_bold_cell), Paragraph("Acc.", style_bold_cell),
         Paragraph("Frais d'analyse", style_bold_cell), Paragraph("Taxe", style_bold_cell),
         Paragraph("Prime TTC", style_bold_cell)],
        [Paragraph(fmt_money(data['prime_nette']), style_cell),
         Paragraph(fmt_money(data['accessoires']), style_cell),
         Paragraph(fmt_money(data['frais_analyse']), style_cell),
         Paragraph(fmt_money(data['taxes']), style_cell),
         Paragraph(f"<b>{fmt_money(data['prime_ttc'])}</b>", style_cell)]
    ]

    prime_table = Table(prime_data, colWidths=[100, 80, 100, 80, 120])
    prime_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    elements.append(prime_table)
    elements.append(Spacer(1, 12))

    # Bandeau gris
    def header_band(title):
        p = Paragraph(f"<b>{title}</b>",
                      ParagraphStyle("hb", textColor=colors.white,
                                     fontSize=10, alignment=0, leftIndent=5))
        band = Table([[p]], colWidths=[A4[0]-40])
        band.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#6e6e6e")),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING", (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ]))
        return band

    elements.append(header_band("Offre soumise sous réserve de nous transmettre :"))
    elements.append(Paragraph("""
        - Modèle de l'acte de caution<br/>
        - Attestations de bonne exécution des marchés similaires déjà réalisés<br/>
        - Documents Administratifs (RCCM - CNI DU GÉRANT - STATUTS - DFE)<br/>
        - Documents Financiers (États financiers des 3 dernières années ou relevé bancaire sur une année)<br/>
        - Contrat de marché signé
    """, style_normal))
    elements.append(Spacer(1, 8))

    if data.get("suretes_text"):
        elements.append(header_band("Sûretés et mesures cumulatives :"))
        suretes = data["suretes_text"].replace('\n', '<br/>')
        elements.append(Paragraph(f'<font color="red">{suretes}</font>', style_normal))
        elements.append(Spacer(1, 8))

    elements.append(header_band("Exclusions :"))
    elements.append(Paragraph("""
        - Dommages et pertes découlant directement ou indirectement des épidémies/pandémies ;<br/>
        - Risque et violence politique, guerre civile ou étrangère
    """, style_normal))
    elements.append(Spacer(1, 12))

    signature_date = format_date_fr(datetime.date.today())
    elements.append(Paragraph(
        f"<para alignment='right'>Fait à Abidjan, le {signature_date}<br/><br/>"
        "<b>POUR LA COMPAGNIE</b><br/>Leadway Assurance IARD</para>",
        style_normal))

    if lots_data:
        elements.append(PageBreak())
        elements.append(Paragraph("<b>DÉTAILS DES LOTS</b>",
                                  ParagraphStyle('Titre', fontSize=12, leading=14, spaceAfter=10)))
        lots_table_data = [[
            Paragraph("<b>Numéro du lot</b>", style_normal),
            Paragraph("<b>Montant à cautionner</b>", style_normal),
            Paragraph("<b>Désignation</b>", style_normal)
        ]]
        for lot in lots_data:
            lots_table_data.append([
                lot.get("Lot", ""),
                fmt_money(lot.get("Montant", 0)),
                lot.get("Désignation", "")
            ])
        lots_table = Table(lots_table_data, colWidths=[100, 130, 310])
        lots_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('VALIGN', (0,0), (-1,-1), "MIDDLE"),
        ]))
        elements.append(lots_table)

    doc.build(elements, onFirstPage=add_footer, onLaterPages=add_footer)
    buffer.seek(0)
    return buffer

# ============================
# PDF CONTRAT CAUTION D'AGRÉMENT
# ============================
def generate_contrat_agrement_pdf(data, detail_agrement, lots_data=None):
    buffer = BytesIO()
    
    # Fonction pour ajouter le bas de page
    def add_footer(canvas, doc):
        canvas.saveState()
        try:
            if os.path.exists(BAS_DE_PAGE_PATH):
                pil_img = PILImage.open(BAS_DE_PAGE_PATH)
                w, h = pil_img.size
                ratio = h / w
                footer_width = A4[0]
                footer_height = footer_width * ratio
                canvas.drawImage(BAS_DE_PAGE_PATH, 0, 0, 
                               width=footer_width, height=footer_height,
                               preserveAspectRatio=True, mask='auto')
        except Exception as e:
            pass
        canvas.restoreState()
    
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=40, leftMargin=40,
                            topMargin=30, bottomMargin=70)
    styles = getSampleStyleSheet()
    elements = []

    style_title = ParagraphStyle('TitleCenter', parent=styles['Title'],
                                 alignment=1, fontSize=14, spaceAfter=20,
                                 fontName="Helvetica-Bold")
    style_center = ParagraphStyle('Center', alignment=1, fontSize=11, spaceAfter=15)
    style_bold = ParagraphStyle('Bold', fontName='Helvetica-Bold',
                                fontSize=10, leading=12, spaceAfter=8, alignment=4)
    style_normal = ParagraphStyle('Normal', fontSize=10, leading=12, alignment=4)
    style_underline = ParagraphStyle('Underline', fontName='Helvetica-Bold',
                                     fontSize=10, leading=12, alignment=1)

    # PAGE 1 - Page de garde
    try:
        pil_img = PILImage.open(LOGO_PATH)
        w, h = pil_img.size
        ratio = h / w
        target_w = 80 * mm
        target_h = target_w * ratio
        logo = Image(LOGO_PATH, width=target_w, height=target_h)
        logo.hAlign = 'CENTER'
        elements.append(logo)
        elements.append(Spacer(1, 40))
    except Exception as e:
        st.warning(f"Logo : {e}")

    elements.append(Paragraph(f"{data['assure']}", style_title))
    elements.append(Spacer(1, 30))
    elements.append(Paragraph("CONDITIONS PARTICULIERES", style_title))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(f"ASSURANCE CAUTION D'AGREMENT/ {detail_agrement}", style_title))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(f"POLICE NUMERO No {data['police_num']}", style_title))
    
    # PAGE 2 - Conditions particulières avec détail
    elements.append(PageBreak())
    
    elements.append(Paragraph(f"<u>CONDITIONS PARTICULIÈRES – {detail_agrement}</u>", style_underline))
    elements.append(Spacer(1, 20))
    
    # Tableau d'informations
    info_data = [
        ["SOUSCRIPTEUR", f": {data['souscripteur']}"],
        ["ASSURE", f": {data['assure']}"],
        ["INTERMEDIAIRE", ": OLEA AFRICA"],
        ["CODE", ": 2003"],
        ["DATE D'EMISSION", f": {data['date_emission']}"],
        ["DATE D'EFFET", f": {data['date_effet']}"],
        ["DATE D'ECHEANCE", f": {data['date_echeance']}"],
        ["A DUREE FERME", ""],
    ]
    info_table = Table(info_data, colWidths=[150, 380])
    info_table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 20))
    
    # Décompte de prime
    elements.append(Paragraph("<b>DECOMPTE DE PRIME & CONTRE-GARANTIE :</b>", style_bold))
    elements.append(Spacer(1, 10))
    
    def fmt_money_no_currency(val):
        try:
            val = float(val)
            return f"{int(round(val)):,}".replace(",", " ")
        except:
            return "0"
    
    elements.append(Paragraph("<b>Détail prime</b>", style_normal))
    prime_detail_data = [
        ["Prime nette", "Frais d'analyse", "Accessoires", "Taxes", "Prime TTC"],
        [fmt_money_no_currency(data['prime_nette']),
         fmt_money_no_currency(data['frais_analyse']),
         fmt_money_no_currency(data['accessoires']),
         fmt_money_no_currency(data['taxes']),
         fmt_money_no_currency(data['prime_ttc'])]
    ]
    prime_detail_table = Table(prime_detail_data, colWidths=[100, 100, 100, 100, 100])
    prime_detail_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
    ]))
    elements.append(prime_detail_table)
    elements.append(Spacer(1, 15))
    
    # Contre-garantie
    montant_contre_garantie = data.get('montant_caution', 0) * 2  # Exemple: 2x le montant
    elements.append(Paragraph(f"<b><i>Contre-garantie à déposer</i></b>        {fmt_money(montant_contre_garantie)}", style_normal))
    elements.append(Spacer(1, 15))
    
    # Texte de constitution
    elements.append(Paragraph("""
    La présente police est constituée par :<br/>
    Des Conditions Générales et des présentes Conditions Particulières dont l'assuré reconnaît avoir reçu un exemplaire.<br/>
    Les conditions particulières annulent et remplacent toutes dispositions des Conditions Générales qui seraient plus restrictives 
    que celles des conditions particulières ou qui présenteraient par rapport à celles-ci une divergence ou une incompatibilité.
    """, style_normal))
    elements.append(Spacer(1, 30))
    
    # Signatures
    sig_img = Image(SIGNATURE_PATH, width=170, height=140) if os.path.exists(SIGNATURE_PATH) else ""
    sig_table = Table([
        [Paragraph("<b>LE SOUSCRIPTEUR</b>", style_normal),
         Paragraph("<b>POUR L'ASSUREUR</b>", style_normal)],
        ["", ""],
        ["", sig_img],
    ], colWidths=[280, 220])
    sig_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    elements.append(sig_table)

    # PAGE 3 - Conditions particulières - Caution professionnelle
    elements.append(PageBreak())
    
    elements.append(Paragraph("<u>CONDITIONS PARTICULIÈRES – CAUTION PROFESSIONNELLE</u>", style_underline))
    elements.append(Spacer(1, 20))
    
    elements.append(Paragraph(f"<b>Police n° : {data['police_num']}</b>", style_bold))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph(f"<b>Souscripteurs / Donneurs d'ordre :</b>", style_bold))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph(f"<b>{data['assure']}</b><br/><i>{data.get('adresse', '01 BP 12792 Abidjan 01')}</i>", style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph(f"<b>Assureur :</b>", style_bold))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph("""
    <b>LEADWAY ASSURANCE IARD, 01 BP 11944 Abidjan 01</b> Société Anonyme au capital de 5 000 000 000 FCFA, 
    dont le siège est à Abidjan, Cocody 7ème Tranche, représenté par Monsieur Tiornan COULIBALY, Son Directeur Général.
    """, style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph(f"<b>Bénéficiaire :</b>", style_bold))
    elements.append(Spacer(1, 5))
    beneficiaire_text = f"<b>{data.get('beneficiaire', data.get('autorite', 'LA DIRECTION DES DOUANES'))}</b>"
    if data.get('adresse_beneficiaire') and data.get('adresse_beneficiaire') != "N/A":
        beneficiaire_text += f"<br/><i>{data.get('adresse_beneficiaire')}</i>"
    elements.append(Paragraph(beneficiaire_text, style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph(f"<b>Identification du marché :</b>", style_bold))
    elements.append(Spacer(1, 5))
    objet_default = "Couvrir le matériel importé depuis l'Afrique du Sud pour réaliser une étude à court terme en Côte d'Ivoire."
    elements.append(Paragraph(data.get('objet', objet_default), style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph(f"<b>Montant cautionné :</b>", style_bold))
    elements.append(Spacer(1, 5))
    montant_lettres = number_to_words(int(data['montant_caution']))
    elements.append(Paragraph(f"{montant_lettres.capitalize()} ({fmt_money(data['montant_caution'])}) francs CFA.", style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph(f"<b>Durée de validité :</b>", style_bold))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph(f"{data.get('duree', '60 jours')} à compter du {data['date_effet']} au {data['date_echeance']}", style_normal))
    elements.append(Spacer(1, 15))
    
    # Article 1
    elements.append(Paragraph("<u><b>Article 1 – Objet de la Garantie</b></u>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("""
    L'assureur se porte caution solidaire et principal débiteur du Souscripteur auprès du bénéficiaire, 
    pour Caution en Douane.
    """, style_normal))
    elements.append(Spacer(1, 10))
    
    # Article 2
    elements.append(Paragraph("<u><b>Article 2 – Engagement de l'Assureur</b></u>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("""
    LEADWAY ASSURANCE s'engage à payer à première demande du bénéficiaire les sommes dues en cas de défaillance 
    de l'entreprise, dans la limite du montant garanti.
    """, style_normal))
    elements.append(Spacer(1, 15))

    # Article 3 - Conditions Financières (SANS SAUT DE PAGE)
    elements.append(Paragraph("<u><b>Article 3 – Conditions Financières</b></u>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("<b>Détail prime</b>", style_normal))
    elements.append(Spacer(1, 5))
    
    prime_art3_data = [
        ["Prime nette", "Frais d'analyse", "Accessoires", "Taxes", "Prime TTC"],
        [fmt_money_no_currency(data['prime_nette']),
         fmt_money_no_currency(data['frais_analyse']),
         fmt_money_no_currency(data['accessoires']),
         fmt_money_no_currency(data['taxes']),
         fmt_money_no_currency(data['prime_ttc'])]
    ]
    prime_art3_table = Table(prime_art3_data, colWidths=[90, 90, 90, 90, 90])
    prime_art3_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
    ]))
    elements.append(prime_art3_table)
    elements.append(Spacer(1, 10))
    elements.append(Paragraph("Payable avant le retrait de l'acte de caution.", style_normal))
    elements.append(Spacer(1, 15))
    
    # Article 4 - Sûretés Accessoires
    elements.append(Paragraph("<u><b>Article 4 – Sûretés Accessoires</b></u>", style_bold))
    elements.append(Spacer(1, 8))
    montant_depot_lettres = number_to_words(int(montant_contre_garantie/2))
    elements.append(Paragraph(f"""
    • Dépôt à terme de <b>{montant_depot_lettres} ({fmt_money(montant_contre_garantie/2)}) F CFA</b><br/>
    • Cautionnement personnel et solidaire des dirigeants<br/>
    • Billet à hauteur de l'engagement a signer
    """, style_normal))
    elements.append(Spacer(1, 15))
    
    # SAUT DE PAGE APRÈS L'ARTICLE 4
    elements.append(PageBreak())
    
    # Article 5
    elements.append(Paragraph("<u><b>Article 5 – Obligation d'Information</b></u>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("""
    Le donneur d'ordre doit :<br/><br/>
    • Communiquer les justificatifs de décaissement
    """, style_normal))
    elements.append(Spacer(1, 15))
    
    # Article 6
    elements.append(Paragraph("<u><b>Article 6 – Retrait de l'Acte</b></u>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("""
    Une fois l'acte retiré, la prime est acquise sauf cas de force majeure empêchant l'utilisation. 
    Les frais d'étude et de dossier restent dus.
    """, style_normal))
    elements.append(Spacer(1, 15))
    
    # Article 7
    elements.append(Paragraph("<u><b>Article 7 – Subrogation</b></u>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("""
    L'assureur est subrogé dans les droits du bénéficiaire en cas de paiement. Le Souscripteur perd 
    le bénéfice de la garantie s'il empêche la subrogation.
    """, style_normal))
    elements.append(Spacer(1, 15))
    
    # Article 8
    elements.append(Paragraph("<u><b>Article 8 – Durée de la Garantie</b></u>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph(f"""
    <b>La caution est valable du {data['date_effet']} au {data['date_echeance']}, sauf libération anticipée.</b>
    """, style_normal))
    elements.append(Spacer(1, 15))

    # Article 9 (SANS SAUT DE PAGE)
    elements.append(Paragraph("<u><b>Article 9 – Conditions d'Appel de la Garantie</b></u>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("""
    <b>La garantie est appelée en cas de défaillance avérée de l'entreprise : incapacité à exécuter le contrat ou 
    rembourser l'avance non amortie, notamment en cas de redressement, liquidation ou force majeure.</b>
    """, style_normal))
    elements.append(Spacer(1, 15))
    
    # Article 10
    elements.append(Paragraph("<u><b>ARTICLE 10 : Restitution du déposit :</b></u>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("""
    Au cas où un dépôt est constitué dans les livres de <b>LEADWAY ASSURANCE IARD</b>, la restitution se fera 
    sur demande expresse du Donneur d'Ordre. Cette demande doit être accompagnée de <b>l'original de l'acte de 
    cautionnement délivré avec la mention « bon pour mainlevée » ou de l'acte de mainlevée délivré par le bénéficiaire.</b><br/>
    Les sommes dues par le Donneur d'Ordre sont prélevées d'office sur le dépôt, le solde lui étant restitué.
    """, style_normal))
    elements.append(Spacer(1, 15))
    
    # Article 11
    elements.append(Paragraph("<u><b>Article 11 – Exclusions</b></u>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("""
    • <b>Non-respect des obligations contractuelles en dehors des cas prévus</b><br/>
    • <b>Utilisation détournée de l'avance par le Souscripteur.</b>
    """, style_normal))
    elements.append(Spacer(1, 40))
    
    # Signatures finales
    sig_img_final = Image(SIGNATURE_PATH, width=170, height=140) if os.path.exists(SIGNATURE_PATH) else ""
    sig_table_final = Table([
        [Paragraph("<b>LE SOUSCRIPTEUR</b>", style_normal),
         Paragraph("<b>POUR L'ASSUREUR</b>", style_normal)],
        ["", ""],
        ["", sig_img_final],
    ], colWidths=[280, 220])
    sig_table_final.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    elements.append(sig_table_final)

    doc.build(elements, onFirstPage=add_footer, onLaterPages=add_footer)
    buffer.seek(0)
    return buffer

# ============================
# PDF CONTRAT
# ============================
def generate_contrat_pdf(data, lots_data=None):
    buffer = BytesIO()
    
    # Fonction pour ajouter le bas de page
    def add_footer(canvas, doc):
        canvas.saveState()
        try:
            # Bas de page
            if os.path.exists(BAS_DE_PAGE_PATH):
                pil_img = PILImage.open(BAS_DE_PAGE_PATH)
                w, h = pil_img.size
                ratio = h / w
                footer_width = A4[0]
                footer_height = footer_width * ratio
                canvas.drawImage(BAS_DE_PAGE_PATH, 0, 0, 
                               width=footer_width, height=footer_height,
                               preserveAspectRatio=True, mask='auto')
        except Exception as e:
            pass
        canvas.restoreState()
    
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=40, leftMargin=40,
                            topMargin=30, bottomMargin=70)
    styles = getSampleStyleSheet()
    elements = []

    style_title = ParagraphStyle('TitleCenter', parent=styles['Title'],
                                 alignment=1, fontSize=14, spaceAfter=20,
                                 fontName="Helvetica-Bold")
    style_center = ParagraphStyle('Center', alignment=1, fontSize=11, spaceAfter=15)
    style_bold = ParagraphStyle('Bold', fontName='Helvetica-Bold',
                                fontSize=10, leading=12, spaceAfter=8, alignment=4)
    style_normal = ParagraphStyle('Normal', fontSize=10, leading=12, alignment=4)

    # PAGE 1
    try:
        pil_img = PILImage.open(LOGO_PATH)
        w, h = pil_img.size
        ratio = h / w
        target_w = 50 * mm
        target_h = target_w * ratio
        logo = Image(LOGO_PATH, width=target_w, height=target_h)
        logo.hAlign = 'CENTER'
        elements.append(logo)
        elements.append(Spacer(1, 12))
    except Exception as e:
        st.warning(f"Logo contrat : {e}")

    elements.append(Paragraph(data["assure"].upper(), style_title))
    elements.append(Paragraph("CONDITIONS PARTICULIERES ET GENERALES", style_title))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("CAUTION", style_title))
    elements.append(Paragraph(f"POLICE NUMERO {data['police_num']}", style_center))

    # PAGE 2
    elements.append(PageBreak())

    elements.append(Spacer(1, 10))

    info_data = [
        ["SOUSCRIPTEUR", f": {data['souscripteur']}"],
        ["ASSURÉ", f": {data['assure']}"],
        ["INTERMEDIAIRE", ": DIRECT"],
        ["CODE", ": 2000"],
        ["DATE D'EMISSION", f": {data['date_emission']}"],
        ["DATE D'EFFET", f": {data['date_effet']}"],
        ["DATE D'ÉCHÉANCE", f": {data['date_echeance']}"],
        ["DURÉE DE LA POLICE", f": {data['duree_police']}"],
    ]
    info_table = Table(info_data, colWidths=[150, 380])
    info_table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 15))

    # DÉCOMPTE DE PRIME
    elements.append(Paragraph("<b>DÉCOMPTE DE PRIME (en F CFA) :</b>", style_bold))
    elements.append(Spacer(1, 8))
    
    def fmt_money_no_currency(val):
        try:
            val = float(val)
            return f"{int(round(val)):,}".replace(",", " ")
        except:
            return "0"
    
    style_cell_contrat = ParagraphStyle(name='CellContrat', fontName='Helvetica',
                                        fontSize=10, alignment=1, leading=12)
    style_bold_cell_contrat = ParagraphStyle(name='BoldCellContrat',
                                              parent=style_cell_contrat, fontName='Helvetica-Bold')
    
    prime_data = [
        [Paragraph("Prime HT", style_bold_cell_contrat), Paragraph("Acc.", style_bold_cell_contrat),
         Paragraph("Frais d'analyse", style_bold_cell_contrat), Paragraph("Taxe", style_bold_cell_contrat),
         Paragraph("Prime TTC", style_bold_cell_contrat)],
        [Paragraph(fmt_money_no_currency(data['prime_nette']), style_cell_contrat),
         Paragraph(fmt_money_no_currency(data['accessoires']), style_cell_contrat),
         Paragraph(fmt_money_no_currency(data['frais_analyse']), style_cell_contrat),
         Paragraph(fmt_money_no_currency(data['taxes']), style_cell_contrat),
         Paragraph(f"<b>{fmt_money_no_currency(data['prime_ttc'])}</b>", style_cell_contrat)]
    ]
    prime_table = Table(prime_data, colWidths=[100, 80, 100, 80, 120])
    prime_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    elements.append(prime_table)
    elements.append(Spacer(1, 12))

    elements.append(Paragraph(f"""
    Aux conditions générales de la police de cautionnement de marché, aux conditions spéciales 
    et particulières qui suivent, <b>LEADWAY ASSURANCE IARD</b> garantit l'Assuré 
    <b>{data['assure']}</b> aux conditions ci-dessous
    """, style_normal))
    elements.append(Spacer(1, 12))

    # Articles 1-6
    montant_lettres = number_to_words(int(data['montant_caution']))

    if data['couverture'] == "Avance sur démarrage":
        article2_text = f"Le montant de la garantie de restitution d'avance est de <b>{fmt_money(data['montant_caution'])}</b> ({montant_lettres} francs CFA)."
    else:
        article2_text = f"Le montant de la garantie est de <b>{fmt_money(data['montant_caution'])}</b> ({montant_lettres} francs CFA)."

    articles = [
        ("ARTICLE 1 : OBJET DE LA GARANTIE",
         f"Le présent contrat a pour objet de garantir le bénéficiaire <b>{data.get('autorite', 'N/A')}</b> "
         f"contre les défaillances de l'Assuré en cas de non-exécution des prestations faisant l'objet "
         f"du marché <b>{data.get('objet', 'N/A')}</b>."),

        ("ARTICLE 2 : MONTANT DE LA GARANTIE", article2_text),

        ("ARTICLE 3 : L'ETENDUE DE LA GARANTIE",
         f"Le présent contrat couvre l'Assuré contre l'acompte perçu du maître d'ouvrage. "
         f"Elle s'épuise au fur et à mesure de l'exécution des travaux pour la caution d'avance de démarrage. "
         f"Toutefois, elle s'épuise après la réception des travaux pour les autres cautions de marché."),

        ("ARTICLE 4 : DURÉE",
         f"Le présent contrat prend effet le <b>{data['date_effet']}</b> "
         f"et prend fin le <b>{data['date_echeance']}</b>."),

        ("ARTICLE 5 : PAIEMENT DES PRIMES À LEADWAY ASSURANCE IARD",
         "Les modalités de paiement de la prime par l'Assuré à <b>LEADWAY ASSURANCE IARD</b> "
         "sont définies et arrêtées comme le stipule l'article 13 nouveau du Code CIMA. "
         "Pas de prime, pas de garantie. L'Assuré est tenu de payer la totalité de la prime "
         "à la délivrance de la caution. Une fois l'acte de caution retiré, la prime ne peut être restituée."),

        ("ARTICLE 6 : OBLIGATIONS D'INFORMATION",
         f"Le Donneur d'Ordre s'engage à transmettre à <b>LEADWAY ASSURANCE IARD</b> "
         f"l'ordre de service dès sa réception. <b>{data['assure']}</b> s'engage à informer régulièrement "
         f"<b>LEADWAY ASSURANCE IARD</b> de l'état d'avancement du marché. "
         f"Après chaque décompte, la société <b>{data['assure']}</b> doit transmettre une copie certifiée "
         f"à <b>LEADWAY ASSURANCE IARD</b> au plus tard dans les 48 heures qui suivent le décompte. "
         f"La non-transmission des documents demandés dans les délais convenus entraînera une amende forfaitaire. "
         f"<b>LEADWAY ASSURANCE IARD</b> a le droit d'exiger de l'Assuré la communication de tous documents "
         f"relatifs aux opérations cautionnées et elle a le droit de procéder à toutes vérifications utiles "
         f"afin de contrôler la sincérité et l'exactitude des déclarations du Donneur d'Ordre.")
    ]

    for title, txt in articles:
        elements.append(Paragraph(f"<b><u>{title}</u></b>", style_bold))
        elements.append(Spacer(1, 8))
        elements.append(Paragraph(txt, style_normal))
        elements.append(Spacer(1, 10))

    # Articles 7-10
    elements.append(Paragraph("<b><u>ARTICLE 7 : VISITE DE CHANTIER</u></b>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph(f"""
    Les parties conviennent d'organiser ensemble au moins deux (02) visites de chantier par an. 
    Ces visites sont organisées à l'initiative de la partie la plus diligente. 
    Les charges relatives à la visite sont supportées par la société <b>{data['assure']}</b> 
    pour seulement deux (02) agents de <b>LEADWAY ASSURANCE IARD</b>.
    """, style_normal))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("<b><u>ARTICLE 8 : MAIN LEVEE</u></b>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph(f"""
    La société <b>{data['assure']}</b> s'engage à diligenter par le Maître d'Ouvrage 
    d'une lettre de mainlevée qui doit être transmise à <b>LEADWAY ASSURANCE IARD</b>.
    """, style_normal))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("<b><u>ARTICLE 9 : RESTITUTION DU DÉPÔT</u></b>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("""
    Au cas où un dépôt est constitué dans les livres de <b>LEADWAY ASSURANCE IARD</b>, 
    la restitution se fera sur demande expresse du Donneur d'Ordre. 
    Cette demande doit être accompagnée de l'original de l'acte de cautionnement délivré 
    avec la mention « Bon pour mainlevée » ou de l'acte de mainlevée délivré par le bénéficiaire. 
    Les sommes dues par le Donneur d'Ordre sont prélevées d'office sur le dépôt, 
    le solde lui étant restitué.
    """, style_normal))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("<b><u>ARTICLE 10 : SUBROGATION</u></b>", style_bold))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("""
    <b>LEADWAY ASSURANCE IARD</b>, qui a payé l'indemnité d'assurance, est subrogée, 
    jusqu'à concurrence de cette indemnité, dans les droits et actions du bénéficiaire 
    de la caution envers qui l'Assuré a été défaillant. 
    <b>LEADWAY ASSURANCE IARD</b> peut être déchargée de tout ou partie de sa garantie 
    envers l'Assuré lorsque la subrogation ne peut plus, par le fait de l'Assuré, 
    s'opérer en faveur de l'Assureur.
    """, style_normal))
    elements.append(Spacer(1, 30))

    # Signature
    sig_img = Image(SIGNATURE_PATH, width=170, height=140) if os.path.exists(SIGNATURE_PATH) else ""
    sig_table = Table([
        [Paragraph("Le Donneur d'Ordre (Assuré)", style_normal),
         Paragraph("Le Garant", style_normal)],
        ["", Paragraph("(L'Assureur)", style_normal)],
        ["", ""],
        ["", sig_img],
    ], colWidths=[280, 220])
    sig_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LINEABOVE', (1,2), (1,2), 0.5, colors.black),
    ]))
    elements.append(sig_table)

    # CONDITIONS GÉNÉRALES
    elements.append(PageBreak())
    
    style_title_cg = ParagraphStyle('TitleCG', fontName='Helvetica-Bold',
                                    fontSize=12, alignment=1, spaceAfter=15)
    
    elements.append(Paragraph("<b>CONDITIONS GENERALES</b>", style_title_cg))
    elements.append(Spacer(1, 10))
    
    # TITRE I
    elements.append(Paragraph("<b>TITRE I : DISPOSITIONS GENERALES</b>", style_bold))
    elements.append(Spacer(1, 8))
    
    # Articles CG - TITRE I
    cg_titre1 = [
        ("Article 1 : Définitions des termes",
         "<b>Donneur d'ordre (Assuré)</b> : La personne à la demande de laquelle il est émis un acte de cautionnement.<br/>"
         "<b>Garant</b> : L'émetteur de l'acte de cautionnement ou de garantie ci-après dénommé « LEADWAY ASSURANCE IARD »<br/>"
         "<b>Bénéficiaire/Maître d'ouvrage</b> : organisme au profit duquel l'acte de cautionnement ou de garantie est émis."),
        
        ("Article 2 : Objet",
         "La présente police a pour objet la définition des conditions générales d'émission, à la demande du Donneur d'Ordre, "
         "d'engagements de signature par le Garant dans le cadre des marchés de travaux ou de prestations de services. "
         "Elle est complétée, précisée ou modifiée par les « Conditions particulières » qui sont convenues pour tous les actes "
         "de cautionnement délivrés par LEADWAY ASSURANCE IARD au profit du Bénéficiaire/Maître d'ouvrage désigné à ces mêmes conditions générales."),
        
        ("Article 3 : Dispositions contractuelles",
         "Les relations entre les parties sont régies par les présentes conditions générales et par tous les accords dont les parties "
         "pourraient convenir. Dans le silence de leurs conventions, les parties se réfèrent aux dispositions du contrat d'assurance "
         "telles stipulées dans le Livre I du CODE CIMA ainsi que le CODE DES MARCHES PUBLIQUES au/ou l'Acte Uniforme portant "
         "organisation des sûretés en ses articles 3 à 38."),
        
        ("Article 4 : Durée et entrée en vigueur du contrat",
         "Le présent contrat est conclu pour la durée de soumission à l'appel d'offres pour la caution de soumission jusqu'à "
         "l'adjudication de l'offre, toutefois il prend effet à partir de la signature du contrat d'exécution des travaux et ce, "
         "jusqu'à la réception définitive des travaux."),
        
        ("Article 5 : Champ d'application",
         "Sont garantis par l'Assureur caution et pouvant être demandés par le Donneur d'ordre au Garant, les cautionnements ou "
         "garanties de soumission, d'avance de démarrage, de bonne exécution et de retenue de garantie ou de toute autre nature ou "
         "appellation qui peuvent être demandés dans le marché de références. Elles s'appliquent aux garanties qui sont demandées par "
         "le Donneur d'ordre au Garant sont des personnes physiques ou morales, de droit public ou de droit privé, nationaux ou étrangers.")
    ]
    
    for title, txt in cg_titre1:
        elements.append(Paragraph(f"<b><u>{title}</u></b>", style_bold))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph(txt, style_normal))
        elements.append(Spacer(1, 10))
    
    # TITRE II
    elements.append(Spacer(1, 5))
    elements.append(Paragraph("<b>TITRE II : DELIVRANCE DES CAUTIONNEMENTS</b>", style_bold))
    elements.append(Spacer(1, 8))
    
    elements.append(Paragraph("<b><u>Article 6 : Demande de cautionnement- Documents à fournir</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    La délivrance des polices de cautionnement est faite sur demande du Donneur d'ordre. Cette demande doit être accompagnée 
    des pièces permettant au Garant d'émettre une offre, le cas échéant, son acte de cautionnement conformément aux prescriptions 
    du dossier d'appel d'offre (DAO). À titre indicatif, le Donneur d'ordre devra accompagner sa demande :
    """, style_normal))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    • Pour les cautionnements de soumission : une demande formelle, une copie du dossier particulier d'appel d'offre (DPAO) et le modèle de l'acte de cautionnement à délivrer.<br/>
    • Pour les cautionnements d'avance de démarrage : une demande formelle, une copie du contrat de marché signé entre le Donneur d'ordre et le Bénéficiaire/Maître d'ouvrage et le modèle de l'acte de cautionnement à délivrer.<br/>
    • Pour les cautionnements de retenue de garantie : une demande formelle, une copie du procès-verbal de réception provisoire des travaux et le modèle de l'acte de cautionnement à délivrer.<br/>
    • Pour les cautionnements de bonne exécution : une demande formelle, une copie du contrat de marché signé entre le Donneur d'ordre et le Bénéficiaire/Maître d'ouvrage et le modèle de l'acte de cautionnement à délivrer.
    """, style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b><u>Article 7 : Délivrance des actes de cautionnement</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    Après étude du dossier du Donneur d'ordre, LEADWAY ASSURANCE IARD délivre éventuellement le cautionnement qui lui est demandé. 
    En cas d'acceptation du Garant de la délivrance des actes de cautionnement dans les conditions habituelles convenues avec le 
    Donneur d'ordre. Le dernier est informé par LEADWAY ASSURANCE IARD par les moyens les plus rapides pour procéder aux retraits 
    des actes de cautionnement à son siège. Au cas où l'acceptation de la délivrance des cautionnements demandés est assujettie à 
    des conditions différentes de celles habituellement pratiquées, LEADWAY ASSURANCE IARD, après en avoir informé le Donneur d'ordre, 
    est tenu de le lui notifier par lettre recommandée avec accusé de réception.
    """, style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b><u>Article 8 : Modalités de délivrance des actes de cautionnement</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    Sauf convention expresse entre les parties, les actes de cautionnement sont délivrés après satisfaction des conditions convenues 
    entre les parties et paiement de la facture y afférente.
    """, style_normal))
    elements.append(Spacer(1, 15))
    
    # TITRE III
    elements.append(Paragraph("<b>TITRE III : OBLIGATION DES PARTIES</b>", style_bold))
    elements.append(Spacer(1, 8))
    
    elements.append(Paragraph("<b><u>Article 9 : Obligation de diligence</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    Le Garant devra répondre avec diligence aux demandes de cautionnements qui lui sont faites par le Donneur d'ordre. 
    Il s'engage à lui donner une réponse dans les 5 jours ouvrés suivant la date du dépôt de la demande et des documents complets 
    et en pièces lui fournir.
    """, style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b><u>Article 10 : Obligation de conformité</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    Le Garant s'oblige avant toute intervention relative à un acte de cautionnement d'en informer le Donneur d'ordre par le 
    transmission d'une copie de la correspondance du Bénéficiaire/Maître d'ouvrage.
    """, style_normal))
    elements.append(Spacer(1, 15))
    
    # TITRE IV
    elements.append(Paragraph("<b>TITRE IV : OBLIGATIONS DU DONNEUR D'ORDRE</b>", style_bold))
    elements.append(Spacer(1, 8))
    
    elements.append(Paragraph("<b><u>Article 11 : Obligation de paiement de primes</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    Le Donneur d'ordre est tenu au paiement de la prime qui constitue la rémunération de LEADWAY ASSURANCE IARD. 
    Sauf convention expresse entre les parties, la prime est payée concomitamment au retrait des actes de cautionnement au siège 
    de LEADWAY ASSURANCE IARD selon les dispositions de l'article 13 nouveau du CODE CIMA. Une fois la prime payée, elle ne peut 
    être restituée sauf si le Donneur d'ordre, pour des raisons imputables au Bénéficiaire/Maître d'ouvrage de la caution ou au 
    Garant, n'a pas pu jouir de l'avantage du cautionnement. Dans ce dernier cas, la restitution portera sur la prime, exceptée 
    les droits d'ouverture de dossier. Il sera également tenu compte au délai pendant lequel le Donneur d'ordre aura gardé par 
    devers lui l'acte de cautionnement, tout trimestre commencé étant dû. Toute augmentation de la durée de validité du 
    cautionnement sera facturée au Donneur d'ordre qui devra régler le complément de la prime, si la perception d'une prime 
    complémentaire calculée prorata temporis.
    """, style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b><u>Article 12 : Obligation de diligence</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    Le Donneur d'ordre s'oblige à exécuter le contrat pour lequel le Garant a donné son cautionnement conformément aux prescriptions 
    du Bénéficiaire/Maître d'ouvrage. Il s'engage à prendre toutes les dispositions utiles pour qu'il ne puisse lui être reproché 
    aucun manquement dans l'exécution des obligations pour lesquelles il a obtenu le cautionnement du Garant. Le donneur d'ordre 
    s'engage pour toute la durée de la présente police à introduire auprès de LEADWAY ASSURANCE IARD toute demande d'augmentation 
    de son cautionnement ou tout nouveau cautionnement exigé par le même Bénéficiaire/Maître d'ouvrage conformément aux dispositions 
    du Code des Assurances relatives aux modifications substantielles des circonstances du contrat. Le fausse déclaration et 
    intentionnelle des capitaux pouvant donner lieu à l'application de la règle proportionnelle de capitaux sur des primes.
    """, style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b><u>Article 13 : Obligation d'information</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    Le Donneur d'ordre s'oblige à tenir informé périodiquement le Garant des dispositions et/ou de l'œuvre pour la bonne réalisation 
    de laquelle lequel le Garant a pris le cautionnement de l'Assureur Caution. Il est tenu également de convier l'Assureur Caution 
    ou son préposé à visiter et inspecter les chantiers et d'organiser avec les différentes parties prenantes la réalisation du 
    marché garanti. Le Donneur d'ordre s'engage en outre à fournir annuellement à l'Assureur Caution Garant ses états financiers 
    annuels certifiés ou approuvés par les organes de contrôle.
    """, style_normal))
    elements.append(Spacer(1, 15))
    
    # TITRE V
    elements.append(Paragraph("<b>TITRE V : INTERVENTION ET RECOURS DE L'ASSUREUR CAUTION</b>", style_bold))
    elements.append(Spacer(1, 8))
    
    elements.append(Paragraph("<b><u>Article 14 : Intervention de l'Assureur caution</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    Lorsque le Bénéficiaire/Maître d'ouvrage demande l'intervention de l'Assureur Caution Garant, il en fait information au 
    Donneur d'ordre qui pourra lui faire opposition, à condition de présentation de pièces régulières établissant l'exécution de 
    ses obligations. Aussi, dans les 72 heures qui suivent la réception de cette information, le Donneur d'ordre est tenu de faire 
    part à l'Assureur Caution de ses appréciations sur la demande du Bénéficiaire/Maître d'ouvrage. A défaut, l'Assureur Caution 
    se réserve le droit de répondre utilement à la demande du Bénéficiaire/Maître d'ouvrage. Le Donneur d'ordre ne pourra opposer 
    à l'Assureur Caution la montant toutes mesures conservatoires au cas où il serait invité à intervenir comme caution ou dès qu'il 
    est averti d'une défaillance du donneur d'ordre vis-à-vis du Bénéficiaire/Maître d'ouvrage.
    """, style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b><u>Article 15 : Indemnisation de l'Assureur Caution</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    En cas de paiement au Bénéficiaire/Maître d'ouvrage, le Donneur d'ordre est tenu de rembourser à l'Assureur Caution le montant 
    total de son intervention, y compris tous les frais et dépenses judiciaires, extrajudiciaires. Dès l'instant qu'un paiement aura 
    été effectué au Bénéficiaire/Maître d'ouvrage, le Donneur d'ordre cède tout droit de créance à Assureur Caution, à concurrence 
    des montants payés. La notification au Donneur des pièces de paiement de LEADWAY ASSURANCE IARD au Bénéficiaire/Maître d'ouvrage 
    vaudront pour le débiteur, une preuve de la cession. Il pourra, par conséquent, se désintéresser toute réclamation de l'Assureur Caution.
    """, style_normal))
    elements.append(Spacer(1, 15))
    
    # TITRE VI
    elements.append(Paragraph("<b>TITRE VI : DISPOSITIONS FINALES</b>", style_bold))
    elements.append(Spacer(1, 8))
    
    elements.append(Paragraph("<b><u>Article 16 : Circulation du contrat</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("Le présent contrat est soumis au droit CIMA.", style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b><u>Article 17 : Résiliation du contrat</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    Le contrat peut être résilié par chacune des parties. La partie qui prend l'initiative de la résiliation est tenue de servir à 
    son cocontractant un préavis, trois (3) mois avant la fin de la période annuelle en cours. Le contrat est également résilié de 
    plein droit en cas de cessation d'activités du Donneur d'ordre ou en cas d'un prononcé à son encontre d'un jugement de cessation 
    de paiement ou de la constatation de n'importe quel autre procédé destiné à lévier ou retracer.
    """, style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b><u>Article 18 : Clause d'arbitrage</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    Tout différend ou contestation qui pourrait survenir entre les parties du fait ou au sujet de l'application du présent contrat 
    pourra être réglé à l'amiable ou par les instances compétentes des Marchés Publics par la négociation sera soumis au Tribunal 
    de Première Instance d'Abidjan.
    """, style_normal))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b><u>Article 19 : Election de domicile</u></b>", style_bold))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("""
    Pour l'exécution des présentes, les parties font élection de domicile à savoir :<br/>
    ◆ LEADWAY ASSURANCE IARD: Siège Social : Angré 7ème tranche, près du Centre Commercial TERA<br/>
    ◆ LE DONNEUR D'ORDRE, dont les références sont données aux Conditions Particulières
    """, style_normal))

    doc.build(elements, onFirstPage=add_footer, onLaterPages=add_footer)
    buffer.seek(0)
    return buffer

# ============================
# UI STREAMLIT
# ============================
st.markdown("<h1 style='text-align:center;color:#000;'>ASSUR DEFENDER - Caution</h1>", unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# Durée & Type
col1, col2 = st.columns([1, 1])
with col1:
    duree = st.selectbox("Durée", ["30 jours", "90 jours", "150 jours", "180 jours", "365 jours"], index=4)
with col2:
    type_caution = st.selectbox("Type de Caution", [
        "Soumission", 
        "Avance sur démarrage", 
        "Bonne exécution",
        "Retenue de garantie", 
        "Provisoire", 
        "Intermédiaire d'assurance",
        "Agence de voyage", 
        "Fondateur d'établissement",
        "Caution d'agrément"
    ])

# Champ Détail pour Caution d'agrément
detail_agrement = ""
if type_caution == "Caution d'agrément":
    detail_agrement = st.text_input("Détail", key="detail_agrement")

# Infos assuré
st.markdown("### <span style='color:#8B00FF;'>Informations sur l'Assuré</span>", unsafe_allow_html=True)

# Question Assuré = Souscripteur
assure_est_souscripteur = st.radio("L'assuré est-il le souscripteur ?", ["Oui", "Non"], horizontal=True, index=0)

# Si Non, demander le nom du souscripteur
nom_souscripteur = ""
if assure_est_souscripteur == "Non":
    nom_souscripteur = st.text_input("Nom du Souscripteur", key="nom_souscripteur")

c1, c2, c3 = st.columns(3)
nom_assure = c1.text_input("Nom de l'Assuré")
siege_social = c2.text_input("Siège Social")
telephone = c3.text_input("Téléphone")

# Question Assuré = Bénéficiaire
assure_est_beneficiaire = st.radio("L'assuré est-il le bénéficiaire ?", ["Oui", "Non"], horizontal=True, index=0, key="beneficiaire_radio")

# Si Non, demander le nom et l'adresse du bénéficiaire
nom_beneficiaire = ""
adresse_beneficiaire = ""
if assure_est_beneficiaire == "Non":
    col_b1, col_b2 = st.columns(2)
    nom_beneficiaire = col_b1.text_input("Nom du Bénéficiaire", key="nom_beneficiaire")
    adresse_beneficiaire = col_b2.text_input("Adresse du Bénéficiaire", key="adresse_beneficiaire")

# Détails marché
st.markdown("### <span style='color:#8B00FF;'>Détails du Marché</span>", unsafe_allow_html=True)
c1, c2, c3 = st.columns(3)
situation_geo = c1.text_input("Situation Géographique du Marché")
numero_marche = c2.text_input("Numéro du Marché")
autorite_contractante = c3.text_input("Autorité Contractante")

c4, c5, c6 = st.columns(3)
date_depot = c4.date_input("Date de Dépôt du Dossier", disabled=(type_caution != "Soumission"))
date_depot_str = format_date_fr(date_depot) if type_caution == "Soumission" else "Selon contrat"
objet_marche = c5.text_input("Objet du Marché")
montant_marche = c6.number_input("Montant du Marché (FCFA)", min_value=0.0, step=1000.0, format="%.0f")

# Lots
st.markdown("### <span style='color:#8B00FF;'>Détails des Lots (Optionnel)</span>", unsafe_allow_html=True)

# Types de caution qui ne nécessitent pas de lots
types_sans_lots = ["Intermédiaire d'assurance", "Agence de voyage", "Fondateur d'établissement", "Caution d'agrément"]

if type_caution in types_sans_lots:
    # Pour ces types, demander simplement le montant à cautionner
    montant_total_caution = st.number_input("Montant à Cautionner (FCFA)",
                                            min_value=0.0, step=1000.0, format="%.0f", key="montant_simple")
    lots_data = []
else:
    # Pour les autres types, proposer la gestion des lots
    nb_lots = st.number_input("Nombre de lots", min_value=0, step=1, value=0)
    lots_data = []
    montant_total_caution = 0.0

    if nb_lots > 0:
        for i in range(1, nb_lots + 1):
            with st.expander(f"Lot {i}"):
                cc1, cc2, cc3 = st.columns([1, 1, 2])
                lot_num = cc1.text_input(f"Numéro du lot {i}", key=f"lotnum_{i}")
                lot_mont = cc2.number_input(f"Montant à cautionner {i}", min_value=0.0,
                                            step=1000.0, key=f"lotmont_{i}", format="%.0f")
                lot_desc = cc3.text_input(f"Désignation {i}", key=f"lotdesc_{i}")
                montant_total_caution += lot_mont
                lots_data.append({"Lot": lot_num, "Montant": lot_mont, "Désignation": lot_desc})
        st.info(f"Montant total calculé : **{fmt_money(montant_total_caution)}**")
    else:
        montant_total_caution = st.number_input("Montant total à Cautionner (FCFA)",
                                                min_value=0.0, step=1000.0, format="%.0f")

# Tarification
st.markdown("### <span style='color:#8B00FF;'>Tarification</span>", unsafe_allow_html=True)
col1, col2, col3, col4 = st.columns(4)
taux_tarif = col1.number_input("Taux (%)", min_value=0.0, step=0.01, value=0.1)
reduction = col2.number_input("Réduction (%)", min_value=0.0, step=0.1, value=0.0)
accessoires_plus = col3.number_input("Accessoires + (FCFA)", min_value=0.0, step=1000.0, format="%.0f")
frais_analyse = col4.number_input("Frais d'analyse (FCFA)", min_value=0.0, step=1000.0, format="%.0f")

# Sûretés
suretes_input = ""
if type_caution != "Soumission":
    st.markdown("### <span style='color:#8B00FF;'>Sûretés (Optionnel)</span>", unsafe_allow_html=True)
    suretes_input = st.text_area("Conditions (une par ligne)", height=120)

# Génération Cotation
if st.button("Générer la Cotation", type="primary", use_container_width=True):
    if not nom_assure.strip() or montant_total_caution <= 0:
        st.error("Nom de l'Assuré et Montant à cautionner obligatoires.")
    else:
        taux_eff = taux_tarif / 100
        red_eff = reduction / 100
        prime_nette = taux_eff * montant_total_caution * (1 - red_eff)
        if prime_nette <= 100_000:
            accessoires_base = 5_000
        elif prime_nette <= 500_000:
            accessoires_base = 7_500
        elif prime_nette <= 1_000_000:
            accessoires_base = 10_000
        elif prime_nette <= 5_000_000:
            accessoires_base = 15_000
        elif prime_nette <= 10_000_000:
            accessoires_base = 20_000
        elif prime_nette <= 50_000_000:
            accessoires_base = 30_000
        else:
            accessoires_base = 50_000
        accessoires = accessoires_base + accessoires_plus
        taxes = 0.145 * (prime_nette + accessoires + frais_analyse)
        prime_ttc = prime_nette + accessoires + frais_analyse + taxes

        date_cotation_str = format_date_fr(datetime.date.today())

        data = {
            "assure": nom_assure,
            "souscripteur": nom_souscripteur if nom_souscripteur else nom_assure,
            "beneficiaire": nom_beneficiaire if nom_beneficiaire else nom_assure,
            "adresse_beneficiaire": adresse_beneficiaire if adresse_beneficiaire else siege_social,
            "adresse": siege_social or "N/A",
            "situation_geo": situation_geo or "N/A",
            "num_marche": numero_marche or "N/A",
            "autorite": autorite_contractante or "N/A",
            "date_depot": date_depot_str,
            "objet": objet_marche or "N/A",
            "couverture": type_caution,
            "montant_marche": montant_marche,
            "duree": duree,
            "montant_caution": montant_total_caution,
            "prime_nette": prime_nette,
            "frais_analyse": frais_analyse,
            "accessoires": accessoires,
            "taxes": taxes,
            "prime_ttc": prime_ttc,
            "date_cotation": date_cotation_str,
            "suretes_text": suretes_input,
        }

        pdf_buffer = generate_caution_pdf(data, lots_data)
        st.success("Cotation PDF générée !")
        st.download_button("Télécharger Cotation", pdf_buffer,
                           f"Cotation_{nom_assure.replace(' ', '_')}.pdf",
                           "application/pdf")

        # Sauvegarde Supabase Étape 1
        new_cotation_id, message = save_cotation_to_supabase(data, lots_data, detail_agrement)
        
        if new_cotation_id:
            st.success(f"Cotation enregistrée dans Supabase (ID: {new_cotation_id}).")
            st.session_state.cotation_data = data
            st.session_state.lots_data = lots_data
            st.session_state.cotation_db_id = new_cotation_id # Stocker l'ID BDD
        else:
            st.error(f"Échec de l'enregistrement Supabase (Cotation): {message}")
            # Réinitialiser au cas où l'enregistrement échoue
            if "cotation_data" in st.session_state:
                del st.session_state.cotation_data
            if "lots_data" in st.session_state:
                del st.session_state.lots_data
            if "cotation_db_id" in st.session_state:
                del st.session_state.cotation_db_id


# Génération Contrat
if "cotation_data" in st.session_state and "cotation_db_id" in st.session_state:
    st.markdown("---")
    if st.button("Générer le Contrat", type="secondary", use_container_width=True):
        data = st.session_state.cotation_data
        cotation_db_id = st.session_state.cotation_db_id # Récupérer l'ID BDD
        
        police_num = f"3240-800{str(uuid.uuid4().int)[:6]}25"
        today = datetime.date.today()
        contrat_data = {
            **data,
            "police_num": police_num,
            "date_emission": format_date_fr(today),
            "date_effet": format_date_fr(today),
            "date_echeance": format_date_fr(today + datetime.timedelta(days=364)),
            "duree_police": "365 jours",
        }

        # Vérifier si c'est une caution d'agrément
        if type_caution == "Caution d'agrément":
            pdf_contrat = generate_contrat_agrement_pdf(contrat_data, detail_agrement, st.session_state.lots_data)
        else:
            pdf_contrat = generate_contrat_pdf(contrat_data, st.session_state.lots_data)
        
        st.success(f"Contrat PDF généré – Police **{police_num}**")
        
        # Sauvegarde Supabase Étape 2
        success, message = save_police_to_supabase(cotation_db_id, contrat_data)
        if success:
            st.success(f"Police {police_num} enregistrée et liée à la cotation {cotation_db_id}.")
        else:
            st.error(f"Échec de l'enregistrement Supabase (Police): {message}")
        
        st.download_button("Télécharger Contrat", pdf_contrat,
                           f"Contrat_{police_num}.pdf", "application/pdf",
                           use_container_width=True)