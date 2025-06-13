from flask import Flask, request, render_template, redirect, url_for, session
import os
import datetime
import csv
import shutil
import fitz  # PyMuPDF
import re
from flaskwebgui import FlaskUI
from collections import defaultdict

# Abréviations FR 4 lettres pour les mois
MOIS_FR = {
    1: 'janv', 2: 'févr', 3: 'mars', 4: 'avr',
    5: 'mai', 6: 'juin', 7: 'juil', 8: 'août',
    9: 'sept', 10: 'oct', 11: 'nov', 12: 'déc'
}

app = Flask(__name__)
app.secret_key = 'archivagebl'

BASE_DIR = r"T:\SECRETARIAT\SECRETARIAT\PERMANENT\JUSTIFICATIFS FACTURES\LOUVET\Archivage_Bons"
FACTURE_DIR = os.path.join(BASE_DIR, "Factures")
UPLOAD_TEMP = "uploads_temp"
HISTO_FILE = os.path.join(BASE_DIR, "historique.csv")

app.config['UPLOAD_FOLDER'] = UPLOAD_TEMP
os.makedirs(BASE_DIR, exist_ok=True)
os.makedirs(FACTURE_DIR, exist_ok=True)
os.makedirs(UPLOAD_TEMP, exist_ok=True)

# Fonction de nettoyage
def sanitize_filename(name):
    return re.sub(r'[\/\\\:\*\?\"\<\>\|]', '-', name)

def extract_infos_from_pdf(filepath):
    try:
        text = ""
        with fitz.open(filepath) as doc:
            for page in doc:
                text += page.get_text()

        lines = text.splitlines()
        entreprise = chantier = engin = ""
        found_client = found_chantier = found_materiel = False

        for i, line in enumerate(lines):
            line_clean = line.strip()

            if not found_client and re.match(r'(?i)^client\s*[:\-–]', line_clean):
                parts = re.split(r'[:\-–]', line_clean, maxsplit=1)
                if len(parts) > 1:
                    entreprise = sanitize_filename(parts[1].strip())
                    found_client = True

            elif not found_chantier and re.match(r'(?i)^chantier\s*[:\-–]', line_clean):
                parts = re.split(r'[:\-–]', line_clean, maxsplit=1)
                if len(parts) > 1:
                    chantier = parts[1].strip()
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        if not re.match(r'(?i)^réf\.?\s*chantier\s*[:\-–]', next_line):
                            chantier += " " + next_line
                    chantier = sanitize_filename(chantier)
                    found_chantier = True

            elif not found_materiel and re.match(r'(?i)^mat[ée]riel\s*[:\-–]', line_clean):
                parts = re.split(r'[:\-–]', line_clean, maxsplit=1)
                if len(parts) > 1:
                    engin = sanitize_filename(parts[1].strip())
                    found_materiel = True

        return entreprise, chantier, engin

    except Exception as e:
        print(f"[ERREUR extraction PDF] {e}")
        return "", "", ""

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", client="", chantier="", machine="", last_file=session.get("last_file_name", ""))

@app.route("/analyser_pdf", methods=["POST"])
def analyser_pdf():
    fichier = request.files.get("pdf_file")
    if fichier and fichier.filename.endswith(".pdf"):
        chemin_temp = os.path.join(app.config['UPLOAD_FOLDER'], fichier.filename)
        fichier.save(chemin_temp)
        session["last_file_path"] = chemin_temp
        session["last_file_name"] = fichier.filename
        entreprise, chantier, engin = extract_infos_from_pdf(chemin_temp)
        return render_template("index.html", client=entreprise, chantier=chantier, machine=engin, last_file=fichier.filename)
    return redirect(url_for("index"))

@app.route("/upload", methods=["POST"])
def upload():
    client = sanitize_filename(request.form.get("client", "").strip())
    chantier = sanitize_filename(request.form.get("chantier", "").strip())
    machine = sanitize_filename(request.form.get("machine", "").strip())
    fichier = request.files.get("fichier")

    if not fichier and "last_file_path" in session:
        fichier = open(session["last_file_path"], "rb")
        fichier.filename = session.get("last_file_name")

    if fichier and client and chantier and machine:
        nom_original = os.path.splitext(fichier.filename)[0]
        extension = os.path.splitext(fichier.filename)[1]
        dossier_cible = os.path.join(BASE_DIR, client, chantier, machine)
        os.makedirs(dossier_cible, exist_ok=True)
        nom_fichier = f"{client}_{chantier}_{machine}_{nom_original}{extension}"
        chemin_fichier = os.path.join(dossier_cible, nom_fichier)
        with open(chemin_fichier, "wb") as f_out:
            shutil.copyfileobj(fichier, f_out)

        now = datetime.datetime.now()
        date = now.strftime("%d/%m/%Y")
        heure = now.strftime("%H:%M")
        write_header = not os.path.exists(HISTO_FILE) or os.path.getsize(HISTO_FILE) == 0

        with open(HISTO_FILE, mode="a", newline="", encoding="utf-8-sig") as csvfile:
            writer = csv.writer(csvfile, delimiter=';')
            if write_header:
                writer.writerow(["Date", "Heure", "Client", "Chantier", "Machine", "Fichier"])
            writer.writerow([date, heure, client, chantier, machine, nom_fichier])

    return redirect(url_for("index"))

@app.route("/facture", methods=["GET", "POST"])
def facture():
    if request.method == "POST":
        nom_facture = sanitize_filename(request.form["nom_facture"].strip())
        fichiers = request.form.getlist("fichiers")
        dossier_facture = os.path.join(FACTURE_DIR, nom_facture)
        os.makedirs(dossier_facture, exist_ok=True)
        for chemin_absolu in fichiers:
            if os.path.exists(chemin_absolu):
                shutil.move(chemin_absolu, os.path.join(dossier_facture, os.path.basename(chemin_absolu)))
        return redirect(url_for("facture"))

    fichiers_groupes = defaultdict(lambda: defaultdict(list))
    for root, _, files in os.walk(BASE_DIR):
        if FACTURE_DIR in root or UPLOAD_TEMP in root:
            continue
        for file in files:
            try:
                chemin = os.path.join(root, file)
                relative_path = os.path.relpath(chemin, BASE_DIR)
                parts = relative_path.split(os.sep)
                if len(parts) >= 3:
                    entreprise, chantier = parts[0], parts[1]
                    fichiers_groupes[entreprise][chantier].append((chemin, file))
            except Exception as e:
                print(f"[ERREUR FACTURE] {file} → {e}")

    return render_template("facture.html", fichiers_groupes=fichiers_groupes)


@app.route('/tri_bons', methods=['GET', 'POST'])
def tri_bons():
    if request.method == 'POST':
        date_sel = request.form.get('date_selection')
        if not date_sel:
            return render_template('index.html', error="Veuillez sélectionner une date.")
        year, month, _ = map(int, date_sel.split('-'))
        date_code = f"{MOIS_FR[month]}{str(year)[2:]}"
        # Collect PDF BL files by modification date
        bons_filtres = []
        for root, _, files in os.walk(BASE_DIR):
            if FACTURE_DIR in root or UPLOAD_TEMP in root:
                continue
            for file in files:
                if file.lower().endswith('.pdf'):
                    chemin = os.path.join(root, file)
                    dt = datetime.datetime.fromtimestamp(os.path.getmtime(chemin))
                    if dt.year == year and dt.month == month:
                        rel = os.path.relpath(chemin, BASE_DIR)
                        parts = rel.split(os.sep)
                        if len(parts) >= 3:
                            client, chantier, engin = parts[0], parts[1], parts[2]
                            bons_filtres.append((client, chantier, engin, chemin, file))
        # Grouping
        groupes = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
        for client, chantier, engin, chemin, filename in bons_filtres:
            groupes[client][chantier][engin].append((chemin, filename))
        return render_template('liste_bons.html', date_code=date_code, groupes=groupes)
    return render_template('index.html')

if __name__ == "__main__":
    ui = FlaskUI(app=app, server="flask", width=900, height=700)
    ui.run()
