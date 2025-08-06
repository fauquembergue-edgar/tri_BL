from flask import Flask, request, render_template, redirect, url_for, session, send_file, abort
import os
import datetime
import csv
import shutil
import fitz  # Librairie PyMuPDF pour lire les PDF
import re
from flaskwebgui import FlaskUI
from collections import defaultdict

# Dictionnaire contenant les abréviations françaises des mois (4 lettres)
MOIS_FR = {
    1: 'janv', 2: 'févr', 3: 'mars', 4: 'avri',
    5: 'mai', 6: 'juin', 7: 'juil', 8: 'août',
    9: 'sept', 10: 'octo', 11: 'nove', 12: 'déce'
}

# Liste utilisée pour identifier les noms de dossiers de mois avec accents en majuscule
MOIS_MAJ_ACCENTS = [
    'JANV', 'FÉVR', 'MARS', 'AVR', 'MAI', 'JUIN', 'JUIL', 'AOÛT', 'SEPT', 'OCT', 'NOV', 'DÉC'
]

# Initialisation de l'application Flask
app = Flask(__name__)
app.secret_key = 'archivagebl'

# Chemins des différents répertoires utilisés
BASE_DIR = r"T:\SECRETARIAT\SECRETARIAT\PERMANENT\JUSTIFICATIFS FACTURES\TPW\Archivage_Bons"
FACTURE_DIR = os.path.join(BASE_DIR, "Factures")
UPLOAD_TEMP = "uploads_temp"
HISTO_FILE = os.path.join(BASE_DIR, "historique.csv")

# Configuration du dossier d'upload temporaire
app.config['UPLOAD_FOLDER'] = UPLOAD_TEMP
os.makedirs(BASE_DIR, exist_ok=True)
os.makedirs(FACTURE_DIR, exist_ok=True)
os.makedirs(UPLOAD_TEMP, exist_ok=True)

# Fonction pour nettoyer les noms de fichiers
def sanitize_filename(name):
    return re.sub(r'[\/\\\:\*\?\"\<\>\|]', '-', name)

# Fonction pour extraire les informations (client, chantier, engin) à partir d’un fichier PDF
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
                    # Vérifie si le chantier continue sur la ligne suivante
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

# Fonction utilisée pour trier chronologiquement les dossiers nommés 'JANV25', 'AOÛT24', etc.
def mois_code_to_tuple(mois_code):
    mois_abrev = mois_code[:-2]
    annee = int(mois_code[-2:])
    mois_num = None
    for num, abbr in MOIS_FR.items():
        if abbr.upper() == mois_abrev:
            mois_num = num
            break
    if mois_num is None:
        mois_num = 99
    return (2000 + annee, mois_num)

# Route principale affichant le formulaire
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", client="", chantier="", machine="", last_file=session.get("last_file_name", ""))

# Route qui analyse le PDF et extrait automatiquement les champs à partir de son contenu
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

# Route qui gère l'enregistrement du fichier PDF dans le bon dossier selon les champs remplis
@app.route("/upload", methods=["POST"])
def upload():
    client = sanitize_filename(request.form.get("client", "").strip())
    chantier = sanitize_filename(request.form.get("chantier", "").strip())
    machine = sanitize_filename(request.form.get("machine", "").strip())
    date_archivage_str = request.form.get("date_archivage", "").strip()
    fichier = request.files.get("fichier")

    if not fichier and "last_file_path" in session:
        fichier = open(session["last_file_path"], "rb")
        fichier.filename = session.get("last_file_name")

    dossier_date = ""
    if date_archivage_str:
        try:
            date_archivage = datetime.datetime.strptime(date_archivage_str, "%Y-%m-%d")
            mois = date_archivage.month
            annee = date_archivage.year % 100
            dossier_date = f"{MOIS_FR[mois].capitalize()}{annee:02d}".upper()
            if dossier_date.startswith("Déc"):
                dossier_date = "DÉC" + dossier_date[3:]
            if dossier_date.startswith("Févr"):
                dossier_date = "FÉVR" + dossier_date[4:]
            if dossier_date.startswith("Août"):
                dossier_date = "AOÛT" + dossier_date[4:]
        except Exception as e:
            print(f"[ERREUR PARSING DATE] {e}")
            dossier_date = ""

    if fichier and client and chantier and machine and dossier_date:
        nom_original = os.path.splitext(fichier.filename)[0]
        extension = os.path.splitext(fichier.filename)[1]
        dossier_cible = os.path.join(BASE_DIR, dossier_date, client, chantier, machine)
        os.makedirs(dossier_cible, exist_ok=True)
        nom_fichier = f"{client}_{chantier}_{machine}_{nom_original}{extension}"
        chemin_fichier = os.path.join(dossier_cible, nom_fichier)
        with open(chemin_fichier, "wb") as f_out:
            shutil.copyfileobj(fichier, f_out)

        now = datetime.datetime.now()
        date = now.strftime("%d/%m/%Y")
        heure = now.strftime("%H:%M")
        write_header = not os.path.exists(HISTO_FILE) or os.path.getsize(HISTO_FILE) == 0

        # Enregistre les informations dans un fichier CSV d'historique
        with open(HISTO_FILE, mode="a", newline="", encoding="utf-8-sig") as csvfile:
            writer = csv.writer(csvfile, delimiter=';')
            if write_header:
                writer.writerow(["Date", "Heure", "Dossier date", "Client", "Chantier", "Machine", "Fichier"])
            writer.writerow([date, heure, dossier_date, client, chantier, machine, nom_fichier])

    return redirect(url_for("index"))

# Route permettant de grouper et trier les fichiers à rattacher à une facture
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
        bc_file = request.files.get("bc_file")
        if bc_file and bc_file.filename:
            bc_filename = sanitize_filename(bc_file.filename)
            bc_path = os.path.join(dossier_facture, f"BC_{bc_filename}")
            bc_file.save(bc_path)
        return redirect(url_for("facture"))

    regex_mois = r"^(" + "|".join(MOIS_MAJ_ACCENTS) + r")\d{2}$"
    mois_disponibles = [
        nom for nom in os.listdir(BASE_DIR)
        if os.path.isdir(os.path.join(BASE_DIR, nom))
        and re.match(regex_mois, nom)
        and nom != "Factures"
    ]
    mois_disponibles.sort(key=mois_code_to_tuple)

    mois_selectionne = request.args.get("mois_filtre", "")
    client_selectionne = request.args.get("client_filtre", "")

    clients_set = set()
    if mois_selectionne:
        mois_dir = os.path.join(BASE_DIR, mois_selectionne)
        if os.path.isdir(mois_dir):
            for client in os.listdir(mois_dir):
                if os.path.isdir(os.path.join(mois_dir, client)):
                    clients_set.add(client)
    else:
        for mois in mois_disponibles:
            mois_dir = os.path.join(BASE_DIR, mois)
            for client in os.listdir(mois_dir):
                if os.path.isdir(os.path.join(mois_dir, client)):
                    clients_set.add(client)
    clients_disponibles = sorted(list(clients_set))

    fichiers_groupes = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(list))))
    for mois in mois_disponibles:
        if mois_selectionne and mois != mois_selectionne:
            continue
        mois_dir = os.path.join(BASE_DIR, mois)
        for client in os.listdir(mois_dir):
            if client_selectionne and client != client_selectionne:
                continue
            client_dir = os.path.join(mois_dir, client)
            if not os.path.isdir(client_dir):
                continue
            for chantier in os.listdir(client_dir):
                chantier_dir = os.path.join(client_dir, chantier)
                if not os.path.isdir(chantier_dir):
                    continue
                for machine in os.listdir(chantier_dir):
                    machine_dir = os.path.join(chantier_dir, machine)
                    if not os.path.isdir(machine_dir):
                        continue
                    for file in os.listdir(machine_dir):
                        chemin = os.path.join(machine_dir, file)
                        if file.lower().endswith('.pdf'):
                            fichiers_groupes[mois][client][chantier][machine].append((chemin, file))

    return render_template(
        "facture.html",
        fichiers_groupes=fichiers_groupes,
        mois_disponibles=mois_disponibles,
        clients_disponibles=clients_disponibles,
        mois_selectionne=mois_selectionne,
        client_selectionne=client_selectionne
    )

# Route proxy pour afficher un fichier PDF localement depuis un chemin sécurisé
@app.route('/static/pdf_proxy')
def pdf_proxy():
    path = request.args.get('path', '')
    if not path or not os.path.isfile(path) or not path.lower().endswith('.pdf'):
        return abort(404)
    if not os.path.abspath(path).startswith(os.path.abspath(BASE_DIR)):
        return abort(403)
    return send_file(path, mimetype='application/pdf')

# Lancement de l'application avec interface graphique FlaskUI
if __name__ == "__main__":
    ui = FlaskUI(app=app, server="flask", width=900, height=700)
    ui.run()
