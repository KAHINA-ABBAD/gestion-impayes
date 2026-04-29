#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Application Web Flask pour l'Outil de Suivi des Impayés
Backend avec API REST
"""

from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
import sys
from pathlib import Path
from datetime import datetime
import tempfile
import shutil
import pandas as pd

# Ajouter le chemin du script principal
sys.path.insert(0, str(Path(__file__).parent))
from generateur_suivi_impayes import GenerateurSuiviImpayes

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB max
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()

@app.route('/logo.png')
def serve_logo():
    logo_path = Path(__file__).parent / 'static' / 'images' / 'Logo_.png'
    if not logo_path.exists():
        logo_path = Path(__file__).parent / 'static' / 'images' / 'Logo.png'
    return send_file(str(logo_path), mimetype='image/png')

# Extensions autorisées
ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'xls'}

def allowed_file(filename):
    """Vérifie si l'extension du fichier est autorisée"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def normaliser_nom_colonne(nom):
    """
    Normalise un nom de colonne en enlevant les espaces et en mettant en minuscules
    """
    return str(nom).strip().lower()


def convertir_format_arche(fichier_path):
    """
    Convertit un fichier au format Arche (10 colonnes) vers le format attendu (6 colonnes)
    Détecte automatiquement si la conversion est nécessaire
    """
    # Lire le fichier
    if str(fichier_path).endswith('.csv'):
        df = pd.read_csv(fichier_path)
    else:
        df = pd.read_excel(fichier_path)
    
    print(f"  → Fichier chargé : {len(df)} lignes, {len(df.columns)} colonnes")
    print(f"  → Colonnes détectées : {list(df.columns)}")
    
    # Normaliser les noms de colonnes pour la détection
    colonnes_normalisees = {normaliser_nom_colonne(col): col for col in df.columns}
    
    # Vérifier si c'est un format Arche (avec Solde et Fact NL)
    a_solde = any(normaliser_nom_colonne(k) == 'solde' for k in df.columns)
    a_fact_nl = any(normaliser_nom_colonne(k) == 'fact nl' for k in df.columns)
    
    # Vérifier si c'est déjà au format standard
    a_solde_au = any(normaliser_nom_colonne(k) == 'solde au' for k in df.columns)
    a_nb_factures = any(normaliser_nom_colonne(k) == 'nombre de factures non reglees' for k in df.columns)
    
    if a_solde and a_fact_nl:
        # C'est un format Arche, on doit le convertir
        print(f"  → Détection du format Arche, conversion en cours...")
        
        # Mapping avec normalisation
        colonnes_mapping = {
            'clients': 'Clients',
            'ville': 'Ville',
            'secteur': 'Secteur',
            'compte tiers': 'Compte Tiers',
            'solde': 'Solde AU',
            'fact nl': 'NOMBRE DE FACTURES NON REGLEES'
        }
        
        # Trouver les colonnes réelles correspondantes
        colonnes_reelles = {}
        colonnes_manquantes = []
        
        for col_norm, col_sortie in colonnes_mapping.items():
            col_trouvee = colonnes_normalisees.get(col_norm)
            if col_trouvee:
                colonnes_reelles[col_trouvee] = col_sortie
            else:
                colonnes_manquantes.append(col_norm)
        
        if colonnes_manquantes:
            raise ValueError(
                f"Colonnes manquantes dans le format Arche : {colonnes_manquantes}. "
                f"Colonnes disponibles : {list(df.columns)}"
            )
        
        # Extraire et renommer
        df_converti = df[list(colonnes_reelles.keys())].copy()
        df_converti.columns = list(colonnes_reelles.values())
        
        # Nettoyer les données - Solde AU
        if df_converti['Solde AU'].dtype == 'object':
            # Remplacer les valeurs vides/invalides par 0
            df_converti['Solde AU'] = df_converti['Solde AU'].astype(str).str.replace(' ', '').str.replace(',', '.')
            df_converti['Solde AU'] = pd.to_numeric(df_converti['Solde AU'], errors='coerce').fillna(0)
        else:
            # Si déjà numérique, remplacer les NaN par 0
            df_converti['Solde AU'] = df_converti['Solde AU'].fillna(0)
        
        # Nettoyer les données - Nombre de factures
        # Remplacer les valeurs vides/invalides par 0
        df_converti['NOMBRE DE FACTURES NON REGLEES'] = pd.to_numeric(
            df_converti['NOMBRE DE FACTURES NON REGLEES'], 
            errors='coerce'
        ).fillna(0).astype(int)
        
        # Enlever les lignes sans Compte Tiers
        df_converti = df_converti.dropna(subset=['Compte Tiers'])
        
        print(f"  → Données nettoyées : {len(df_converti)} lignes valides")
        
        # Sauvegarder le fichier converti
        fichier_converti = Path(fichier_path).parent / f"converti_{Path(fichier_path).name}"
        if str(fichier_path).endswith('.csv'):
            df_converti.to_csv(fichier_converti, index=False)
        else:
            df_converti.to_excel(fichier_converti, index=False)
        
        print(f"  → Conversion réussie : {len(df_converti)} lignes")
        return str(fichier_converti)
    
    # Sinon, vérifier si c'est déjà au bon format
    elif a_solde_au and a_nb_factures:
        print(f"  → Format déjà conforme, aucune conversion nécessaire")
        return str(fichier_path)
    
    else:
        # Format non reconnu - afficher un diagnostic détaillé
        message_erreur = (
            "Format de fichier non reconnu.\n"
            f"Colonnes détectées : {list(df.columns)}\n\n"
            "Le fichier doit contenir soit :\n"
            "1. Format Arche : Clients, Ville, Secteur, Compte Tiers, Solde, Fact NL\n"
            "2. Format Standard : Clients, Ville, Secteur, Compte Tiers, Solde AU, NOMBRE DE FACTURES NON REGLEES\n\n"
            "Vérifiez que les noms de colonnes sont corrects (espaces, accents, majuscules)."
        )
        raise ValueError(message_erreur)


def convertir_comptes_clotures_arche(fichier_path):
    """
    Convertit un fichier de comptes clôturés au format Arche
    """
    # Lire le fichier
    if str(fichier_path).endswith('.csv'):
        df = pd.read_csv(fichier_path)
    else:
        df = pd.read_excel(fichier_path)
    
    # Si le fichier a plus d'une colonne, extraire seulement "Compte Tiers"
    if len(df.columns) > 1:
        if 'Compte Tiers' not in df.columns:
            raise ValueError("La colonne 'Compte Tiers' n'a pas été trouvée dans le fichier des comptes clôturés")
        
        df_converti = df[['Compte Tiers']].copy()
        df_converti = df_converti.dropna(subset=['Compte Tiers'])
        
        # Sauvegarder
        fichier_converti = Path(fichier_path).parent / f"converti_{Path(fichier_path).name}"
        if str(fichier_path).endswith('.csv'):
            df_converti.to_csv(fichier_converti, index=False)
        else:
            df_converti.to_excel(fichier_converti, index=False)
        
        print(f"  → Fichier comptes clôturés converti : {len(df_converti)} comptes")
        return str(fichier_converti)
    
    return str(fichier_path)


@app.route('/')
def index():
    """Page principale"""
    return render_template('index.html')


@app.route('/api/generate', methods=['POST'])
def generate_file():
    """
    Endpoint pour générer le fichier Excel
    """
    try:
        # Récupérer les paramètres
        mode = request.form.get('mode', 'creer')
        date_ref = request.form.get('date', datetime.now().strftime("%Y-%m-%d"))
        
        # Vérifier les fichiers uploadés
        if 'fichier_brut' not in request.files:
            return jsonify({'error': 'Le fichier export comptable est obligatoire'}), 400
        
        fichier_brut = request.files['fichier_brut']
        fichier_clotures = request.files.get('fichier_clotures')
        fichier_existant = request.files.get('fichier_existant')
        
        if fichier_brut.filename == '':
            return jsonify({'error': 'Aucun fichier sélectionné'}), 400
        
        if not allowed_file(fichier_brut.filename):
            return jsonify({'error': 'Format de fichier non autorisé'}), 400
        
        # Sauvegarder les fichiers temporairement
        upload_folder = Path(app.config['UPLOAD_FOLDER'])
        
        brut_path = upload_folder / secure_filename(fichier_brut.filename)
        fichier_brut.save(str(brut_path))
        
        # CONVERSION AUTOMATIQUE DU FORMAT ARCHE
        try:
            brut_path_converti = convertir_format_arche(brut_path)
            # Convertir en Path si c'est une string
            brut_path_converti = Path(brut_path_converti)
        except Exception as e:
            return jsonify({'error': f'Erreur de conversion du fichier export comptable : {str(e)}'}), 400
        
        clotures_path = None
        clotures_path_converti = None
        if fichier_clotures and fichier_clotures.filename != '':
            clotures_path = upload_folder / secure_filename(fichier_clotures.filename)
            fichier_clotures.save(str(clotures_path))
            
            # Conversion des comptes clôturés si nécessaire
            try:
                clotures_path_converti = convertir_comptes_clotures_arche(clotures_path)
                # Convertir en Path si c'est une string
                clotures_path_converti = Path(clotures_path_converti)
            except Exception as e:
                return jsonify({'error': f'Erreur de conversion du fichier comptes clôturés : {str(e)}'}), 400
        
        existant_path = None
        if mode == 'actualiser' and fichier_existant and fichier_existant.filename != '':
            existant_path = upload_folder / secure_filename(fichier_existant.filename)
            fichier_existant.save(str(existant_path))
        
        # Créer le générateur avec un dossier temporaire
        temp_dir = tempfile.mkdtemp()
        generateur = GenerateurSuiviImpayes(dossier_base=temp_dir)
        
        # Générer le fichier
        if mode == 'creer':
            chemin_fichier = generateur.creer_fichier_initial(
                fichier_brut=str(brut_path_converti),
                fichier_comptes_clotures=str(clotures_path_converti) if clotures_path_converti else None,
                date_reference=date_ref
            )
        else:  # actualiser
            if not existant_path:
                return jsonify({'error': 'Le fichier existant est requis en mode actualisation'}), 400
            
            chemin_fichier = generateur.actualiser_fichier(
                fichier_existant=str(existant_path),
                fichier_brut=str(brut_path_converti),
                fichier_comptes_clotures=str(clotures_path_converti) if clotures_path_converti else None,
                date_reference=date_ref
            )
        
        # Nettoyer les fichiers uploadés et convertis
        # Supprimer le fichier original
        if brut_path.exists():
            brut_path.unlink()
        # Supprimer le fichier converti (s'il est différent)
        if brut_path_converti != brut_path and brut_path_converti.exists():
            brut_path_converti.unlink()
        
        if clotures_path and clotures_path.exists():
            clotures_path.unlink()
        if clotures_path_converti and clotures_path_converti != clotures_path and clotures_path_converti.exists():
            clotures_path_converti.unlink()
        
        if existant_path and existant_path.exists():
            existant_path.unlink()
        
        # Vérifier que le fichier généré existe
        if not Path(chemin_fichier).exists():
            raise FileNotFoundError(f"Le fichier généré n'a pas été créé : {chemin_fichier}")
        
        # Vérifier que le fichier a une taille > 0
        taille = Path(chemin_fichier).stat().st_size
        if taille == 0:
            raise ValueError(f"Le fichier généré est vide : {chemin_fichier}")
        
        print(f"  → Fichier généré : {chemin_fichier} ({taille} octets)")

        # Lire en mémoire avant de supprimer le dossier temp (évite WinError 32)
        import io
        nom_fichier = Path(chemin_fichier).name
        with open(chemin_fichier, 'rb') as f:
            file_data = io.BytesIO(f.read())

        shutil.rmtree(temp_dir, ignore_errors=True)

        return send_file(
            file_data,
            as_attachment=True,
            download_name=nom_fichier,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        import traceback
        error_msg = str(e)
        stack_trace = traceback.format_exc()
        print(f"Erreur: {error_msg}")
        print(stack_trace)
        return jsonify({'error': f'Erreur lors de la génération : {error_msg}'}), 500


@app.route('/api/health', methods=['GET'])
def health_check():
    """Endpoint de vérification de l'état du serveur"""
    return jsonify({
        'status': 'ok',
        'version': '1.1',
        'timestamp': datetime.now().isoformat()
    })


if __name__ == '__main__':
    import webbrowser
    import threading

    from waitress import serve
    threading.Timer(1.5, lambda: webbrowser.open("http://localhost:5000")).start()
    serve(app, host='0.0.0.0', port=5000, threads=4)
