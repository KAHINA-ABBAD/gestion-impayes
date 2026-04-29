#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Générateur de fichier de suivi des impayés clients
Version 1.0 - Conforme à la spécification technique

Ce script gère :
- La création initiale du fichier structuré Excel
- L'actualisation mensuelle avec préservation des saisies manuelles
- L'application des règles de formatage et protection
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule
import pandas as pd
from datetime import datetime
from pathlib import Path
import logging
from typing import Dict, Optional, List

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class GenerateurSuiviImpayes:
    """
    Classe principale pour la génération et l'actualisation du fichier de suivi des impayés.
    """
    
    # Définition de la structure figée du fichier (14 colonnes)
    STRUCTURE_COLONNES = [
        {'nom': 'Clients', 'type': 'auto', 'largeur': 25},
        {'nom': 'Ville', 'type': 'auto', 'largeur': 20},
        {'nom': 'Secteur', 'type': 'auto', 'largeur': 20},
        {'nom': 'Compte Tiers', 'type': 'auto', 'largeur': 15},
        {'nom': 'Solde AU', 'type': 'auto', 'largeur': 15},
        {'nom': 'Évolution', 'type': 'auto', 'largeur': 12},  # NOUVELLE COLONNE
        {'nom': 'NOMBRE DE FACTURES NON REGLEES', 'type': 'auto', 'largeur': 25},
        {'nom': 'Le client a-t-il de la famille ?', 'type': 'manuel', 'largeur': 25},
        {'nom': 'Si oui faut-il contacter la famille ou le client ?', 'type': 'manuel', 'largeur': 35},
        {'nom': 'Qui contacter le Client/famille ?', 'type': 'manuel', 'largeur': 30},
        {'nom': 'Le client est-il en capacité de bien comprendre les informations ?', 'type': 'manuel', 'largeur': 40},
        {'nom': 'Compte clôturé ?', 'type': 'auto', 'largeur': 18},
        {'nom': 'Commentaires RS', 'type': 'manuel', 'largeur': 40},
        {'nom': 'Commentaires compta', 'type': 'manuel', 'largeur': 40}
    ]
    
    # Colonnes automatiques (indices 0-based) - AJOUT colonne 5 (Évolution)
    COLONNES_AUTO = [0, 1, 2, 3, 4, 5, 6, 11]
    
    # Colonnes manuelles (indices 0-based)
    COLONNES_MANUELLES = [7, 8, 9, 10, 12, 13]
    
    # Index de la colonne "Compte Tiers" (clé primaire)
    COL_COMPTE_TIERS = 3
    
    # Index de la colonne "Compte clôturé ?"
    COL_COMPTE_CLOTURE = 10
    
    def __init__(self, dossier_base: str = "."):
        """
        Initialise le générateur avec les chemins de travail.
        
        Args:
            dossier_base: Chemin du dossier racine du projet
        """
        self.dossier_base = Path(dossier_base)
        self.dossier_data = self.dossier_base / "data"
        self.dossier_output = self.dossier_base / "output"
        self.dossier_archives = self.dossier_base / "archives"
        
        # Créer les dossiers s'ils n'existent pas
        for dossier in [self.dossier_data, self.dossier_output, self.dossier_archives]:
            dossier.mkdir(parents=True, exist_ok=True)
    
    def creer_fichier_initial(
        self, 
        fichier_brut: str, 
        fichier_comptes_clotures: Optional[str] = None,
        date_reference: Optional[str] = None
    ) -> str:
        """
        Crée le fichier initial de suivi des impayés.
        
        Args:
            fichier_brut: Chemin vers le fichier d'export comptable (Excel ou CSV)
            fichier_comptes_clotures: Chemin vers le fichier des comptes clôturés (optionnel)
            date_reference: Date de référence au format AAAA-MM-JJ (défaut: aujourd'hui)
        
        Returns:
            Chemin du fichier généré
        """
        logger.info("=== CRÉATION DU FICHIER INITIAL ===")
        
        # Déterminer la date de référence
        if date_reference is None:
            date_reference = datetime.now().strftime("%Y-%m-%d")
        
        # Charger les données sources
        logger.info(f"Chargement du fichier brut: {fichier_brut}")
        df_brut = self._charger_fichier(fichier_brut)
        
        # Valider la structure du fichier brut
        self._valider_structure_brut(df_brut)
        
        # Charger les comptes clôturés si fournis
        comptes_clotures = set()
        if fichier_comptes_clotures:
            logger.info(f"Chargement des comptes clôturés: {fichier_comptes_clotures}")
            comptes_clotures = self._charger_comptes_clotures(fichier_comptes_clotures)
            logger.info(f"Nombre de comptes clôturés: {len(comptes_clotures)}")
        
        # Préparer le DataFrame final
        df_final = self._preparer_dataframe_initial(df_brut, comptes_clotures)
        
        # Créer le workbook Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Suivi Impayés"
        
        # Écrire les données
        self._ecrire_donnees(ws, df_final)
        
        # Appliquer le formatage
        self._appliquer_formatage(ws, len(df_final))
        
        # Créer le tableau structuré
        self._creer_tableau_structure(ws, len(df_final))
        
        # Ajouter les validations de données
        self._ajouter_validations(ws, len(df_final))
        
        # Appliquer la mise en forme conditionnelle
        self._appliquer_mise_en_forme_conditionnelle(ws, len(df_final))
        
        # Protéger les colonnes automatiques
        self._proteger_colonnes(ws)
        
        # Créer l'onglet Notice
        self._creer_onglet_notice(wb)
        
        # Créer l'onglet Paramètres
        self._creer_onglet_parametres(wb)
        
        # Sauvegarder le fichier
        nom_fichier = f"Suivi_Impayes_{date_reference}.xlsx"
        chemin_fichier = self.dossier_output / nom_fichier
        
        logger.info(f"Sauvegarde du fichier: {chemin_fichier}")
        wb.save(chemin_fichier)
        
        logger.info(f"✓ Fichier créé avec succès: {chemin_fichier}")
        logger.info(f"  - Nombre de lignes: {len(df_final)}")
        logger.info(f"  - Colonnes: 14 (dont nouvelle colonne Évolution)")
        
        return str(chemin_fichier)
    
    def actualiser_fichier(
        self,
        fichier_existant: str,
        fichier_brut: str,
        fichier_comptes_clotures: Optional[str] = None,
        date_reference: Optional[str] = None
    ) -> str:
        """
        Actualise un fichier existant avec de nouvelles données tout en préservant les saisies manuelles.
        Crée un dossier par mois et maintient un fichier ACTUEL toujours à jour.
        
        Args:
            fichier_existant: Chemin vers le fichier Excel existant
            fichier_brut: Chemin vers le nouveau fichier d'export comptable
            fichier_comptes_clotures: Chemin vers le fichier des comptes clôturés (optionnel)
            date_reference: Date de référence au format AAAA-MM-JJ (défaut: aujourd'hui)
        
        Returns:
            Chemin du fichier actualisé (fichier ACTUEL)
        """
        logger.info("=== ACTUALISATION DU FICHIER EXISTANT ===")
        
        # Déterminer la date de référence
        if date_reference is None:
            date_reference = datetime.now().strftime("%Y-%m-%d")
        
        # Extraire année et mois
        date_obj = datetime.strptime(date_reference, "%Y-%m-%d")
        annee_mois = date_obj.strftime("%Y-%m")
        dernier_jour_mois = self._dernier_jour_mois(date_obj.year, date_obj.month)
        
        # Archiver l'ancien fichier dans son dossier mensuel
        self._archiver_fichier_par_mois(fichier_existant)
        
        # Charger les données existantes
        logger.info(f"Chargement du fichier existant: {fichier_existant}")
        df_existant = self._charger_fichier_existant(fichier_existant)
        
        # Charger les nouvelles données
        logger.info(f"Chargement du nouveau fichier brut: {fichier_brut}")
        df_nouveau = self._charger_fichier(fichier_brut)
        self._valider_structure_brut(df_nouveau)
        
        # Charger les comptes clôturés
        comptes_clotures = set()
        if fichier_comptes_clotures:
            logger.info(f"Chargement des comptes clôturés: {fichier_comptes_clotures}")
            comptes_clotures = self._charger_comptes_clotures(fichier_comptes_clotures)
            logger.info(f"Nombre de comptes clôturés: {len(comptes_clotures)}")
        
        # Fusionner les données (préservation des saisies manuelles)
        df_final = self._fusionner_donnees(df_existant, df_nouveau, comptes_clotures)
        
        # Créer le nouveau workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Suivi Impayés"
        
        # Écrire les données
        self._ecrire_donnees(ws, df_final)
        
        # Appliquer le formatage
        self._appliquer_formatage(ws, len(df_final))
        
        # Créer le tableau structuré
        self._creer_tableau_structure(ws, len(df_final))
        
        # Ajouter les validations de données
        self._ajouter_validations(ws, len(df_final))
        
        # Appliquer la mise en forme conditionnelle
        self._appliquer_mise_en_forme_conditionnelle(ws, len(df_final))
        
        # Protéger les colonnes automatiques
        self._proteger_colonnes(ws)
        
        # Créer l'onglet Notice
        self._creer_onglet_notice(wb)
        
        # Créer l'onglet Paramètres
        self._creer_onglet_parametres(wb)
        
        # === SAUVEGARDE DOUBLE ===
        
        # 1. Sauvegarder dans le dossier du mois (archive)
        dossier_mois = self.dossier_output / annee_mois
        dossier_mois.mkdir(parents=True, exist_ok=True)
        
        nom_fichier_mois = f"Suivi_Impayes_{dernier_jour_mois}.xlsx"
        chemin_fichier_mois = dossier_mois / nom_fichier_mois
        
        logger.info(f"Sauvegarde archive mensuelle: {chemin_fichier_mois}")
        wb.save(chemin_fichier_mois)
        
        # 2. Sauvegarder le fichier ACTUEL (toujours à jour)
        nom_fichier_actuel = "Suivi_Impayes_ACTUEL.xlsx"
        chemin_fichier_actuel = self.dossier_output / nom_fichier_actuel
        
        logger.info(f"Sauvegarde fichier actuel: {chemin_fichier_actuel}")
        wb.save(chemin_fichier_actuel)
        
        logger.info(f"✓ Fichier actualisé avec succès")
        logger.info(f"  - Archive mensuelle: {chemin_fichier_mois}")
        logger.info(f"  - Fichier actuel: {chemin_fichier_actuel}")
        logger.info(f"  - Lignes: {len(df_final)}")
        logger.info(f"  - Mois: {annee_mois}")
        
        return str(chemin_fichier_actuel)
    
    # ==================== MÉTHODES PRIVÉES ====================
    
    def _charger_fichier(self, chemin_fichier: str) -> pd.DataFrame:
        """
        Charge un fichier Excel ou CSV et retourne un DataFrame.
        """
        chemin = Path(chemin_fichier)
        
        if not chemin.exists():
            raise FileNotFoundError(f"Le fichier {chemin_fichier} n'existe pas")
        
        if chemin.suffix.lower() in ['.xlsx', '.xls']:
            df = pd.read_excel(chemin_fichier)
        elif chemin.suffix.lower() == '.csv':
            df = pd.read_csv(chemin_fichier, encoding='utf-8')
        else:
            raise ValueError(f"Format de fichier non supporté: {chemin.suffix}")
        
        logger.info(f"Fichier chargé: {len(df)} lignes, {len(df.columns)} colonnes")
        return df
    
    def _valider_structure_brut(self, df: pd.DataFrame):
        """
        Valide que le fichier brut contient les colonnes nécessaires.
        """
        colonnes_requises = [
            'Clients', 'Ville', 'Secteur', 'Compte Tiers', 
            'Solde AU', 'NOMBRE DE FACTURES NON REGLEES'
        ]
        
        colonnes_manquantes = [col for col in colonnes_requises if col not in df.columns]
        
        if colonnes_manquantes:
            raise ValueError(f"Colonnes manquantes dans le fichier brut: {colonnes_manquantes}")
        
        logger.info("✓ Structure du fichier brut validée")
    
    def _charger_comptes_clotures(self, chemin_fichier: str) -> set:
        """
        Charge la liste des comptes clôturés depuis un fichier Excel ou CSV.
        Retourne un set de comptes tiers clôturés.
        """
        df = self._charger_fichier(chemin_fichier)
        
        # Chercher la colonne contenant les comptes tiers
        # (peut s'appeler "Compte Tiers", "Compte", "CompteClient", etc.)
        col_compte = None
        for col in df.columns:
            if 'compte' in col.lower() and 'tiers' in col.lower():
                col_compte = col
                break
        
        if col_compte is None:
            # Prendre la première colonne par défaut
            col_compte = df.columns[0]
            logger.warning(f"Colonne 'Compte Tiers' non trouvée, utilisation de '{col_compte}'")
        
        # Convertir en set et nettoyer
        comptes = set()
        for compte in df[col_compte].dropna():
            # Convertir en string et nettoyer
            compte_str = str(compte).strip()
            if compte_str:
                comptes.add(compte_str)
        
        return comptes
    
    def _preparer_dataframe_initial(self, df_brut: pd.DataFrame, comptes_clotures: set) -> pd.DataFrame:
        """
        Prépare le DataFrame initial avec les 14 colonnes dans le bon ordre.
        """
        df = pd.DataFrame()
        
        # Colonnes automatiques (1-5)
        df['Clients'] = df_brut['Clients']
        df['Ville'] = df_brut['Ville']
        df['Secteur'] = df_brut['Secteur']
        df['Compte Tiers'] = df_brut['Compte Tiers'].astype(str)
        df['Solde AU'] = pd.to_numeric(df_brut['Solde AU'], errors='coerce')
        
        # Colonne automatique (6) - Évolution (vide pour création initiale)
        df['Évolution'] = ''
        
        # Colonne automatique (7)
        df['NOMBRE DE FACTURES NON REGLEES'] = pd.to_numeric(df_brut['NOMBRE DE FACTURES NON REGLEES'], errors='coerce').astype('Int64')
        
        # Colonnes manuelles (8-11) - initialement vides
        df['Le client a-t-il de la famille ?'] = ''
        df['Si oui faut-il contacter la famille ou le client ?'] = ''
        df['Qui contacter le Client/famille ?'] = ''
        df['Le client est-il en capacité de bien comprendre les informations ?'] = ''
        
        # Colonne automatique (12) - Compte clôturé
        df['Compte clôturé ?'] = df['Compte Tiers'].apply(
            lambda x: 'OUI' if str(x) in comptes_clotures else ''
        )
        
        # Colonnes manuelles (13-14) - commentaires
        df['Commentaires RS'] = ''
        df['Commentaires compta'] = ''
        
        # FILTRAGE : Supprimer les clients avec 0 factures ou vide
        nb_lignes_avant = len(df)
        df = df[
            (df['NOMBRE DE FACTURES NON REGLEES'].notna()) & 
            (df['NOMBRE DE FACTURES NON REGLEES'] > 0)
        ].copy()
        nb_lignes_apres = len(df)
        
        if nb_lignes_avant > nb_lignes_apres:
            logger.info(f"✓ Clients supprimés (0 factures) : {nb_lignes_avant - nb_lignes_apres}")
        
        return df
    
    def _charger_fichier_existant(self, chemin_fichier: str) -> pd.DataFrame:
        """
        Charge un fichier existant et extrait toutes les données (y compris saisies manuelles).
        """
        wb = load_workbook(chemin_fichier)
        ws = wb["Suivi Impayés"]
        
        # Lire toutes les données
        data = []
        headers = None
        
        for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if row_idx == 1:
                headers = list(row)
            else:
                # Ignorer les lignes complètement vides
                if any(cell is not None and str(cell).strip() for cell in row):
                    data.append(list(row))
        
        df = pd.DataFrame(data, columns=headers)
        
        # S'assurer que Compte Tiers est en string
        df['Compte Tiers'] = df['Compte Tiers'].astype(str)
        
        logger.info(f"Données existantes chargées: {len(df)} lignes")
        return df
    
    def _preserver_valeur(self, valeur):
        """
        Préserve une valeur existante en gérant les cas None, NaN, ou valeurs vides.
        Retourne la valeur si elle existe, sinon une chaîne vide.
        """
        import pandas as pd
        
        # Si la valeur est None ou NaN
        if valeur is None or (isinstance(valeur, float) and pd.isna(valeur)):
            return ''
        
        # Convertir en string et retourner
        return str(valeur)
    
    def _fusionner_donnees(
        self, 
        df_existant: pd.DataFrame, 
        df_nouveau: pd.DataFrame, 
        comptes_clotures: set
    ) -> pd.DataFrame:
        """
        Fusionne les données existantes et nouvelles en préservant les saisies manuelles.
        
        Logique:
        - Pour les comptes existants: mise à jour des colonnes auto, préservation des colonnes manuelles
        - Pour les nouveaux comptes: ajout avec colonnes manuelles vides
        - Pour les comptes disparus: conservation avec marquage visuel (optionnel)
        """
        logger.info("Fusion des données avec préservation des saisies manuelles...")
        
        # Préparer le nouveau DataFrame
        df_nouveau_prep = pd.DataFrame()
        df_nouveau_prep['Clients'] = df_nouveau['Clients']
        df_nouveau_prep['Ville'] = df_nouveau['Ville']
        df_nouveau_prep['Secteur'] = df_nouveau['Secteur']
        df_nouveau_prep['Compte Tiers'] = df_nouveau['Compte Tiers'].astype(str)
        df_nouveau_prep['Solde AU'] = pd.to_numeric(df_nouveau['Solde AU'], errors='coerce')
        df_nouveau_prep['NOMBRE DE FACTURES NON REGLEES'] = pd.to_numeric(df_nouveau['NOMBRE DE FACTURES NON REGLEES'], errors='coerce').astype('Int64')
        
        # Calculer le statut de clôture
        df_nouveau_prep['Compte clôturé ?'] = df_nouveau_prep['Compte Tiers'].apply(
            lambda x: 'OUI' if str(x) in comptes_clotures else ''
        )
        
        # Créer un dictionnaire des données existantes indexé par Compte Tiers
        dict_existant = {}
        for _, row in df_existant.iterrows():
            compte_tiers = str(row['Compte Tiers'])
            dict_existant[compte_tiers] = row
        
        # Construire le DataFrame final
        rows_final = []
        comptes_traites = set()
        
        # Traiter les comptes du nouveau fichier
        for _, row_nouveau in df_nouveau_prep.iterrows():
            compte_tiers = str(row_nouveau['Compte Tiers'])
            comptes_traites.add(compte_tiers)
            
            if compte_tiers in dict_existant:
                # Compte existant: fusionner
                row_existant = dict_existant[compte_tiers]
                
                # CALCULER L'ÉVOLUTION DU SOLDE
                ancien_solde = pd.to_numeric(row_existant.get('Solde AU', 0), errors='coerce')
                nouveau_solde = row_nouveau['Solde AU']
                evolution = ""
                
                if pd.notna(ancien_solde) and pd.notna(nouveau_solde):
                    if nouveau_solde > ancien_solde:
                        # Le solde a augmenté (plus négatif) = mauvais
                        evolution = "↑"
                    elif nouveau_solde < ancien_solde:
                        # Le solde a baissé (moins négatif) = bon
                        evolution = "↓"
                    # Sinon (égal), on laisse vide
                
                row_final = {
                    # Colonnes automatiques: prendre les nouvelles valeurs
                    'Clients': row_nouveau['Clients'],
                    'Ville': row_nouveau['Ville'],
                    'Secteur': row_nouveau['Secteur'],
                    'Compte Tiers': compte_tiers,
                    'Solde AU': row_nouveau['Solde AU'],
                    'Évolution': evolution,  # NOUVELLE COLONNE
                    'NOMBRE DE FACTURES NON REGLEES': row_nouveau['NOMBRE DE FACTURES NON REGLEES'],
                    
                    # Colonnes manuelles: préserver les valeurs existantes (même vides)
                    # Utiliser la valeur existante si elle existe, sinon chaîne vide
                    'Le client a-t-il de la famille ?': self._preserver_valeur(row_existant.get('Le client a-t-il de la famille ?')),
                    'Si oui faut-il contacter la famille ou le client ?': self._preserver_valeur(row_existant.get('Si oui faut-il contacter la famille ou le client ?')),
                    'Qui contacter le Client/famille ?': self._preserver_valeur(row_existant.get('Qui contacter le Client/famille ?')),
                    'Le client est-il en capacité de bien comprendre les informations ?': self._preserver_valeur(row_existant.get('Le client est-il en capacité de bien comprendre les informations ?')),
                    
                    # Colonne automatique
                    'Compte clôturé ?': row_nouveau['Compte clôturé ?'],
                    
                    # Colonnes manuelles: commentaires (TOUJOURS PRÉSERVÉS)
                    'Commentaires RS': self._preserver_valeur(row_existant.get('Commentaires RS')),
                    'Commentaires compta': self._preserver_valeur(row_existant.get('Commentaires compta'))
                }
            else:
                # Nouveau compte: créer avec colonnes manuelles vides
                row_final = {
                    'Clients': row_nouveau['Clients'],
                    'Ville': row_nouveau['Ville'],
                    'Secteur': row_nouveau['Secteur'],
                    'Compte Tiers': compte_tiers,
                    'Solde AU': row_nouveau['Solde AU'],
                    'Évolution': '',  # NOUVELLE COLONNE - vide pour nouveau compte
                    'NOMBRE DE FACTURES NON REGLEES': row_nouveau['NOMBRE DE FACTURES NON REGLEES'],
                    'Le client a-t-il de la famille ?': '',
                    'Si oui faut-il contacter la famille ou le client ?': '',
                    'Qui contacter le Client/famille ?': '',
                    'Le client est-il en capacité de bien comprendre les informations ?': '',
                    'Compte clôturé ?': row_nouveau['Compte clôturé ?'],
                    'Commentaires RS': '',
                    'Commentaires compta': ''
                }
            
            rows_final.append(row_final)
        
        # Traiter les comptes disparus (présents dans existant mais pas dans nouveau)
        comptes_disparus = set(dict_existant.keys()) - comptes_traites
        
        if comptes_disparus:
            logger.info(f"Comptes disparus du nouveau fichier: {len(comptes_disparus)}")
            logger.info("Ces comptes sont conservés pour traçabilité")
            
            for compte_tiers in comptes_disparus:
                row_existant = dict_existant[compte_tiers]
                rows_final.append(row_existant.to_dict())
        
        df_final = pd.DataFrame(rows_final)
        
        # FILTRAGE : Supprimer les clients avec 0 factures ou vide
        nb_lignes_avant = len(df_final)
        df_final = df_final[
            (df_final['NOMBRE DE FACTURES NON REGLEES'].notna()) & 
            (df_final['NOMBRE DE FACTURES NON REGLEES'] > 0)
        ].copy()
        nb_lignes_apres = len(df_final)
        
        if nb_lignes_avant > nb_lignes_apres:
            logger.info(f"✓ Clients supprimés (0 factures) : {nb_lignes_avant - nb_lignes_apres}")
        
        logger.info(f"✓ Fusion terminée: {len(df_final)} lignes")
        logger.info(f"  - Comptes mis à jour: {len(comptes_traites & set(dict_existant.keys()))}")
        logger.info(f"  - Nouveaux comptes: {len(comptes_traites - set(dict_existant.keys()))}")
        logger.info(f"  - Comptes conservés (disparus): {len(comptes_disparus)}")
        
        return df_final
    
    def _ecrire_donnees(self, ws, df: pd.DataFrame):
        """
        Écrit les données du DataFrame dans la feuille Excel.
        """
        # Écrire les en-têtes
        for col_idx, col_info in enumerate(self.STRUCTURE_COLONNES, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = col_info['nom']
        
        # Écrire les données
        for row_idx, row_data in enumerate(df.itertuples(index=False), start=2):
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                
                # Gérer les valeurs None et NaN
                if pd.isna(value):
                    cell.value = ''
                else:
                    cell.value = value
        
        logger.info(f"✓ Données écrites: {len(df)} lignes")
    
    def _appliquer_formatage(self, ws, nb_lignes: int):
        """
        Applique le formatage de base (largeurs, alignements, polices).
        """
        # Définir les largeurs de colonnes
        for col_idx, col_info in enumerate(self.STRUCTURE_COLONNES, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = col_info['largeur']
        
        # Formater l'en-tête
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        for col_idx in range(1, 15):  # 14 colonnes
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
        
        # Formater le corps
        body_font = Font(name='Arial', size=10)
        
        for row_idx in range(2, nb_lignes + 2):
            for col_idx in range(1, 15):  # 14 colonnes
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = body_font
                
                # Alignement spécifique par type de colonne
                if col_idx in [1, 2, 3, 8, 9, 10, 11, 13, 14]:  # Colonnes texte
                    cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                elif col_idx == 4:  # Compte Tiers
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif col_idx == 5:  # Solde AU
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    cell.number_format = '#,##0.00 [$€-fr-FR]'
                elif col_idx in [6, 7, 12]:  # Évolution, Nombre factures, Compte clôturé
                    cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Figer la première ligne
        ws.freeze_panes = 'A2'
        
        # Activer les filtres automatiques sur toutes les colonnes (A-N = 14 colonnes)
        ws.auto_filter.ref = f"A1:N{nb_lignes + 1}"
        
        logger.info("✓ Formatage appliqué")
        logger.info("✓ Filtres automatiques activés")
    
    def _creer_tableau_structure(self, ws, nb_lignes: int):
        """
        Convertit les données en tableau structuré Excel.
        """
        # Définir la plage du tableau (A-N = 14 colonnes)
        ref = f"A1:N{nb_lignes + 1}"
        
        # Créer le tableau
        tab = Table(displayName="Tab_SuiviImpayes", ref=ref)
        
        # Appliquer un style de tableau
        style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        tab.tableStyleInfo = style
        
        ws.add_table(tab)
        
        logger.info("✓ Tableau structuré créé: Tab_SuiviImpayes")
    
    def _ajouter_validations(self, ws, nb_lignes: int):
        """
        Ajoute les validations de données (listes déroulantes).
        """
        # Colonne 8 (H): Le client a-t-il de la famille ?
        dv_col8 = DataValidation(
            type="list",
            formula1='"Oui,Non,À vérifier"',
            allow_blank=True
        )
        dv_col8.error = 'Valeur invalide'
        dv_col8.errorTitle = 'Erreur de saisie'
        ws.add_data_validation(dv_col8)
        dv_col8.add(f'H2:H{nb_lignes + 1}')
        
        # Colonne 9 (I): Si oui faut-il contacter la famille ou le client ?
        dv_col9 = DataValidation(
            type="list",
            formula1='"Client,Famille,Les deux,À définir"',
            allow_blank=True
        )
        dv_col9.error = 'Valeur invalide'
        dv_col9.errorTitle = 'Erreur de saisie'
        ws.add_data_validation(dv_col9)
        dv_col9.add(f'I2:I{nb_lignes + 1}')
        
        # Colonne 11 (K): Le client est-il en capacité de bien comprendre les informations ?
        dv_col11 = DataValidation(
            type="list",
            formula1='"Oui,Non,Partiel,À évaluer"',
            allow_blank=True
        )
        dv_col11.error = 'Valeur invalide'
        dv_col11.errorTitle = 'Erreur de saisie'
        ws.add_data_validation(dv_col11)
        dv_col11.add(f'K2:K{nb_lignes + 1}')
        
        logger.info("✓ Validations de données ajoutées")
    
    def _appliquer_mise_en_forme_conditionnelle(self, ws, nb_lignes: int):
        """
        Applique les règles de mise en forme conditionnelle.
        """
        # Règle 1: Comptes clôturés (colonne L = "OUI")
        fill_rouge = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        font_rouge_cloture = Font(color="CC0000")

        for row_idx in range(2, nb_lignes + 2):
            if ws.cell(row=row_idx, column=12).value == 'OUI':  # Colonne L (12)
                for col_idx in range(1, 15):  # 14 colonnes
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.fill = fill_rouge
                    cell.font = font_rouge_cloture
        
        # Règle 2: Soldes élevés (> 1000 €)
        fill_orange = PatternFill(start_color="FFDAB9", end_color="FFDAB9", fill_type="solid")
        font_orange = Font(color="FF8C00", bold=True)
        
        for row_idx in range(2, nb_lignes + 2):
            solde = ws.cell(row=row_idx, column=5).value  # Colonne E (5)
            if solde and float(solde) > 1000:
                ws.cell(row=row_idx, column=5).fill = fill_orange
                ws.cell(row=row_idx, column=5).font = font_orange
        
        # Règle 3: Cellules vides dans colonnes de saisie
        fill_jaune = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")
        
        for row_idx in range(2, nb_lignes + 2):
            for col_idx in [8, 9, 10, 11]:  # Colonnes H, I, J, K (8-11)
                cell = ws.cell(row=row_idx, column=col_idx)
                if not cell.value or str(cell.value).strip() == '':
                    cell.fill = fill_jaune
        
        # Règle 4: Capacité de compréhension limitée
        fill_bleu = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        
        for row_idx in range(2, nb_lignes + 2):
            capacite = ws.cell(row=row_idx, column=11).value  # Colonne K
            if capacite in ['Non', 'Partiel']:
                ws.cell(row=row_idx, column=11).fill = fill_bleu
        
        # Règle 5: Évolution du solde (NOUVEAU)
        font_rouge = Font(color="FF0000", size=16, bold=True)  # Flèche rouge ↑
        font_vert = Font(color="00B050", size=16, bold=True)   # Flèche verte ↓
        
        for row_idx in range(2, nb_lignes + 2):
            evolution = ws.cell(row=row_idx, column=6).value  # Colonne F (Évolution)
            cell = ws.cell(row=row_idx, column=6)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            
            if evolution == '↑':
                # Solde a augmenté (mauvais) = rouge
                cell.font = font_rouge
            elif evolution == '↓':
                # Solde a baissé (bon) = vert
                cell.font = font_vert
        
        logger.info("✓ Mise en forme conditionnelle appliquée")
    
    def _proteger_colonnes(self, ws):
        """
        Protection désactivée - Toutes les cellules sont modifiables.
        """
        # Déverrouiller toutes les cellules
        for row in ws.iter_rows():
            for cell in row:
                cell.protection = openpyxl.styles.Protection(locked=False)
        
        # NE PAS protéger la feuille - tout est modifiable
        ws.protection.sheet = False
        
        logger.info("✓ Aucune protection - Fichier entièrement modifiable")
    
    def _creer_onglet_notice(self, wb):
        """
        Crée l'onglet Notice avec la documentation utilisateur.
        """
        ws_notice = wb.create_sheet("Notice")
        
        notice_content = [
            ["NOTICE D'UTILISATION", ""],
            ["Outil de suivi des impayés clients", ""],
            ["", ""],
            ["1. STRUCTURE DU FICHIER", ""],
            ["", ""],
            ["Le fichier contient 13 colonnes réparties en deux catégories :", ""],
            ["", ""],
            ["COLONNES AUTOMATIQUES (non modifiables):", ""],
            ["- Clients, Ville, Secteur, Compte Tiers", ""],
            ["- Solde AU, NOMBRE DE FACTURES NON REGLEES", ""],
            ["- Compte clôturé ?", ""],
            ["", ""],
            ["COLONNES MANUELLES (éditables):", ""],
            ["- Le client a-t-il de la famille ?", "Liste déroulante: Oui, Non, À vérifier"],
            ["- Si oui faut-il contacter la famille ou le client ?", "Liste déroulante: Client, Famille, Les deux, À définir"],
            ["- Qui contacter le Client/famille ?", "Texte libre"],
            ["- Le client est-il en capacité de bien comprendre les informations ?", "Liste déroulante: Oui, Non, Partiel, À évaluer"],
            ["- Commentaires RS", "Texte libre pour les Responsables de Secteur"],
            ["- Commentaires compta", "Texte libre pour le service comptabilité"],
            ["", ""],
            ["2. UTILISATION DES FILTRES", ""],
            ["", ""],
            ["Les filtres automatiques sont actifs sur toutes les colonnes.", ""],
            ["Vous pouvez filtrer par secteur, ville, montant de solde, etc.", ""],
            ["", ""],
            ["3. CODES COULEURS", ""],
            ["", ""],
            ["GRIS: Compte clôturé (à régulariser, ne pas relancer)", ""],
            ["ORANGE: Solde > 1000 € (prioritaire)", ""],
            ["JAUNE: Information manquante (à compléter)", ""],
            ["BLEU: Capacité de compréhension limitée (attention particulière)", ""],
            ["", ""],
            ["4. ACTUALISATION MENSUELLE", ""],
            ["", ""],
            ["Le fichier est actualisé chaque mois automatiquement.", ""],
            ["VOS SAISIES DANS LES COLONNES MANUELLES SONT TOUJOURS PRÉSERVÉES.", ""],
            ["", ""],
            ["5. SUPPORT", ""],
            ["", ""],
            ["En cas de problème, contactez le service informatique.", ""],
        ]
        
        for row_idx, row_data in enumerate(notice_content, start=1):
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws_notice.cell(row=row_idx, column=col_idx)
                cell.value = value
                
                # Formater les titres
                if row_idx == 1:
                    cell.font = Font(name='Arial', size=14, bold=True, color="1F4E78")
                elif row_idx == 2:
                    cell.font = Font(name='Arial', size=12, italic=True)
                elif ":" in str(value) and col_idx == 1:
                    cell.font = Font(name='Arial', size=10, bold=True)
        
        ws_notice.column_dimensions['A'].width = 60
        ws_notice.column_dimensions['B'].width = 40
        
        logger.info("✓ Onglet Notice créé")
    
    def _creer_onglet_parametres(self, wb):
        """
        Crée l'onglet Paramètres avec les listes de valeurs.
        """
        ws_params = wb.create_sheet("Paramètres")
        
        params_content = [
            ["Liste: Le client a-t-il de la famille ?", "Liste: Si oui contacter qui ?", "Liste: Capacité de compréhension"],
            ["Oui", "Client", "Oui"],
            ["Non", "Famille", "Non"],
            ["À vérifier", "Les deux", "Partiel"],
            ["", "À définir", "À évaluer"]
        ]
        
        for row_idx, row_data in enumerate(params_content, start=1):
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws_params.cell(row=row_idx, column=col_idx)
                cell.value = value
                
                if row_idx == 1:
                    cell.font = Font(bold=True)
        
        ws_params.column_dimensions['A'].width = 40
        ws_params.column_dimensions['B'].width = 40
        ws_params.column_dimensions['C'].width = 40
        
        logger.info("✓ Onglet Paramètres créé")
    
    def _dernier_jour_mois(self, annee: int, mois: int) -> str:
        """
        Retourne le dernier jour du mois au format AAAA-MM-JJ.
        """
        from calendar import monthrange
        dernier_jour = monthrange(annee, mois)[1]
        return f"{annee:04d}-{mois:02d}-{dernier_jour:02d}"
    
    def _archiver_fichier_par_mois(self, chemin_fichier: str):
        """
        Archive l'ancien fichier dans un dossier correspondant à son mois.
        Ne modifie JAMAIS les archives des mois passés.
        """
        import shutil
        
        chemin_source = Path(chemin_fichier)
        
        if not chemin_source.exists():
            return
        
        # Extraire la date du fichier (depuis le nom ou depuis la date de modification)
        try:
            # Essayer d'extraire depuis le nom du fichier (format: Suivi_Impayes_2024-11-30.xlsx)
            if "Suivi_Impayes" in chemin_source.stem:
                parts = chemin_source.stem.split('_')
                if len(parts) >= 3 and '-' in parts[-1]:
                    date_str = parts[-1]  # 2024-11-30
                    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
                else:
                    # Si pas de date dans le nom, utiliser la date de modification
                    date_obj = datetime.fromtimestamp(chemin_source.stat().st_mtime)
            else:
                date_obj = datetime.fromtimestamp(chemin_source.stat().st_mtime)
        except:
            date_obj = datetime.fromtimestamp(chemin_source.stat().st_mtime)
        
        # Créer le dossier du mois
        annee_mois = date_obj.strftime("%Y-%m")
        dossier_mois = self.dossier_output / annee_mois
        dossier_mois.mkdir(parents=True, exist_ok=True)
        
        # Nom de l'archive avec timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nom_archive = f"{chemin_source.stem}_archive_{timestamp}{chemin_source.suffix}"
        chemin_archive = dossier_mois / nom_archive
        
        shutil.copy2(chemin_source, chemin_archive)
        logger.info(f"✓ Fichier archivé dans {annee_mois}: {chemin_archive}")
    
    def _archiver_fichier(self, chemin_fichier: str):
        """
        Archive l'ancien fichier avant actualisation.
        """
        import shutil
        
        chemin_source = Path(chemin_fichier)
        
        if chemin_source.exists():
            # Extraire la date du nom de fichier ou utiliser la date de modification
            timestamp = datetime.fromtimestamp(chemin_source.stat().st_mtime).strftime("%Y-%m-%d_%H%M%S")
            nom_archive = f"{chemin_source.stem}_archive_{timestamp}{chemin_source.suffix}"
            chemin_archive = self.dossier_archives / nom_archive
            
            shutil.copy2(chemin_source, chemin_archive)
            logger.info(f"✓ Fichier archivé: {chemin_archive}")


# ==================== SCRIPT PRINCIPAL ====================

def main():
    """
    Point d'entrée principal du script.
    """
    import argparse
    
    parser = argparse.ArgumentParser(description="Générateur de fichier de suivi des impayés clients")
    
    parser.add_argument('mode', choices=['creer', 'actualiser'], help="Mode d'opération")
    parser.add_argument('--fichier-brut', required=True, help="Chemin vers le fichier brut (export comptable)")
    parser.add_argument('--fichier-comptes-clotures', help="Chemin vers le fichier des comptes clôturés")
    parser.add_argument('--fichier-existant', help="Chemin vers le fichier existant (mode actualiser uniquement)")
    parser.add_argument('--date', help="Date de référence au format AAAA-MM-JJ (défaut: aujourd'hui)")
    parser.add_argument('--dossier-base', default="/home/claude/suivi_impayes", help="Dossier racine du projet")
    
    args = parser.parse_args()
    
    # Créer le générateur
    generateur = GenerateurSuiviImpayes(dossier_base=args.dossier_base)
    
    try:
        if args.mode == 'creer':
            # Mode création
            chemin_fichier = generateur.creer_fichier_initial(
                fichier_brut=args.fichier_brut,
                fichier_comptes_clotures=args.fichier_comptes_clotures,
                date_reference=args.date
            )
            
        elif args.mode == 'actualiser':
            # Mode actualisation
            if not args.fichier_existant:
                raise ValueError("Le paramètre --fichier-existant est requis en mode actualiser")
            
            chemin_fichier = generateur.actualiser_fichier(
                fichier_existant=args.fichier_existant,
                fichier_brut=args.fichier_brut,
                fichier_comptes_clotures=args.fichier_comptes_clotures,
                date_reference=args.date
            )
        
        print(f"\n{'='*60}")
        print(f"✓ SUCCÈS - Fichier généré: {chemin_fichier}")
        print(f"{'='*60}\n")
        
    except Exception as e:
        logger.error(f"ERREUR: {str(e)}", exc_info=True)
        print(f"\n{'='*60}")
        print(f"ECHEC - {str(e)}")
        print(f"{'='*60}\n")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())
