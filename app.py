# ============================================================
# CVSmart V11 - VERSION FINALE COMPLÈTE
# Toutes les fonctionnalités intégrées
# ============================================================

import os
import re
import json
import shutil
import warnings
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Optional

warnings.filterwarnings('ignore')

import fitz
import docx
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.pipeline import Pipeline
from gestionnaire_email import GestionnaireEmail, tester_connexion

# ============================================================
# CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="CVSmart - L'intelligence qui trie vos talents",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================================
# 50 DOMAINES AVEC MÉTIERS
# ============================================================
CATEGORIES = {
    'Informatique': {
        'emoji': '💻', 'couleur': '#2B3F8C',
        'metiers': {
            'Développeur': ['développeur', 'programmeur', 'full stack', 'front-end', 'back-end', 'python', 'java', 'javascript', 'php', 'react', 'node'],
            'Administrateur réseau': ['administrateur réseau', 'admin réseau', 'réseau', 'ccna', 'cisco', 'windows server'],
            'DevOps': ['devops', 'docker', 'kubernetes', 'ci/cd', 'jenkins', 'ansible'],
            'Data Scientist': ['data scientist', 'data analyst', 'machine learning', 'big data'],
            'Cybersécurité': ['cybersécurité', 'sécurité informatique', 'pentest', 'ceh'],
            'Chef de projet IT': ['chef de projet informatique', 'scrum master', 'product owner'],
            'Support IT': ['support informatique', 'helpdesk', 'technicien informatique'],
        },
        'mots_cles': ['informatique', 'logiciel', 'application', 'web', 'mobile', 'cloud', 'sql', 'api', 'git', 'linux']
    },
    'Génie Civil': {
        'emoji': '🏗️', 'couleur': '#FF6B00',
        'metiers': {
            'Ingénieur civil': ['ingénieur civil', 'ingénieur génie civil', 'ingénieur structure'],
            'Conducteur de travaux': ['conducteur de travaux', 'chef de chantier'],
            'Topographe': ['topographe', 'géomètre'],
            'Dessinateur projeteur': ['dessinateur', 'projeteur', 'autocad'],
        },
        'mots_cles': ['génie civil', 'chantier', 'béton', 'structure', 'construction', 'plans']
    },
    'Génie Électrique': {
        'emoji': '⚡', 'couleur': '#FFD700',
        'metiers': {
            'Ingénieur électrique': ['ingénieur électrique', 'ingénieur électricité'],
            'Électricien': ['électricien', 'électricité', 'câblage', 'installation électrique'],
            'Technicien électricité': ['technicien électricité', 'technicien électrique'],
            'Automaticien': ['automaticien', 'automate', 'plc', 'scada'],
        },
        'mots_cles': ['électrique', 'électricité', 'courant', 'tension', 'circuit']
    },
    'Génie Mécanique': {
        'emoji': '⚙️', 'couleur': '#8B00FF',
        'metiers': {
            'Ingénieur mécanique': ['ingénieur mécanique'],
            'Technicien mécanique': ['technicien mécanique', 'mécanicien industriel'],
            'Dessinateur industriel': ['dessinateur industriel', 'solidworks', 'catia'],
            'Responsable maintenance': ['responsable maintenance', 'maintenance industrielle'],
        },
        'mots_cles': ['mécanique', 'machine', 'moteur', 'usinage', 'fabrication']
    },
    'Génie des Procédés': {
        'emoji': '🧪', 'couleur': '#E63946',
        'metiers': {
            'Ingénieur procédés': ['ingénieur procédés', 'génie des procédés', 'procédés industriels'],
            'Ingénieur chimiste': ['ingénieur chimiste', 'chimie industrielle'],
            'Responsable production': ['responsable production', 'chef de production'],
            'Technicien procédés': ['technicien procédés', 'opérateur procédés'],
        },
        'mots_cles': ['procédés', 'chimie', 'pétrole', 'raffinage', 'industriel']
    },
    'Mécatronique': {
        'emoji': '🤖', 'couleur': '#00C4CC',
        'metiers': {
            'Ingénieur mécatronique': ['ingénieur mécatronique', 'mécatronique'],
            'Roboticien': ['roboticien', 'robotique', 'robot'],
            'Technicien automatisme': ['automatisme', 'automate', 'plc'],
        },
        'mots_cles': ['mécatronique', 'robot', 'automatisation', 'électronique']
    },
    'Électrotechnique': {
        'emoji': '🔌', 'couleur': '#FFAA00',
        'metiers': {
            'Ingénieur électrotechnique': ['électrotechnique', 'ingénieur électrotechnique'],
            'Technicien électrotechnique': ['technicien électrotechnique'],
        },
        'mots_cles': ['électrotechnique', 'moteur électrique', 'transformateur']
    },
    'Télécommunications': {
        'emoji': '📡', 'couleur': '#7B68EE',
        'metiers': {
            'Ingénieur télécom': ['ingénieur télécom', 'ingénieur télécommunications'],
            'Technicien réseau': ['technicien réseau', 'technicien télécom'],
            'Installateur': ['installateur antenne', 'installation réseau'],
        },
        'mots_cles': ['télécom', 'télécommunications', 'fibre', 'antenne', '4g', '5g']
    },
    'BTP et Construction': {
        'emoji': '🔨', 'couleur': '#A0522D',
        'metiers': {
            'Maçon': ['maçon', 'maçonnerie'],
            'Coffreur': ['coffreur', 'coffrage'],
            'Ferrailleur': ['ferrailleur', 'ferraillage'],
            'Chef de chantier': ['chef de chantier'],
            'Peintre bâtiment': ['peintre bâtiment', 'peinture'],
            'Carreleur': ['carreleur', 'carrelage'],
        },
        'mots_cles': ['btp', 'bâtiment', 'chantier', 'maçonnerie']
    },
    'Énergie et Pétrole': {
        'emoji': '🛢️', 'couleur': '#2F4F4F',
        'metiers': {
            'Ingénieur pétrole': ['ingénieur pétrole', 'ingénieur pétrolier'],
            'Ingénieur énergie': ['ingénieur énergie', 'énergies renouvelables'],
            'Technicien pétrole': ['technicien pétrole', 'technicien forage'],
        },
        'mots_cles': ['pétrole', 'gaz', 'énergie', 'forage', 'raffinage', 'solaire']
    },
    'Comptabilité': {
        'emoji': '💰', 'couleur': '#E63946',
        'metiers': {
            'Comptable': ['comptable', 'comptabilité', 'tenue de comptes'],
            'Auditeur': ['auditeur', 'audit'],
            'Contrôleur de gestion': ['contrôleur de gestion', 'contrôle de gestion'],
            'Expert-comptable': ['expert-comptable', 'expert comptable'],
            'Assistant comptable': ['assistant comptable', 'aide-comptable'],
        },
        'mots_cles': ['comptabilité', 'bilan', 'grand livre', 'sage', 'écritures']
    },
    'Finance': {
        'emoji': '💵', 'couleur': '#00C4CC',
        'metiers': {
            'Analyste financier': ['analyste financier', 'analyse financière'],
            'Directeur financier': ['directeur financier', 'daf'],
            'Contrôleur financier': ['contrôleur financier'],
            'Trésorier': ['trésorier', 'trésorerie'],
        },
        'mots_cles': ['finance', 'budget', 'investissement', 'portefeuille', 'risque']
    },
    'Banque et Assurance': {
        'emoji': '🏦', 'couleur': '#2B3F8C',
        'metiers': {
            'Chargé de clientèle': ['chargé de clientèle', 'conseiller clientèle'],
            'Banquier': ['banquier', 'banque'],
            'Assureur': ['assureur', 'assurance'],
            'Analyste crédit': ['analyste crédit', 'analyse crédit'],
        },
        'mots_cles': ['banque', 'assurance', 'crédit', 'placement', 'microfinance']
    },
    'Ressources Humaines': {
        'emoji': '👥', 'couleur': '#8B00FF',
        'metiers': {
            'Chargé RH': ['chargé rh', 'chargé ressources humaines'],
            'Recruteur': ['recruteur', 'recrutement', 'talent acquisition'],
            'Gestionnaire de paie': ['gestionnaire paie', 'paie', 'payroll'],
            'Responsable formation': ['responsable formation'],
            'DRH': ['drh', 'directeur ressources humaines'],
        },
        'mots_cles': ['rh', 'ressources humaines', 'recrutement', 'paie', 'sirh']
    },
    'Marketing': {
        'emoji': '📢', 'couleur': '#FF6B00',
        'metiers': {
            'Marketing digital': ['marketing digital', 'digital marketing', 'community manager'],
            'Chef de produit': ['chef de produit', 'product manager'],
            'Brand manager': ['brand manager', 'branding', 'marque'],
            'SEO/SEA': ['seo', 'sem', 'référencement'],
            'Chargé de communication': ['chargé de communication', 'communication'],
        },
        'mots_cles': ['marketing', 'communication', 'digital', 'campagne', 'publicité']
    },
    'Commercial et Vente': {
        'emoji': '🤝', 'couleur': '#FF0055',
        'metiers': {
            'Commercial': ['commercial', 'commerciale', 'représentant'],
            'Business Developer': ['business developer', 'business development'],
            'Responsable des ventes': ['responsable des ventes', 'chef des ventes'],
            'Téléconseiller': ['téléconseiller', 'télémarketing'],
        },
        'mots_cles': ['vente', 'commerce', 'négociation', 'client', 'prospection']
    },
    'Logistique et Transport': {
        'emoji': '🚚', 'couleur': '#FFAA00',
        'metiers': {
            'Logisticien': ['logisticien', 'logistique'],
            'Chauffeur': ['chauffeur', 'conducteur', 'permis'],
            'Magasinier': ['magasinier', 'magasin', 'gestion de stock'],
            'Gestionnaire de stock': ['gestionnaire stock', 'gestion des stocks'],
            'Transitaire': ['transitaire', 'transit', 'douane'],
            'Chef de dépôt': ['chef de dépôt', 'responsable dépôt'],
        },
        'mots_cles': ['logistique', 'transport', 'supply chain', 'approvisionnement', 'distribution']
    },
    'Achats et Approvisionnement': {
        'emoji': '📦', 'couleur': '#E63946',
        'metiers': {
            'Acheteur': ['acheteur', 'acheteuse', 'achats'],
            'Responsable achats': ['responsable achats'],
            'Approvisionneur': ['approvisionneur', 'approvisionnement'],
        },
        'mots_cles': ['achat', 'approvisionnement', 'fournisseur', 'commande']
    },
    'Administration': {
        'emoji': '📋', 'couleur': '#666666',
        'metiers': {
            'Secrétaire': ['secrétaire', 'secrétariat'],
            'Assistant de direction': ['assistant de direction', 'assistante de direction'],
            'Office Manager': ['office manager'],
            'Agent administratif': ['agent administratif', 'administration'],
        },
        'mots_cles': ['administration', 'secrétariat', 'assistant', 'planning', 'agenda']
    },
    'Direction Générale': {
        'emoji': '👔', 'couleur': '#2B3F8C',
        'metiers': {
            'Directeur Général': ['directeur général', 'dg'],
            'Directeur Général Adjoint': ['directeur général adjoint', 'dga'],
            'Président': ['président', 'pdg', 'ceo'],
            'Directeur': ['directeur', 'directrice'],
            'Manager': ['manager', 'responsable'],
        },
        'mots_cles': ['direction', 'général', 'président', 'manager', 'responsable']
    },
    'Santé et Médical': {
        'emoji': '⚕️', 'couleur': '#FF0055',
        'metiers': {
            'Médecin': ['médecin', 'docteur', 'doctorat en médecine'],
            'Infirmier': ['infirmier', 'infirmière'],
            'Sage-femme': ['sage-femme', 'sage femme'],
            'Kinésithérapeute': ['kinésithérapeute', 'kiné'],
        },
        'mots_cles': ['santé', 'médical', 'hôpital', 'patient', 'soin']
    },
    'Pharmacie': {
        'emoji': '💊', 'couleur': '#E63946',
        'metiers': {
            'Pharmacien': ['pharmacien', 'pharmacie'],
            'Préparateur en pharmacie': ['préparateur pharmacie'],
        },
        'mots_cles': ['pharmacie', 'médicament', 'pharmaceutique']
    },
    'Paramédical': {
        'emoji': '🩺', 'couleur': '#00C4CC',
        'metiers': {
            'Laborantin': ['laborantin', 'laboratoire'],
            'Radiologue': ['radiologue', 'imagerie'],
            'Technicien laboratoire': ['technicien laboratoire'],
        },
        'mots_cles': ['paramédical', 'laboratoire', 'analyse', 'imagerie']
    },
    'Action Sociale': {
        'emoji': '🤲', 'couleur': '#FF6B00',
        'metiers': {
            'Assistant social': ['assistant social', 'assistante sociale'],
            'Éducateur spécialisé': ['éducateur', 'éducateur spécialisé'],
        },
        'mots_cles': ['social', 'aide', 'humanitaire', 'ong']
    },
    'Vétérinaire': {
        'emoji': '🐕', 'couleur': '#E63946',
        'metiers': {
            'Vétérinaire': ['vétérinaire', 'veterinaire'],
            'Assistant vétérinaire': ['assistant vétérinaire'],
        },
        'mots_cles': ['vétérinaire', 'animaux', 'élevage']
    },
    'Enseignement': {
        'emoji': '🎓', 'couleur': '#2B3F8C',
        'metiers': {
            'Enseignant': ['enseignant', 'professeur', 'instituteur'],
            'Professeur': ['professeur', 'maître de conférences'],
            'Directeur d\'école': ['directeur école', 'proviseur'],
        },
        'mots_cles': ['enseignement', 'école', 'université', 'cours', 'élève']
    },
    'Formation professionnelle': {
        'emoji': '📚', 'couleur': '#8B00FF',
        'metiers': {
            'Formateur': ['formateur', 'formatrice'],
            'Coach': ['coach', 'coaching'],
            'Responsable formation': ['responsable formation'],
        },
        'mots_cles': ['formation', 'formateur', 'coaching', 'pédagogie']
    },
    'Recherche scientifique': {
        'emoji': '🔬', 'couleur': '#00C4CC',
        'metiers': {
            'Chercheur': ['chercheur', 'chercheuse'],
            'Assistant de recherche': ['assistant recherche'],
        },
        'mots_cles': ['recherche', 'scientifique', 'laboratoire', 'publication']
    },
    'Droit et Juridique': {
        'emoji': '⚖️', 'couleur': '#2B3F8C',
        'metiers': {
            'Juriste': ['juriste', 'juridique'],
            'Avocat': ['avocat', 'avocate'],
            'Assistant juridique': ['assistant juridique'],
        },
        'mots_cles': ['droit', 'juridique', 'loi', 'contrat', 'contentieux']
    },
    'Notariat': {
        'emoji': '📜', 'couleur': '#666666',
        'metiers': {
            'Notaire': ['notaire', 'notariat'],
            'Clerc de notaire': ['clerc de notaire'],
        },
        'mots_cles': ['notaire', 'notariat', 'acte']
    },
    'Communication et Journalisme': {
        'emoji': '📰', 'couleur': '#FF6B00',
        'metiers': {
            'Journaliste': ['journaliste', 'journalisme'],
            'Rédacteur': ['rédacteur', 'rédactrice', 'rédaction'],
            'Attaché de presse': ['attaché de presse'],
        },
        'mots_cles': ['journalisme', 'presse', 'média', 'rédaction', 'reportage']
    },
    'Art et Design': {
        'emoji': '🎨', 'couleur': '#FF0055',
        'metiers': {
            'Graphiste': ['graphiste', 'graphisme', 'photoshop', 'illustrator'],
            'Designer': ['designer', 'webdesigner', 'ux', 'ui'],
            'Directeur artistique': ['directeur artistique', 'da'],
        },
        'mots_cles': ['design', 'graphique', 'créatif', 'art', 'illustration']
    },
    'Architecture et Urbanisme': {
        'emoji': '🏛️', 'couleur': '#8B00FF',
        'metiers': {
            'Architecte': ['architecte', 'architecture'],
            'Urbaniste': ['urbaniste', 'urbanisme'],
            'Paysagiste': ['paysagiste'],
        },
        'mots_cles': ['architecture', 'urbanisme', 'plans', 'maquette']
    },
    'Audiovisuel et Cinéma': {
        'emoji': '🎬', 'couleur': '#FFAA00',
        'metiers': {
            'Réalisateur': ['réalisateur', 'réalisatrice'],
            'Monteur': ['monteur', 'montage'],
            'Cadreur': ['cadreur', 'caméraman'],
            'Photographe': ['photographe', 'photographie'],
        },
        'mots_cles': ['audiovisuel', 'cinéma', 'vidéo', 'film', 'tournage']
    },
    'Mode et Textile': {
        'emoji': '👗', 'couleur': '#FF0055',
        'metiers': {
            'Styliste': ['styliste', 'stylisme'],
            'Couturier': ['couturier', 'couturière', 'couture'],
            'Modéliste': ['modéliste'],
        },
        'mots_cles': ['mode', 'textile', 'couture', 'vêtement', 'stylisme']
    },
    'Hôtellerie et Restauration': {
        'emoji': '🍽️', 'couleur': '#FF6B00',
        'metiers': {
            'Cuisinier': ['cuisinier', 'cuisinière', 'chef cuisinier'],
            'Chef de cuisine': ['chef de cuisine'],
            'Serveur': ['serveur', 'serveuse'],
            'Réceptionniste': ['réceptionniste', 'réception'],
            'Barman': ['barman', 'barmaid'],
        },
        'mots_cles': ['hôtel', 'restaurant', 'cuisine', 'service', 'tourisme']
    },
    'Tourisme': {
        'emoji': '✈️', 'couleur': '#00C4CC',
        'metiers': {
            'Guide touristique': ['guide touristique', 'guide'],
            'Agent de voyage': ['agent de voyage', 'agence de voyage'],
        },
        'mots_cles': ['tourisme', 'voyage', 'guide', 'agence']
    },
    'Agriculture': {
        'emoji': '🌾', 'couleur': '#E63946',
        'metiers': {
            'Agronome': ['agronome', 'agronomie'],
            'Agriculteur': ['agriculteur', 'agricultrice', 'agriculture'],
            'Technicien agricole': ['technicien agricole'],
        },
        'mots_cles': ['agriculture', 'culture', 'champ', 'récolte', 'élevage']
    },
    'Agroalimentaire': {
        'emoji': '🍞', 'couleur': '#FFAA00',
        'metiers': {
            'Ingénieur agroalimentaire': ['agroalimentaire', 'ingénieur agro'],
            'Technicien qualité': ['technicien qualité', 'contrôle qualité'],
        },
        'mots_cles': ['agroalimentaire', 'alimentaire', 'transformation', 'qualité']
    },
    'Environnement': {
        'emoji': '🌍', 'couleur': '#E63946',
        'metiers': {
            'Ingénieur environnement': ['ingénieur environnement'],
            'Écologiste': ['écologiste', 'écologie'],
        },
        'mots_cles': ['environnement', 'écologie', 'durable', 'déchets']
    },
    'QHSE': {
        'emoji': '🦺', 'couleur': '#FF6B00',
        'metiers': {
            'Responsable QHSE': ['qhse', 'responsable qhse', 'hse'],
            'Animateur sécurité': ['animateur sécurité', 'sécurité'],
            'Contrôleur qualité': ['contrôleur qualité', 'qualité'],
        },
        'mots_cles': ['qhse', 'hse', 'sécurité', 'hygiène', 'qualité', 'environnement']
    },
    'Artisanat': {
        'emoji': '🔧', 'couleur': '#A0522D',
        'metiers': {
            'Artisan': ['artisan', 'artisanat'],
            'Soudeur': ['soudeur', 'soudure'],
            'Bijoutier': ['bijoutier', 'bijouterie'],
            'Cordonnier': ['cordonnier', 'cordonnerie'],
        },
        'mots_cles': ['artisanat', 'artisan', 'manuel', 'fabrication']
    },
    'Menuiserie et Charpente': {
        'emoji': '🪚', 'couleur': '#A0522D',
        'metiers': {
            'Menuisier': ['menuisier', 'menuiserie', 'bois'],
            'Charpentier': ['charpentier', 'charpente'],
            'Ébéniste': ['ébéniste', 'ébénisterie'],
        },
        'mots_cles': ['menuiserie', 'bois', 'charpente', 'meuble']
    },
    'Plomberie': {
        'emoji': '🔧', 'couleur': '#2B3F8C',
        'metiers': {
            'Plombier': ['plombier', 'plomberie'],
            'Installateur sanitaire': ['installation sanitaire'],
        },
        'mots_cles': ['plomberie', 'tuyau', 'sanitaire', 'eau']
    },
    'Mécanique Auto': {
        'emoji': '🚗', 'couleur': '#FF0055',
        'metiers': {
            'Mécanicien auto': ['mécanicien auto', 'mécanique auto'],
            'Carrossier': ['carrossier', 'carrosserie'],
            'Électricien auto': ['électricien auto'],
            'Vidangeur': ['vidangeur', 'vidange', 'graissage'],
        },
        'mots_cles': ['automobile', 'voiture', 'mécanique', 'moteur', 'garage']
    },
    'Métallurgie et Soudure': {
        'emoji': '⚒️', 'couleur': '#666666',
        'metiers': {
            'Soudeur': ['soudeur', 'soudure', 'soudage'],
            'Chaudronnier': ['chaudronnier', 'chaudronnerie'],
            'Tôlier': ['tôlier', 'tôlerie'],
        },
        'mots_cles': ['métallurgie', 'métal', 'soudure', 'acier']
    },
    'Sécurité': {
        'emoji': '🛡️', 'couleur': '#2B3F8C',
        'metiers': {
            'Agent de sécurité': ['agent de sécurité', 'vigile'],
            'Gardien': ['gardien', 'gardiennage'],
            'Chef de sécurité': ['chef de sécurité', 'responsable sécurité'],
        },
        'mots_cles': ['sécurité', 'gardiennage', 'surveillance', 'protection']
    },
    'Nettoyage et Entretien': {
        'emoji': '🧹', 'couleur': '#00C4CC',
        'metiers': {
            'Agent d\'entretien': ['agent entretien', 'agent de propreté'],
            'Femme de ménage': ['femme de ménage', 'ménage'],
            'Laveur': ['laveur', 'lavage'],
        },
        'mots_cles': ['nettoyage', 'entretien', 'propreté', 'ménage']
    },
    'Immobilier': {
        'emoji': '🏠', 'couleur': '#E63946',
        'metiers': {
            'Agent immobilier': ['agent immobilier', 'immobilier'],
            'Promoteur': ['promoteur', 'promotion immobilière'],
            'Gestionnaire locatif': ['gestionnaire locatif', 'gestion locative'],
        },
        'mots_cles': ['immobilier', 'vente', 'location', 'bien']
    },
    'Sport et Loisirs': {
        'emoji': '⚽', 'couleur': '#FF6B00',
        'metiers': {
            'Éducateur sportif': ['éducateur sportif', 'entraîneur'],
            'Coach sportif': ['coach sportif'],
        },
        'mots_cles': ['sport', 'loisirs', 'animation', 'fitness']
    },
    'Non classé': {'emoji': '❓', 'couleur': '#888888', 'metiers': {}, 'mots_cles': []}
}

DOSSIER_TRI = "CV_classes"
DOSSIER_EXCEL = "exports"
DOSSIER_FILTRES = "CV_filtres"
for d in [DOSSIER_TRI, DOSSIER_EXCEL, DOSSIER_FILTRES]:
    os.makedirs(d, exist_ok=True)

# ============================================================
# CSS
# ============================================================
def inject_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    section[data-testid="stSidebar"] { display: none !important; }
    .stApp { background: linear-gradient(135deg, #f0f4ff 0%, #e0f7ff 50%, #f0fff4 100%); }
    .main .block-container { max-width: 1500px !important; padding-top: 1rem !important;
        padding-left: 2rem !important; padding-right: 2rem !important; }
    .main-title { font-size: 2.2rem; font-weight: 900; color: #2B3F8C; margin-bottom: 0.25rem; letter-spacing: -0.5px; }
    .main-subtitle { color: #2d4a7c; font-size: 1rem; font-weight: 500; margin-bottom: 2rem; }
    .cvsmart-header {
    background: linear-gradient(135deg, #2B3F8C 0%, #3A4FA8 100%);
    border-radius: 14px;
    padding: 1rem 1.8rem;
    box-shadow: 0 8px 24px rgba(43,63,140,0.3);
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: relative;
    overflow: hidden; }
.cvsmart-header::after {
    content: '';
    position: absolute;
    right: -60px;
    top: -60px;
    width: 180px;
    height: 180px;
    background: #E63946;
    transform: rotate(45deg);
    opacity: 0.15; }
    .cvsmart-logo-text { font-size: 1.8rem; font-weight: 900; color: white; }
    .cvsmart-logo-sub { font-size: 0.8rem; color: rgba(255,255,255,0.95); font-weight: 500; }
    .stButton > button { border-radius: 10px; font-weight: 700; font-size: 0.9rem;
        padding: 0.6rem 1.2rem; transition: all 0.2s ease; }
    .stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #E63946 0%, #C62828 100%);
    color: white !important;
    border: none;
    box-shadow: 0 4px 14px rgba(230,57,70,0.4);
}
.stButton > button[kind="primary"]:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 20px rgba(230,57,70,0.5);
}
    .stButton > button[kind="secondary"] {
    background: white;
    color: #2B3F8C !important;
    border: 2px solid #2B3F8C;
    font-weight: 700;
}
.stButton > button[kind="secondary"]:hover {
    background: #2B3F8C;
    color: white !important;
}
    .stat-card { background: white; border-radius: 14px; padding: 1.3rem;
        border: 2px solid #e0e7ff; box-shadow: 0 4px 14px rgba(43,63,14,0,0.08);
        display: flex; justify-content: space-between; align-items: center; transition: all 0.2s ease; }
    .stat-label { color: #1a3a6b; font-size: 0.8rem; font-weight: 700; text-transform: uppercase; }
    .stat-value { font-size: 1.8rem; font-weight: 900; color: #2B3F8C; line-height: 1; }
    .stat-icon { font-size: 2rem; }
    .cat-card { background: white; border-radius: 12px; padding: 0.9rem 1.1rem;
        border-left: 6px solid #2B3F8C; box-shadow: 0 4px 14px rgba(0,0,0,0.08);
        display: flex; align-items: center; gap: 0.8rem; margin-bottom: 0.7rem; }
    .cat-emoji { font-size: 1.8rem; }
    .cat-name { font-weight: 800; color: #2B3F8C; font-size: 0.9rem; }
    .cat-count { color: #2d4a7c; font-size: 0.85rem; font-weight: 600; }
    .skill-tag { display: inline-block; padding: 0.4rem 0.9rem; margin: 0.25rem;
        border-radius: 20px; background: linear-gradient(135deg, #2B3F8C 0%, #3A4FA8 100%);
        color: white; font-size: 0.85rem; font-weight: 700; box-shadow: 0 2px 8px rgba(43,63,140,0.3); }
    .badge-cat { display: inline-block; padding: 0.4rem 1rem; border-radius: 20px;
        font-weight: 800; font-size: 0.8rem; color: white; }
    .badge-metier { display: inline-block; padding: 0.35rem 0.9rem; border-radius: 20px;
        font-weight: 700; font-size: 0.75rem; background: #FF6B00; color: white; margin-left: 0.4rem; }
    .badge-ras { display: inline-block; padding: 0.2rem 0.6rem; border-radius: 6px;
        background: #f0f0f0; color: #888; font-size: 0.85rem; font-weight: 600; font-style: italic; }
    .step-box { text-align: center; padding: 1rem 0.7rem; background: white;
        border-radius: 12px; border: 2px solid #e0e7ff; }
    .step-icon { font-size: 2rem; margin-bottom: 0.4rem; }
    .step-title { font-weight: 800; color: #2B3F8C; font-size: 0.8rem; }
    .step-desc { color: #4a6fa5; font-size: 0.7rem; font-weight: 500; }
    #MainMenu, footer, header { visibility: hidden; }
    p, div, span, label { color: #2B3F8C; }
    .stTextInput > div > div > input, .stTextArea > div > div > textarea {
        background: white !important; color: #2B3F8C !important;
        border: 2px solid #c7d7f0 !important; border-radius: 10px !important;
        font-size: 0.95rem !important; font-weight: 500 !important;
        padding: 0.6rem 0.8rem !important; }
    textarea { background-color: white !important; color: #2B3F8C !important; }
    .filter-container { background: white; border-radius: 14px; padding: 1rem;
        border: 2px solid #e0e7ff; }
    .filter-header { font-size: 1rem; font-weight: 800; color: #2B3F8C;
        padding-bottom: 0.5rem; border-bottom: 2px solid #e0e7ff; margin-bottom: 1rem; }
    .filter-group { margin-bottom: 0.9rem; padding-bottom: 0.9rem;
        border-bottom: 1px solid #f0f0f0; }
    .filter-group-title { font-size: 0.8rem; font-weight: 800; color: #1a3a6b;
        text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 0.4rem; }
    .result-item { background: white; border-radius: 12px; padding: 1.1rem;
        border: 2px solid #e0e7ff; margin-bottom: 0.7rem; }
    .result-item:hover { border-color: #2B3F8C; }
    </style>
    """, unsafe_allow_html=True)

# ============================================================
# CONSTANTES APPRENTISSAGE
# ============================================================
FICHIER_APPRENTISSAGE = "apprentissage.json"
SEUIL_CONFIANCE_DEFAUT = 70

SEUILS_CHAMPS = {
    'nom': 70, 'prenom': 70, 'email': 80, 'telephone': 80,
    'sexe': 80, 'age': 75, 'niveau_etudes': 70, 'ville': 70,
    'domaine': 70, 'metier': 60,
}

# ============================================================
# EXTRACTEUR
# ============================================================
class ExtracteurInfos:
    PRENOMS_FEMININS = {
        'marie', 'sophie', 'julie', 'claire', 'lucie', 'camille',
        'léa', 'chloé', 'céline', 'sandrine', 'nathalie', 'isabelle',
        'caroline', 'émilie', 'fatou', 'aïssatou', 'mariama',
        'aminata', 'kadiatou', 'awa', 'grace', 'espérance',
        'bénédicte', 'charlotte', 'sylvie', 'madeleine', 'jericae',
        'jeanne', 'françoise', 'monique', 'véronique', 'patricia',
        'catherine', 'christine', 'valérie', 'séverine', 'audrey',
        'jennifer', 'jessica', 'sarah', 'laura', 'emma', 'léna',
        'juliette', 'manon', 'élise', 'élodie', 'anaïs', 'océane',
        'jerica', 'ange', 'angele', 'angèle'
    }
    
    PRENOMS_MASCULINS = {
        'jean', 'pierre', 'paul', 'jacques', 'lucas', 'hugo',
        'antoine', 'nicolas', 'sébastien', 'thomas', 'julien',
        'maxime', 'alexandre', 'guillaume', 'françois', 'philippe',
        'laurent', 'moussa', 'ibrahima', 'amadou', 'oumar',
        'abdoulaye', 'mamadou', 'samuel', 'david', 'daniel',
        'emmanuel', 'joseph', 'alphonse', 'bernard', 'charles',
        'étienne', 'franck', 'gaston', 'henri', 'jules', 'kévin',
        'léon', 'michel', 'noël', 'oscar', 'pascal', 'raymond',
        'serge', 'victor', 'richard', 'sacha', 'marc', 'luc',
        'eric', 'éric', 'patrick', 'christophe', 'olivier',
        'vincent', 'benjamin', 'adrien', 'arthur', 'louis'
    }
    
    MOTS_INTERDITS = {
        'EXPÉRIENCES', 'EXPERIENCES', 'EXPÉRIENCE', 'EXPERIENCE',
        'FORMATION', 'FORMATIONS', 'ÉDUCATION', 'EDUCATION',
        'CERTIFICATIONS', 'CERTIFICATION', 'DIPLOMES', 'DIPLÔMES',
        'COMPÉTENCES', 'COMPETENCES', 'COMPÉTENCE', 'COMPETENCE',
        'LANGUES', 'LANGUE', 'INTÉRÊTS', 'INTERETS', 'INTÉRÊT',
        'CENTRES', 'CENTRE', 'LOISIRS', 'PROFIL', 'OBJECTIF',
        'OBJECTIFS', 'CONTACT', 'CONTACTS', 'INFORMATIONS',
        'RÉFÉRENCES', 'REFERENCES', 'PROJETS', 'PROJET',
        'ÉCOLE', 'ECOLE', 'ÉCOLES', 'UNIVERSITÉ', 'UNIVERSITE',
        'LYCÉE', 'LYCEE', 'COLLÈGE', 'COLLEGE', 'INSTITUT',
        'ENTREPRISE', 'ENTREPRISES', 'SOCIÉTÉ', 'SOCIETE',
        'STAGE', 'STAGES', 'EMPLOI', 'EMPLOIS', 'POSTE', 'POSTES',
        'CV', 'CURRICULUM', 'VITAE', 'RESUME', 'RÉSUMÉ',
        'NOM', 'PRENOM', 'PRÉNOM', 'NOMS', 'PRENOMS',
        'TEL', 'TÉL', 'TELEPHONE', 'TÉLÉPHONE', 'MOBILE',
        'EMAIL', 'MAIL', 'ADRESSE', 'VILLE', 'PAYS',
        'JANVIER', 'FÉVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN',
        'JUILLET', 'AOÛT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DÉCEMBRE',
        'COMMUNICATION', 'INNOVATION', 'PRODUIT',
        'ANNÉE', 'PREMIÈRE', 'DEUXIÈME', 'TROISIÈME',
        'TCF', 'TOEIC', 'TOEFL', 'MARKETING', 'DIGITAL'
    }
    
    def __init__(self):
        self.villes = [
            'yaoundé', 'douala', 'bafoussam', 'garoua', 'bamenda',
            'maroua', 'ngaoundéré', 'bertoua', 'ebolowa',
            'kribi', 'limbe', 'buea', 'kumba', 'edea', 'dschang', 'foumban',
            'paris', 'lyon', 'marseille', 'toulouse', 'nice', 'nantes',
            'abidjan', 'dakar', 'lomé', 'cotonou', 'libreville',
            'brazzaville', 'kinshasa', 'casablanca', 'tunis', 'londres',
            'montréal', 'bruxelles'
        ]
    
    def extraire_email(self, texte):
        for p in [r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}',
                  r'[a-zA-Z0-9._%+-]+\s*@\s*[a-zA-Z0-9.-]+\s*\.\s*[a-zA-Z]{2,}']:
            m = re.search(p, texte, re.IGNORECASE)
            if m:
                return re.sub(r'\s+', '', m.group(0)).lower()
        for tld in ['com', 'fr', 'net', 'org', 'edu', 'cm', 'uk', 'de', 'ca']:
            pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9-]+' + tld + r'\b'
            m = re.search(pattern, texte, re.IGNORECASE)
            if m:
                email = m.group(0)
                return (email[:-len(tld)] + '.' + tld).lower()
        return "RAS"
    
    def extraire_tel(self, texte):
        for p in [r'\+237[\s.-]?\d[\s.-]?\d{2}[\s.-]?\d{2}[\s.-]?\d{2}[\s.-]?\d{2}',
                  r'00237[\s.-]?\d[\s.-]?\d{2}[\s.-]?\d{2}[\s.-]?\d{2}[\s.-]?\d{2}',
                  r'\b6[\s.-]?\d{2}[\s.-]?\d{2}[\s.-]?\d{2}[\s.-]?\d{2}\b',
                  r'0[1-9][\s.-]?\d{2}[\s.-]?\d{2}[\s.-]?\d{2}[\s.-]?\d{2}',
                  r'\+33[\s.-]?\d[\s.-]?\d{2}[\s.-]?\d{2}[\s.-]?\d{2}[\s.-]?\d{2}',
                  r'\+\d{1,3}[\s.-]?\(?\d{1,4}\)?[\s.-]?\d{2,4}[\s.-]?\d{2,4}[\s.-]?\d{0,4}',
                  r'\b\d{8,15}\b']:
            for m in re.findall(p, texte):
                tel = re.sub(r'[^\d+]', '', m)
                chiffres = re.sub(r'\D', '', tel)
                if len(chiffres) == 9 and chiffres.startswith('6'):
                    return f"+237 {chiffres[0]} {chiffres[1:3]} {chiffres[3:5]} {chiffres[5:7]} {chiffres[7:9]}"
                if 8 <= len(chiffres) <= 15:
                    return tel
        return "RAS"
    
    def extraire_nom_prenom(self, texte, email="RAS"):
        nom, prenom = "", ""
        for p in [r'(?:^|\n)\s*NOM\s*(?:DE\s*FAMILLE)?\s*[:\-]\s*([A-ZÀ-Ý][a-zà-ÿA-ZÀ-Ý\s\-]{1,30})']:
            m = re.search(p, texte, re.MULTILINE)
            if m:
                nom = m.group(1).strip().split()[0]
                break
        for p in [r'(?:^|\n)\s*PR[EÉ]NOM[S]?\s*[:\-]\s*([A-ZÀ-Ý][a-zà-ÿA-ZÀ-Ý\s\-]{1,30})']:
            m = re.search(p, texte, re.MULTILINE)
            if m:
                prenom = m.group(1).strip().split()[0]
                break
        if nom and prenom:
            return nom, prenom
        
        lignes = [l.strip() for l in texte.split('\n') if l.strip()]
        for ligne in lignes[:3]:
            ligne_clean = re.sub(r'[·•\-/–—|]', ' ', ligne)
            ligne_clean = re.sub(r'\s+', ' ', ligne_clean).strip()
            if len(ligne_clean) < 4 or len(ligne_clean) > 45:
                continue
            mots = re.findall(r'\b[A-Za-zÀ-ÿ]{2,20}\b', ligne_clean)
            if not (2 <= len(mots) <= 3):
                continue
            if any(m.upper() in self.MOTS_INTERDITS for m in mots):
                continue
            if not all(m[0].isupper() for m in mots):
                continue
            if len(mots) == 2:
                prenom = mots[0].capitalize()
                nom = mots[1].upper()
            elif len(mots) == 3:
                if mots[0].isupper() and mots[1].isupper() and mots[2].isupper():
                    nom = mots[0].upper()
                    prenom = f"{mots[1].capitalize()} {mots[2].capitalize()}"
                else:
                    prenom = f"{mots[0].capitalize()} {mots[1].capitalize()}"
                    nom = mots[2].upper()
            return nom, prenom
        
        if email and email != "RAS" and '@' in email:
            partie = re.sub(r'\d+', '', email.split('@')[0])
            for sep in ['.', '_', '-']:
                if sep in partie:
                    parts = [p for p in partie.split(sep) if p and p.upper() not in self.MOTS_INTERDITS and len(p) > 1]
                    if len(parts) >= 2:
                        return parts[1].upper(), parts[0].capitalize()
        return "RAS", "RAS"
    
    def extraire_niveau(self, texte):
        t = texte.lower()
        meilleur, pts = "RAS", 0
        for m in re.finditer(r'\bbac\s*\+\s*(\d+)\b', t):
            n = int(m.group(1))
            if n >= 8: niveau, p = 'Doctorat', 30
            elif n >= 5: niveau, p = 'Master/Ingénieur', 25
            elif n >= 3: niveau, p = 'Licence', 15
            elif n >= 2: niveau, p = 'BTS/DUT', 10
            else: niveau, p = 'Bac', 8
            if p > pts: pts, meilleur = p, niveau
        patterns = {
            r'\b(doctorat|phd|thèse|these)\b': ('Doctorat', 30),
            r'\b(master|mastère|msc|mba|dess|dea|ingénieur)\b': ('Master/Ingénieur', 25),
            r'\b(licence|bachelor|licencié)\b': ('Licence', 15),
            r'\b(bts|dut|deug|hnd)\b': ('BTS/DUT', 10),
            r'\b(baccalauréat|bac|gce)\b': ('Bac', 8),
        }
        for p, (niv, points) in patterns.items():
            if re.search(p, t, re.IGNORECASE) and points > pts:
                pts, meilleur = points, niv
        return meilleur, pts
    
    def extraire_experience(self, texte):
        """
        Calcule les années d'expérience :
        1. "X ans d'expérience" → direct
        2. Plage de dates (2018-2023) → 5 ans
        3. "Depuis 2018" → 2026 - 2018
        4. Sinon → 0
        """
        annee_actuelle = datetime.now().year
        durees = []
        
        # MÉTHODE 1 : "X ans d'expérience" (direct)
        for m in re.finditer(r'(\d{1,2})\s*(?:ans?|années?)\s*(?:d\'?\s*)?(?:expérience|exp)', texte, re.IGNORECASE):
            try:
                v = int(m.group(1))
                if 0 <= v <= 50:
                    durees.append(v)
            except: pass
        
        for m in re.finditer(r'expérience\s*(?:de|d\')\s*(\d{1,2})\s*(?:ans?|années?)', texte, re.IGNORECASE):
            try:
                v = int(m.group(1))
                if 0 <= v <= 50:
                    durees.append(v)
            except: pass
        
        # MÉTHODE 2 : Plages de dates "2018-2023" ou "2018 - 2023" ou "2018 à 2023"
        periodes = re.findall(
            r'(19\d{2}|20\d{2})\s*[-–—à]+\s*(19\d{2}|20\d{2}|aujourd\'hui|présent|present|maintenant)',
            texte, re.IGNORECASE
        )
        for debut, fin in periodes:
            try:
                d = int(debut)
                if fin.lower() in ["aujourd'hui", 'présent', 'present', 'maintenant']:
                    f = annee_actuelle
                else:
                    f = int(fin)
                if 1950 <= d < f <= annee_actuelle + 1:
                    durees.append(f - d)
            except: pass
        
        # MÉTHODE 3 : "Depuis 2018" ou "Depuis 2018 à ce jour"
        for m in re.finditer(r'depuis\s+(19\d{2}|20\d{2})', texte, re.IGNORECASE):
            try:
                d = int(m.group(1))
                if 1950 <= d <= annee_actuelle:
                    durees.append(annee_actuelle - d)
            except: pass
        
        # MÉTHODE 4 : "De 2015 à 2020"
        for m in re.finditer(r'de\s+(19\d{2}|20\d{2})\s+(?:à|a)\s+(19\d{2}|20\d{2})', texte, re.IGNORECASE):
            try:
                d = int(m.group(1))
                f = int(m.group(2))
                if 1950 <= d < f <= annee_actuelle + 1:
                    durees.append(f - d)
            except: pass
        
        # Si rien trouvé → 0
        if not durees:
            return 0
        
        # Filtrer les valeurs aberrantes (> 50 ans)
        durees = [d for d in durees if 0 <= d <= 50]
        if not durees:
            return 0
        
        # Retourner la MÉDIANE (plus fiable que la moyenne)
        return int(np.median(durees))
    
    def extraire_age(self, texte):
        """Extrait l'âge : par année de naissance sinon RAS"""
        annee_actuelle = datetime.now().year
        ages_trouves = []
        
        # MÉTHODE 1 : "Né(e) le 15/03/1995" ou "Né le 15-03-1995"
        for m in re.finditer(r'(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})', texte):
            try:
                annee_naiss = int(m.group(3))
                if 1940 <= annee_naiss <= annee_actuelle:
                    age = annee_actuelle - annee_naiss
                    if 15 <= age <= 85:
                        ages_trouves.append((age, 0.98))
            except: pass
        
        # MÉTHODE 2 : "Né(e) en 1995" ou "Née en 1995"
        for m in re.finditer(r'n[ée]e?\s+(?:le\s+\d{1,2}[/\-.]\d{1,2}[/\-.]|en\s+)?(19\d{2}|20\d{2})', texte, re.IGNORECASE):
            try:
                annee_naiss = int(m.group(1))
                if 1940 <= annee_naiss <= annee_actuelle:
                    age = annee_actuelle - annee_naiss
                    if 15 <= age <= 85:
                        ages_trouves.append((age, 0.95))
            except: pass
        
        # MÉTHODE 3 : "Date de naissance : 1995"
        for m in re.finditer(r'date\s+de\s+naissance\s*[:\-]?\s*\d{0,2}[/\-.]?\d{0,2}[/\-.]?(19\d{2}|20\d{2})', texte, re.IGNORECASE):
            try:
                annee_naiss = int(m.group(1))
                if 1940 <= annee_naiss <= annee_actuelle:
                    age = annee_actuelle - annee_naiss
                    if 15 <= age <= 85:
                        ages_trouves.append((age, 0.97))
            except: pass
        
        # MÉTHODE 4 : "J'ai 28 ans" ou "28 ans"
        for m in re.finditer(r"j['']?ai\s+(\d{1,2})\s*ans", texte, re.IGNORECASE):
            try:
                age = int(m.group(1))
                if 15 <= age <= 85:
                    ages_trouves.append((age, 0.90))
            except: pass
        
        # MÉTHODE 5 : "Âge : 28" ou "Age : 28"
        for m in re.finditer(r'[âa]ge\s*[:\-]\s*(\d{1,2})', texte, re.IGNORECASE):
            try:
                age = int(m.group(1))
                if 15 <= age <= 85:
                    ages_trouves.append((age, 0.95))
            except: pass
        
        # MÉTHODE 6 : "(28 ans)" entre parenthèses
        for m in re.finditer(r'\((\d{1,2})\s*ans\)', texte, re.IGNORECASE):
            try:
                age = int(m.group(1))
                if 15 <= age <= 85:
                    ages_trouves.append((age, 0.85))
            except: pass
        
        # Si rien trouvé → RAS (ne PAS inventer)
        if not ages_trouves:
            return "RAS"
        
        # Prendre la valeur avec la meilleure confiance
        ages_trouves.sort(key=lambda x: x[1], reverse=True)
        return str(ages_trouves[0][0])
    
    def extraire_sexe(self, texte, prenom="RAS"):
        """Détermine le sexe : civilité > prénom connu > RAS"""
        t = texte.lower()
        
        # PRIORITÉ 1 : Civilité explicite (99% confiance)
        if re.search(r'\b(madame|mme|mlle|mademoiselle)\b', t):
            return "Femme"
        if re.search(r'\b(monsieur|mr)\b', t):
            return "Homme"
        
        # PRIORITÉ 2 : Prénom connu
        if prenom and prenom != "RAS":
            for p in prenom.lower().split():
                p = p.strip()
                if p in self.PRENOMS_FEMININS:
                    return "Femme"
                if p in self.PRENOMS_MASCULINS:
                    return "Homme"
        
        # PRIORITÉ 3 : Chercher d'autres indices
        # "épouse", "mariée", "mère de"
        if re.search(r'\b(épouse|mari[ée]e|m[èe]re\s+de|elle)\b', t):
            return "Femme"
        if re.search(r'\b(époux|mari[ée]|p[èe]re\s+de|il\s+est)\b', t):
            return "Homme"
        
        # Sinon → RAS (prénom ambigu)
        return "RAS"
    
    def extraire_situation(self, texte):
        t = texte.lower()
        if re.search(r'\b(mari[ée]e?|mariage)\b', t): return "Marié(e)"
        if re.search(r'\b(c[ée]libataire|c[ée]libat)\b', t): return "Célibataire"
        if re.search(r'\b(divorc[ée]e?)\b', t): return "Divorcé(e)"
        if re.search(r'\b(veuf|veuve)\b', t): return "Veuf(ve)"
        return "RAS"
    
    def extraire_nationalite(self, texte):
        t = texte.lower()
        for cle, val in [('camerounais', 'Camerounaise'), ('camerounaise', 'Camerounaise'),
                         ('cameroun', 'Camerounaise'), ('français', 'Française'),
                         ('française', 'Française'), ('sénégalais', 'Sénégalaise'),
                         ('sénégalaise', 'Sénégalaise'), ('ivoirien', 'Ivoirienne'),
                         ('ivoirienne', 'Ivoirienne'), ('gabonais', 'Gabonaise'),
                         ('gabonaise', 'Gabonaise'), ('congolais', 'Congolaise'),
                         ('congolaise', 'Congolaise'), ('tchadien', 'Tchadienne'),
                         ('tchadienne', 'Tchadienne')]:
            if re.search(r'\b' + re.escape(cle) + r'\b', t):
                return val
        return "RAS"
    
    def extraire_langues(self, texte):
        t = texte.lower()
        trouvees = []
        for cle, nom in [('français', 'Français'), ('francais', 'Français'), ('french', 'Français'),
                         ('anglais', 'Anglais'), ('english', 'Anglais'),
                         ('espagnol', 'Espagnol'), ('spanish', 'Espagnol'),
                         ('allemand', 'Allemand'), ('italien', 'Italien'),
                         ('chinois', 'Chinois'), ('arabe', 'Arabe'),
                         ('portugais', 'Portugais'), ('russe', 'Russe')]:
            if re.search(r'\b' + re.escape(cle) + r'\b', t):
                if nom not in trouvees: trouvees.append(nom)
        return ', '.join(trouvees) if trouvees else "RAS"
    
    def extraire_competences(self, texte):
        t = texte.lower()
        trouvees = set()
        for data in CATEGORIES.values():
            for mot in data.get('mots_cles', []):
                if re.search(r'\b' + re.escape(mot) + r'\b', t):
                    trouvees.add(mot.capitalize())
        return sorted(trouvees)[:20]
    
    def extraire_certifications(self, texte):
        t = texte.lower()
        trouvees = []
        for c in ['pmp', 'prince2', 'itil', 'scrum master', 'aws', 'azure',
                  'google cloud', 'cisco', 'ccna', 'ccnp', 'microsoft',
                  'oracle', 'sap', 'comptia', 'ceh', 'google analytics',
                  'hubspot', 'toeic', 'toefl', 'ielts', 'tcf', 'delf',
                  'cpa', 'acca', 'cfa', 'cia', 'cisa', 'excel']:
            if re.search(r'\b' + re.escape(c) + r'\b', t):
                trouvees.append(c.upper())
        return ', '.join(trouvees) if trouvees else "RAS"
    
    def extraire_etablissement(self, texte):
        etabs = []
        for p in [r'universit[ée]\s+(?:de\s+|d\')?[A-ZÀ-Ý][a-zà-ÿ]+(?:\s+[A-ZÀ-Ý][a-zà-ÿ]+)?',
                  r'[ée]cole\s+(?:de\s+|des\s+|nationale\s+|sup[ée]rieure\s+)?[A-ZÀ-Ý][a-zà-ÿ]+(?:\s+[A-ZÀ-Ý][a-zà-ÿ]+)?',
                  r'institut\s+[A-ZÀ-Ý][a-zà-ÿ]+(?:\s+[A-ZÀ-Ý][a-zà-ÿ]+)?',
                  r'[A-ZÀ-Ý][a-zà-ÿ]+\s+College', r'Digital\s+College']:
            for m in re.finditer(p, texte):
                etab = m.group(0).strip()
                if etab not in etabs and 5 < len(etab) < 60:
                    etabs.append(etab)
        return ', '.join(etabs[:3]) if etabs else "RAS"
    
    def extraire_poste_actuel(self, texte):
        t = texte.lower()
        for poste in ['directeur', 'directrice', 'manager', 'responsable',
                      'chef de projet', 'ingénieur', 'consultant', 'analyste',
                      'assistant', 'assistante', 'technicien', 'comptable',
                      'développeur', 'commercial', 'secrétaire', 'stagiaire']:
            if re.search(r'\b' + re.escape(poste) + r'\b', t):
                return poste.capitalize()
        return "RAS"
    
    def extraire_ville(self, texte):
        t = texte.lower()
        for v in self.villes:
            if re.search(r'\b' + re.escape(v) + r'\b', t):
                return v.capitalize()
        return "RAS"
    
    def extraire_tout(self, texte):
        email = self.extraire_email(texte)
        nom, prenom = self.extraire_nom_prenom(texte, email)
        niveau, pts = self.extraire_niveau(texte)
        return {
            'nom': nom or "RAS", 'prenom': prenom or "RAS",
            'email': email, 'telephone': self.extraire_tel(texte),
            'age': self.extraire_age(texte),
            'sexe': self.extraire_sexe(texte, prenom),
            'situation': self.extraire_situation(texte),
            'nationalite': self.extraire_nationalite(texte),
            'niveau_etudes': niveau, 'points_niveau': pts,
            'etablissement': self.extraire_etablissement(texte),
            'annees_experience': self.extraire_experience(texte),
            'poste_actuel': self.extraire_poste_actuel(texte),
            'competences': self.extraire_competences(texte),
            'langues': self.extraire_langues(texte),
            'certifications': self.extraire_certifications(texte),
            'ville': self.extraire_ville(texte),
            'texte_brut': texte
        }

# ============================================================
# MODULE : ÂGE CALCULÉ + SEXE CONFIANCE
# ============================================================
def calculer_age_depuis_annee(texte):
    annee_actuelle = datetime.now().year
    ages_candidats = []
    for m in re.finditer(r'n[ée]e?\s*(?:le|en)?\s*(19\d{2}|20\d{2})', texte, re.IGNORECASE):
        try:
            age = annee_actuelle - int(m.group(1))
            if 16 <= age <= 80: ages_candidats.append((age, 0.95))
        except: pass
    for m in re.finditer(r'(\d{2})[/\-.](\d{2})[/\-.](\d{4})', texte):
        try:
            annee_naiss = int(m.group(3))
            if 1950 <= annee_naiss <= annee_actuelle:
                age = annee_actuelle - annee_naiss
                if 16 <= age <= 80: ages_candidats.append((age, 0.98))
        except: pass
    for m in re.finditer(r"j['']?ai\s*(\d{1,2})\s*ans", texte, re.IGNORECASE):
        try:
            age = int(m.group(1))
            if 16 <= age <= 80: ages_candidats.append((age, 0.90))
        except: pass
    for m in re.finditer(r'[âa]ge\s*[:\-]\s*(\d{1,2})', texte, re.IGNORECASE):
        try:
            age = int(m.group(1))
            if 16 <= age <= 80: ages_candidats.append((age, 0.95))
        except: pass
    if not ages_candidats: return "RAS", 0.0
    ages_candidats.sort(key=lambda x: x[1], reverse=True)
    return str(ages_candidats[0][0]), ages_candidats[0][1]


def determiner_sexe_avec_confiance(texte, prenom="RAS"):
    t = texte.lower()
    if re.search(r'\b(madame|mme|mlle|mademoiselle)\b', t): return "Femme", 0.99
    if re.search(r'\b(monsieur|mr)\b', t): return "Homme", 0.99
    if re.search(r'\b(elle|mari[ée]e|épouse|mère\s+de)\b', t): return "Femme", 0.85
    if re.search(r'\b(il|mari[ée]|époux|p[èe]re\s+de)\b', t): return "Homme", 0.85
    if prenom and prenom != "RAS":
        for p in prenom.lower().split():
            if p in ExtracteurInfos.PRENOMS_FEMININS: return "Femme", 0.90
            if p in ExtracteurInfos.PRENOMS_MASCULINS: return "Homme", 0.90
    return "RAS", 0.0


def calculer_confiance_champ(valeur, champ, texte=""):
    if not valeur or valeur == "RAS": return 0.0
    if champ == 'email': return 0.98 if '@' in valeur and '.' in valeur else 0.5
    if champ == 'telephone':
        chiffres = re.sub(r'\D', '', valeur)
        return 0.95 if 8 <= len(chiffres) <= 15 else 0.5
    if champ in ['nom', 'prenom']: return 0.85 if len(valeur) >= 2 else 0.4
    if champ == 'niveau_etudes':
        return 0.90 if valeur in ['Bac', 'BTS/DUT', 'Licence', 'Master/Ingénieur', 'Doctorat'] else 0.4
    if champ == 'ville': return 0.85
    if champ == 'age': return 0.90
    if champ == 'sexe': return 0.85
    return 0.5


def verifier_champs_a_valider(infos, domaine, metier, seuil):
    champs_a_valider = []
    seuils = {
        'nom': seuil, 'prenom': seuil, 'email': 80, 'telephone': 80,
        'sexe': 80, 'age': 75, 'niveau_etudes': seuil, 'ville': seuil,
    }
    for champ, s in seuils.items():
        valeur = infos.get(champ, 'RAS')
        confiance = calculer_confiance_champ(valeur, champ)
        if confiance * 100 < s:
            champs_a_valider.append({
                'champ': champ, 'valeur_actuelle': valeur,
                'confiance': confiance, 'seuil': s
            })
    return champs_a_valider

# ============================================================
# MODULE : APPRENTISSAGE
# ============================================================
def charger_apprentissage():
    if os.path.exists(FICHIER_APPRENTISSAGE):
        try:
            with open(FICHIER_APPRENTISSAGE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except: pass
    return {'sexe': {}, 'domaines': {}, 'metiers': {},
            'stats': {'total_corrections': 0, 'derniere_maj': None}}


def sauvegarder_apprentissage(app):
    try:
        app['stats']['derniere_maj'] = datetime.now().isoformat()
        with open(FICHIER_APPRENTISSAGE, 'w', encoding='utf-8') as f:
            json.dump(app, f, indent=2, ensure_ascii=False)
        return True
    except: return False


def appliquer_apprentissage(infos, app):
    if infos.get('sexe') == "RAS":
        prenom = infos.get('prenom', '').lower()
        if prenom in app.get('sexe', {}):
            infos['sexe'] = app['sexe'][prenom]
    return infos


def enregistrer_correction(app, champ, ancienne, nouvelle):
    if not nouvelle or nouvelle == ancienne: return app
    if champ == 'sexe' and ancienne != 'RAS':
        prenom = ancienne.lower().strip()
        if prenom:
            app['sexe'][prenom] = nouvelle
    app['stats']['total_corrections'] = app['stats'].get('total_corrections', 0) + 1
    sauvegarder_apprentissage(app)
    return app


def get_stats_apprentissage(app):
    return {
        'sexe_appris': len(app.get('sexe', {})),
        'domaines_appris': len(app.get('domaines', {})),
        'metiers_appris': len(app.get('metiers', {})),
        'total_corrections': app.get('stats', {}).get('total_corrections', 0),
        'derniere_maj': app.get('stats', {}).get('derniere_maj', 'Jamais')
    }

# ============================================================
# CLASSIFICATEUR INTELLIGENT
# ============================================================
class Classificateur:
    def __init__(self):
        self.pipeline = None
    
    def entrainer(self):
        textes, labels = [], []
        for cat, data in CATEGORIES.items():
            if cat == 'Non classé': continue
            for mot in data.get('mots_cles', []):
                textes.append(f"expérience {mot} compétences {mot} maîtrise {mot}")
                textes.append(f"spécialiste {mot} expert {mot}")
                labels.extend([cat, cat])
            for metier, mots in data.get('metiers', {}).items():
                for mot in mots:
                    textes.append(f"{mot} {cat} {metier}")
                    labels.append(cat)
        self.pipeline = Pipeline([
            ('tfidf', TfidfVectorizer(max_features=15000, stop_words=None,
                                      ngram_range=(1, 2), min_df=1, sublinear_tf=True)),
            ('clf', MultinomialNB(alpha=0.05))
        ])
        self.pipeline.fit(textes, labels)
    
    def classifier(self, texte):
        if not self.pipeline: self.entrainer()
        t = texte.lower()
        if len(t) < 20: return 'Non classé', 'RAS', 0.0
        
        # 1. Chercher LE POSTE dans les 15 premières lignes
        lignes = [l.strip() for l in texte.split('\n') if l.strip()]
        zone_titre = " ".join(lignes[:15]).lower()
        
        METIERS_PRIORITAIRES = {
            'logisticien': ('Logistique et Transport', 'Logisticien'),
            'magasinier': ('Logistique et Transport', 'Magasinier'),
            'gestionnaire de stock': ('Logistique et Transport', 'Gestionnaire de stock'),
            'transitaire': ('Logistique et Transport', 'Transitaire'),
            'chauffeur': ('Logistique et Transport', 'Chauffeur'),
            'chef de dépôt': ('Logistique et Transport', 'Chef de dépôt'),
            'directeur général adjoint': ('Direction Générale', 'Directeur Général Adjoint'),
            'directeur général': ('Direction Générale', 'Directeur Général'),
            'dga': ('Direction Générale', 'Directeur Général Adjoint'),
            'dg': ('Direction Générale', 'Directeur Général'),
            'développeur': ('Informatique', 'Développeur'),
            'administrateur réseau': ('Informatique', 'Administrateur réseau'),
            'devops': ('Informatique', 'DevOps'),
            'data scientist': ('Informatique', 'Data Scientist'),
            'comptable': ('Comptabilité', 'Comptable'),
            'auditeur': ('Comptabilité', 'Auditeur'),
            'expert-comptable': ('Comptabilité', 'Expert-comptable'),
            'analyste financier': ('Finance', 'Analyste financier'),
            'ressources humaines': ('Ressources Humaines', 'Chargé RH'),
            'recruteur': ('Ressources Humaines', 'Recruteur'),
            'drh': ('Ressources Humaines', 'DRH'),
            'marketing digital': ('Marketing', 'Marketing digital'),
            'community manager': ('Marketing', 'Marketing digital'),
            'commercial': ('Commercial et Vente', 'Commercial'),
            'médecin': ('Santé et Médical', 'Médecin'),
            'infirmier': ('Santé et Médical', 'Infirmier'),
            'enseignant': ('Enseignement', 'Enseignant'),
            'formateur': ('Formation professionnelle', 'Formateur'),
            'maçon': ('BTP et Construction', 'Maçon'),
            'ingénieur': ('Génie Civil', 'Ingénieur civil'),
            'qhse': ('QHSE', 'Responsable QHSE'),
            'mécanicien': ('Mécanique Auto', 'Mécanicien auto'),
            'vidangeur': ('Mécanique Auto', 'Vidangeur'),
            'agent de sécurité': ('Sécurité', 'Agent de sécurité'),
            'cuisinier': ('Hôtellerie et Restauration', 'Cuisinier'),
            'serveur': ('Hôtellerie et Restauration', 'Serveur'),
            'secrétaire': ('Administration', 'Secrétaire'),
            'juriste': ('Droit et Juridique', 'Juriste'),
            'architecte': ('Architecture et Urbanisme', 'Architecte'),
        }
        
        metiers_trouves = []
        for mot_cle, (domaine, metier) in METIERS_PRIORITAIRES.items():
            if mot_cle in zone_titre:
                pos = zone_titre.find(mot_cle)
                contexte = zone_titre[max(0, pos-30):pos+len(mot_cle)+30]
                if not any(neg in contexte for neg in ['pas de ', 'non ', 'aucun ']):
                    metiers_trouves.append((domaine, metier, len(mot_cle)))
        
        if metiers_trouves:
            metiers_trouves.sort(key=lambda x: x[2], reverse=True)
            return metiers_trouves[0][0], metiers_trouves[0][1], 0.85
        
        metiers_trouves = []
        for mot_cle, (domaine, metier) in METIERS_PRIORITAIRES.items():
            if mot_cle in t:
                metiers_trouves.append((domaine, metier, len(mot_cle)))
        
        if metiers_trouves:
            metiers_trouves.sort(key=lambda x: x[2], reverse=True)
            return metiers_trouves[0][0], metiers_trouves[0][1], 0.7
        
        try:
            t_clean = re.sub(r'\s+', ' ', t)
            pred = self.pipeline.predict([t_clean])[0]
            proba = max(self.pipeline.predict_proba([t_clean])[0])
            if proba < 0.15: return 'Non classé', 'RAS', proba
            return pred, 'RAS', proba
        except:
            return 'Non classé', 'RAS', 0.0

# ============================================================
# SCORE
# ============================================================
def calculer_score(infos):
    pts_dip = min(infos.get('points_niveau', 0), 30)
    exp = infos.get('annees_experience', 0)
    if exp >= 10: pts_exp = 40
    elif exp >= 7: pts_exp = 35
    elif exp >= 5: pts_exp = 30
    elif exp >= 3: pts_exp = 22
    elif exp >= 2: pts_exp = 15
    elif exp >= 1: pts_exp = 8
    else: pts_exp = 0
    pts_comp = 0
    if infos.get('nom') and infos.get('nom') != "RAS": pts_comp += 3
    if infos.get('prenom') and infos.get('prenom') != "RAS": pts_comp += 3
    if infos.get('email') and infos.get('email') != "RAS": pts_comp += 8
    if infos.get('telephone') and infos.get('telephone') != "RAS": pts_comp += 8
    if infos.get('competences') and len(infos['competences']) >= 3: pts_comp += 4
    elif infos.get('competences'): pts_comp += 2
    if infos.get('ville') and infos.get('ville') != "RAS": pts_comp += 4
    pts_comp = min(pts_comp, 30)
    return {'total': pts_dip + pts_exp + pts_comp,
            'diplome': pts_dip, 'experience': pts_exp, 'completude': pts_comp}

# ============================================================
# LECTURE CV
# ============================================================
def est_cv_image(chemin):
    try:
        if not chemin.lower().endswith('.pdf'): return False
        doc = fitz.open(chemin)
        texte_total = ""
        for page in doc: texte_total += page.get_text()
        doc.close()
        return len(texte_total.strip()) < 200
    except: return False


def lire_avec_easyocr(chemin):
    reader = st.session_state.get('easyocr_reader')
    if not reader: return ""
    try:
        if chemin.lower().endswith('.pdf'):
            doc = fitz.open(chemin)
            texte_final = ""
            for page in doc:
                pix = page.get_pixmap(dpi=100)
                img_bytes = pix.tobytes("png")
                resultats = reader.readtext(img_bytes, detail=0, paragraph=True,
                                            batch_size=4, workers=4)
                texte_final += "\n".join(resultats) + "\n"
            doc.close()
            return texte_final
        else:
            resultats = reader.readtext(chemin, detail=0, paragraph=True)
            return "\n".join(resultats)
    except: return ""


@st.cache_data(show_spinner=False)
def lire_cv(chemin):
    texte = ""
    try:
        if chemin.lower().endswith('.pdf'):
            doc = fitz.open(chemin)
            for page in doc: texte += page.get_text() + "\n"
            doc.close()
        elif chemin.lower().endswith('.docx'):
            d = docx.Document(chemin)
            texte = "\n".join(p.text for p in d.paragraphs)
            for tab in d.tables:
                for r in tab.rows:
                    for c in r.cells: texte += " " + c.text
            return texte
    except: pass
    if len(texte.strip()) >= 300: return texte
    texte_ocr = lire_avec_easyocr(chemin)
    if texte_ocr and len(texte_ocr.strip()) > len(texte.strip()):
        return texte_ocr
    return texte

# ============================================================
# EXPORT EXCEL
# ============================================================
def exporter_excel(resultats, nom_fichier=None):
    if not resultats: return None
    data = []
    for r in resultats:
        if not r.get('succes'): continue
        i = r['infos']
        s = r['scores']
        data.append({
            'Nom': i.get('nom', 'RAS'), 'Prénom': i.get('prenom', 'RAS'),
            'Email': i.get('email', 'RAS'), 'Téléphone': i.get('telephone', 'RAS'),
            'Âge': i.get('age', 'RAS'), 'Sexe': i.get('sexe', 'RAS'),
            'Situation': i.get('situation', 'RAS'),
            'Nationalité': i.get('nationalite', 'RAS'),
            'Ville': i.get('ville', 'RAS'),
            'Niveau': i.get('niveau_etudes', 'RAS'),
            'Établissement': i.get('etablissement', 'RAS'),
            'Expérience': i.get('annees_experience', 0),
            'Poste actuel': i.get('poste_actuel', 'RAS'),
            'Compétences': ', '.join(i.get('competences', [])),
            'Langues': i.get('langues', 'RAS'),
            'Certifications': i.get('certifications', 'RAS'),
            'Domaine': r['categorie'], 'Métier': r.get('metier', 'RAS'),
            'Confiance': f"{r['confiance']*100:.0f}%",
            'Score Total': s['total'], 'Score Diplôme': s['diplome'],
            'Score Expérience': s['experience'], 'Score Complétude': s['completude'],
            'Résumé': r.get('resume', '')
        })
    df = pd.DataFrame(data).sort_values('Score Total', ascending=False)
    df.insert(0, 'Rang', range(1, len(df) + 1))
    if not nom_fichier:
        nom_fichier = f"CVSmart_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    chemin = os.path.join(DOSSIER_EXCEL, nom_fichier)
    with pd.ExcelWriter(chemin, engine='openpyxl') as w:
        df.to_excel(w, sheet_name='CV', index=False)
        ws = w.sheets['CV']
        for col in ws.columns:
            m = max(len(str(c.value)) if c.value else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(m + 2, 50)
    return chemin

# ============================================================
# MATCHING
# ============================================================
def calculer_matching(infos, offre):
    offre_l = offre.lower()
    cv_l = infos.get('texte_brut', '').lower()
    mots_offre = set(re.findall(r'\b[a-zà-ÿ]{3,}\b', offre_l))
    mots_offre -= {'avec', 'pour', 'dans', 'vous', 'nous', 'les', 'des', 'une',
                   'que', 'qui', 'sur', 'est', 'sont', 'être', 'avoir', 'plus'}
    mots_cv = set(re.findall(r'\b[a-zà-ÿ]{3,}\b', cv_l))
    communs = mots_offre & mots_cv
    score = (len(communs) / len(mots_offre) * 100) if mots_offre else 0
    m = re.search(r'(\d+)\s*(?:ans?|années?)\s*(?:d\')?expérience', offre_l)
    if m:
        exp_d = int(m.group(1))
        exp_cv = infos.get('annees_experience', 0)
        score = min(score + 10, 100) if exp_cv >= exp_d else max(score - 15, 0)
    return {'score': round(score, 1),
            'communs': sorted(communs)[:10],
            'manquants': sorted(mots_offre - mots_cv)[:5]}

# ============================================================
# RÉSUMÉ
# ============================================================
def generer_resume(infos, categorie, metier="RAS"):
    nom = f"{infos.get('prenom', '')} {infos.get('nom', '')}".strip()
    if nom in ["RAS RAS", "RAS", ""]: nom = "Ce candidat"
    niveau = infos.get('niveau_etudes', 'RAS')
    exp = infos.get('annees_experience', 0)
    comps = infos.get('competences', [])[:5]
    ville = infos.get('ville', 'RAS')
    p1 = f"{nom}"
    if metier and metier != "RAS": p1 += f", {metier}"
    elif niveau != "RAS": p1 += f", titulaire d'un {niveau}"
    if ville != "RAS": p1 += f", basé à {ville}"
    p1 += "."
    if exp > 0:
        p2 = f"Possède {exp} an{'s' if exp > 1 else ''} d'expérience"
        if categorie and categorie != 'Non classé':
            p2 += f" dans le domaine {categorie.lower()}"
        p2 += "."
    else:
        p2 = "Profil débutant, prêt à s'investir."
    p3 = f"Compétences clés : {', '.join(comps[:4])}." if comps else ""
    return f"{p1} {p2} {p3}".strip()

# ============================================================
# INIT SESSION
# ============================================================
@st.cache_resource
def charger_easyocr():
    try:
        import easyocr
        return easyocr.Reader(['fr', 'en'], gpu=False, verbose=False)
    except: return None


@st.cache_resource
def charger_classificateur():
    classif = Classificateur()
    classif.entrainer()
    return classif


def init_session():
    if 'classif' not in st.session_state:
        with st.spinner("🚀 Initialisation..."):
            st.session_state.classif = charger_classificateur()
    if 'extract' not in st.session_state:
        st.session_state.extract = ExtracteurInfos()
    if 'easyocr_reader' not in st.session_state:
        st.session_state.easyocr_reader = charger_easyocr()
    if 'resultats' not in st.session_state:
        st.session_state.resultats = []
    if 'page_active' not in st.session_state:
        st.session_state.page_active = "Accueil"
    if 'apprentissage' not in st.session_state:
        st.session_state.apprentissage = charger_apprentissage()
    if 'cv_a_valider' not in st.session_state:
        st.session_state.cv_a_valider = []
    if 'seuil_confiance' not in st.session_state:
        st.session_state.seuil_confiance = SEUIL_CONFIANCE_DEFAUT
        # ============================================================
# MENU
# ============================================================
def afficher_menu():
    st.markdown("""
    <div class="cvsmart-header">
        <div style="display: flex; align-items: center; gap: 1rem;">
            <div style="font-size: 2.5rem;">📄</div>
            <div>
                <div class="cvsmart-logo-text">CVSmart</div>
                <div class="cvsmart-logo-sub">L'intelligence qui trie vos talents</div>
            </div>
        </div>
        <div style="color: white; font-weight: 700; font-size: 1rem;">v11.0</div>
    </div>
    """, unsafe_allow_html=True)
    
    if st.session_state.get('resultats'):
        nb_total = len(st.session_state.resultats)
        nb_ok = sum(1 for r in st.session_state.resultats if r.get('succes'))
        nb_valider = len(st.session_state.cv_a_valider)
        nb_appris = st.session_state.apprentissage['stats']['total_corrections']
        st.markdown(f"""
        <div style="display: flex; justify-content: flex-end; gap: 1.5rem;
                    padding: 0.5rem 1rem; margin-bottom: 0.5rem;">
            <span style="color:#1a3a6b; font-size:0.9rem; font-weight:600;">
                📄 <b style="color:#2B3F8C;">{nb_total}</b>
            </span>
            <span style="color:#1a3a6b; font-size:0.9rem; font-weight:600;">
                ✅ <b style="color:#E63946;">{nb_ok - nb_valider}</b>
            </span>
            <span style="color:#1a3a6b; font-size:0.9rem; font-weight:600;">
                ⚠️ <b style="color:#FF6B00;">{nb_valider}</b>
            </span>
            <span style="color:#1a3a6b; font-size:0.9rem; font-weight:600;">
                🧠 <b style="color:#8B00FF;">{nb_appris}</b>
            </span>
        </div>
        """, unsafe_allow_html=True)
    
    pages = [
        ("🏠 Accueil", "Accueil"),
        ("📧 Emails", "Emails"),
        ("🔍 Analyser", "Analyser"),
        ("📊 Résultats", "Résultats"),
        ("🎯 Filtrer", "Filtrer"),
        ("⚠️ Validation", "Validation"),
        ("🧠 Apprentissage", "Apprentissage"),
        ("🤝 Matching", "Matching"),
        ("⚙️ Paramètres", "Paramètres"),
    ]
    
    cols = st.columns(len(pages))
    for i, (label, page_name) in enumerate(pages):
        with cols[i]:
            is_active = st.session_state.page_active == page_name
            btn_type = "primary" if is_active else "secondary"
            # Badge pour validation
            affichage_label = label
            if page_name == "Validation" and st.session_state.cv_a_valider:
                affichage_label = f"{label} ({len(st.session_state.cv_a_valider)})"
            if st.button(affichage_label, key=f"nav_{page_name}",
                         use_container_width=True, type=btn_type):
                st.session_state.page_active = page_name
                st.rerun()
    
    st.markdown("<hr style='border:none; border-top:2px solid #e0e7ff; margin:1rem 0;'>",
                unsafe_allow_html=True)
    return st.session_state.page_active


# ============================================================
# PAGE ACCUEIL
# ============================================================
def page_accueil():
    st.markdown('<div class="main-title">Bienvenue sur CVSmart</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-subtitle">Classez, filtrez et exportez vos CV. L\'IA apprend de vos corrections.</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #2B3F8C 0%, #3A4FA8 100%);
                    border-radius: 20px; padding: 2.5rem; text-align: center;
                    box-shadow: 0 12px 32px rgba(43,63,140,0.3); margin: 1rem 0;">
            <div style="font-size: 4rem; margin-bottom: 1rem;">📄</div>
            <h1 style="color: white; margin: 0; font-weight: 900; font-size: 2.5rem;">CVSmart</h1>
            <p style="color: rgba(255,255,255,0.95); margin-top: 0.5rem;
                      font-size: 1.1rem; font-weight: 600;">L'intelligence qui trie vos talents</p>
            <p style="color: rgba(255,255,255,0.85); font-size: 0.9rem; margin-top: 1rem;">
                50 domaines · 250+ métiers · Apprentissage IA
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("""
        <div class="stat-card"><div>
            <div class="stat-label">⚡ Rapide</div>
            <div style="color:#1a3a6b; font-size:0.85rem; font-weight:600;">Analyse optimisée</div>
        </div><div class="stat-icon">⚡</div></div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class="stat-card"><div>
            <div class="stat-label">🎯 Précis</div>
            <div style="color:#1a3a6b; font-size:0.85rem; font-weight:600;">Domaine + Métier + Confiance</div>
        </div><div class="stat-icon">🎯</div></div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown("""
        <div class="stat-card"><div>
            <div class="stat-label">🧠 Intelligent</div>
            <div style="color:#1a3a6b; font-size:0.85rem; font-weight:600;">Apprend de vos corrections</div>
        </div><div class="stat-icon">🧠</div></div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### 🚀 Comment ça marche")
    
    etapes = [
        ("📂", "1. Import", "Sélection des CV"),
        ("📄", "2. Extraction", "Lecture + OCR"),
        ("🧠", "3. Analyse", "25+ infos"),
        ("🎯", "4. Classification", "Domaine + Métier"),
        ("⚠️", "5. Validation", "Si confiance faible"),
        ("📊", "6. Export", "Excel + Dossiers")
    ]
    cols = st.columns(6)
    for col, (icon, t, d) in zip(cols, etapes):
        with col:
            st.markdown(f"""
            <div class="step-box">
                <div class="step-icon">{icon}</div>
                <div class="step-title">{t}</div>
                <div class="step-desc">{d}</div>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #2B3F8C 0%, #1e3a6b 100%);
                    color: white; border-radius: 16px; padding: 1.5rem; text-align: center;">
            <div style="font-size: 1.8rem; margin-bottom: 0.5rem;">🧠</div>
            <div style="font-weight: 800; font-size: 1.1rem;">
                L'IA apprend de vos corrections et s'améliore à chaque utilisation
            </div>
        </div>
        """, unsafe_allow_html=True)


# ============================================================
# PAGE ANALYSER (V2 avec système de confiance)
# ============================================================
def page_analyser():
    st.markdown('<div class="main-title">Analyse et classement des CV</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-subtitle">Sélectionnez un dossier contenant vos CV.</div>', unsafe_allow_html=True)
    
    with st.expander("⚙️ Configuration du seuil de confiance", expanded=False):
        st.session_state.seuil_confiance = st.slider(
            "Seuil de confiance pour classement automatique (%)",
            min_value=30, max_value=95,
            value=st.session_state.seuil_confiance,
            help="En dessous de ce seuil, l'app vous demandera de valider"
        )
        st.info(f"📊 Seuil actuel : **{st.session_state.seuil_confiance}%**")
    
    col1, col2 = st.columns([4, 1])
    with col1:
        chemin_dossier = st.text_input(
            "📁 Dossier sélectionné",
            value=os.path.abspath("cv_a_traiter"),
            key="chemin_dossier_input"
        )
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("Changer", use_container_width=True, type="secondary", key="btn_changer")
    
    if st.session_state.resultats:
        nb_total = len(st.session_state.resultats)
        nb_ok = sum(1 for r in st.session_state.resultats if r.get('succes'))
        nb_valider = len(st.session_state.cv_a_valider)
        
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"""
            <div class="stat-card"><div>
                <div class="stat-label">📄 Total</div>
                <div class="stat-value">{nb_total}</div>
            </div><div class="stat-icon">📄</div></div>
            """, unsafe_allow_html=True)
        with c2:
            st.markdown(f"""
            <div class="stat-card"><div>
                <div class="stat-label">✅ Auto</div>
                <div class="stat-value" style="color:#E63946;">{nb_ok - nb_valider}</div>
            </div><div class="stat-icon">✅</div></div>
            """, unsafe_allow_html=True)
        with c3:
            st.markdown(f"""
            <div class="stat-card"><div>
                <div class="stat-label">⚠️ À valider</div>
                <div class="stat-value" style="color:#FF6B00;">{nb_valider}</div>
            </div><div class="stat-icon">⚠️</div></div>
            """, unsafe_allow_html=True)
        with c4:
            st.markdown(f"""
            <div class="stat-card"><div>
                <div class="stat-label">🧠 Appris</div>
                <div class="stat-value" style="color:#8B00FF;">{st.session_state.apprentissage['stats']['total_corrections']}</div>
            </div><div class="stat-icon">🧠</div></div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        lancer = st.button("🚀 Lancer l'analyse", type="primary", use_container_width=True, key="btn_lancer")
    with c2:
        vider = st.button("🗑️ Vider la session", use_container_width=True, type="secondary", key="btn_vider")
    
    if vider:
        st.session_state.resultats = []
        st.session_state.cv_a_valider = []
        st.rerun()
    
    if lancer:
        if not os.path.isdir(chemin_dossier):
            st.error(f"❌ Dossier introuvable : {chemin_dossier}")
            return
        
        fichiers = [f for f in os.listdir(chemin_dossier)
                    if f.lower().endswith(('.pdf', '.docx', '.doc'))]
        if not fichiers:
            st.warning(f"⚠️ Aucun CV dans {chemin_dossier}")
            return
        
        st.info(f"📊 {len(fichiers)} CV trouvés. Analyse en cours...")
        barre = st.progress(0)
        journal = st.empty()
        temps_estime = st.empty()
        logs = []
        temps_par_cv = []
        
        for i, nom in enumerate(fichiers):
            chemin = os.path.join(chemin_dossier, nom)
            heure = datetime.now().strftime("%H:%M:%S")
            logs.append(f"🔄 {heure} - **{nom}**...")
            journal.markdown("  \n".join(logs[-5:]))
            
            if temps_par_cv:
                moyenne = sum(temps_par_cv) / len(temps_par_cv)
                restant = moyenne * (len(fichiers) - i)
                temps_estime.info(f"⏱️ Temps estimé : **{int(restant)}s**")
            
            debut_cv = datetime.now()
            
            texte = lire_cv(chemin)
            if not texte:
                st.session_state.resultats.append({
                    'fichier': nom, 'succes': False, 'erreur': 'Lecture impossible'
                })
                continue
            
            # Extraction
            infos = st.session_state.extract.extraire_tout(texte)
            
            # Âge amélioré
            age_calculé, conf_age = calculer_age_depuis_annee(texte)
            if age_calculé != "RAS":
                infos['age'] = age_calculé
            
            # Sexe amélioré
            sexe_calculé, conf_sexe = determiner_sexe_avec_confiance(texte, infos.get('prenom', 'RAS'))
            infos['sexe'] = sexe_calculé
            
            # Apprentissage
            infos = appliquer_apprentissage(infos, st.session_state.apprentissage)
            
            # Classification
            cat, metier, conf = st.session_state.classif.classifier(texte)
            scores = calculer_score(infos)
            resume = generer_resume(infos, cat, metier)
            
            idx = len(st.session_state.resultats)
            st.session_state.resultats.append({
                'fichier': nom, 'succes': True,
                'categorie': cat, 'metier': metier, 'confiance': conf,
                'infos': infos, 'scores': scores, 'resume': resume
            })
            
            # Vérifier si validation nécessaire
            seuil = st.session_state.seuil_confiance / 100
            champs_faibles = verifier_champs_a_valider(infos, cat, metier, st.session_state.seuil_confiance)
            
            if conf < seuil or len(champs_faibles) >= 2:
                st.session_state.cv_a_valider.append(idx)
                logs.append(f"⚠️ {heure} - **{nom}** → {cat} > {metier} (À VALIDER)")
            else:
                logs.append(f"✅ {heure} - **{nom}** → {cat} > {metier}")
            
            # Copie dossier
            dest = os.path.join(DOSSIER_TRI, cat.replace(" ", "_"))
            os.makedirs(dest, exist_ok=True)
            try: shutil.copy2(chemin, os.path.join(dest, nom))
            except: pass
            
            duree_cv = (datetime.now() - debut_cv).total_seconds()
            temps_par_cv.append(duree_cv)
            barre.progress((i + 1) / len(fichiers))
        
        journal.markdown("  \n".join(logs[-8:]))
        temps_estime.empty()
        
        if st.session_state.cv_a_valider:
            st.warning(f"⚠️ {len(st.session_state.cv_a_valider)} CV nécessitent votre validation")
            st.info("👉 Allez dans **⚠️ Validation** pour les corriger")
        else:
            st.success(f"🎉 Analyse terminée ! {len(fichiers)} CV traités automatiquement.")
        st.balloons()


# ============================================================
# PAGE RÉSULTATS
# ============================================================
def page_resultats():
    st.markdown('<div class="main-title">Résultats de classification</div>', unsafe_allow_html=True)
    
    if not st.session_state.resultats:
        st.info("📊 Aucun résultat. Allez dans **Analyser** pour commencer.")
        return
    
    resultats = [r for r in st.session_state.resultats if r.get('succes')]
    total = len(st.session_state.resultats)
    classes = len(resultats)
    nc = total - classes
    
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("📄 Total des CV", total)
    with c2: st.metric("✅ Classés", classes)
    with c3: st.metric("❓ Non classés", nc)
    
    st.markdown("---")
    st.subheader("Répartition par domaine")
    
    compteur = {}
    for r in resultats:
        cat = r['categorie']
        compteur[cat] = compteur.get(cat, 0) + 1
    if nc > 0: compteur['Non classé'] = nc
    
    cols = st.columns(4)
    for i, (cat, nb) in enumerate(compteur.items()):
        with cols[i % 4]:
            data = CATEGORIES.get(cat, {'emoji': '❓', 'couleur': '#888888'})
            st.markdown(f"""
            <div class="cat-card" style="border-left-color: {data['couleur']};">
                <div class="cat-emoji">{data['emoji']}</div>
                <div>
                    <div class="cat-name">{cat}</div>
                    <div class="cat-count">{nb} CV</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    tab1, tab2 = st.tabs(["🏆 Classement", "📋 Détails"])
    
    with tab1:
        data = []
        for idx, r in enumerate(st.session_state.resultats):
            if not r.get('succes'): continue
            i = r['infos']
            a_valider = idx in st.session_state.cv_a_valider
            data.append({
                'Rang': 0,
                '⚠️': "⚠️" if a_valider else "✅",
                'Nom': i.get('nom', 'RAS'),
                'Prénom': i.get('prenom', 'RAS'),
                'Domaine': r['categorie'],
                'Métier': r.get('metier', 'RAS'),
                'Niveau': i.get('niveau_etudes', 'RAS'),
                'Ville': i.get('ville', 'RAS'),
                'Âge': i.get('age', 'RAS'),
                'Sexe': i.get('sexe', 'RAS'),
                'Exp.': i.get('annees_experience', 0),
                'Email': i.get('email', 'RAS'),
                'Téléphone': i.get('telephone', 'RAS'),
                'Score': r['scores']['total']
            })
        if data:
            df = pd.DataFrame(data).sort_values('Score', ascending=False)
            df['Rang'] = range(1, len(df) + 1)
            st.dataframe(df, use_container_width=True, height=400)
    
    with tab2:
        for idx, r in enumerate(st.session_state.resultats):
            if not r.get('succes'): continue
            i = r['infos']
            nom_complet = f"{i.get('prenom', 'RAS')} {i.get('nom', 'RAS')}".strip()
            data_cat = CATEGORIES.get(r['categorie'], {'emoji': '❓', 'couleur': '#888'})
            metier = r.get('metier', 'RAS')
            a_valider = idx in st.session_state.cv_a_valider
            
            marqueur = "⚠️" if a_valider else "✅"
            with st.expander(f"{marqueur} {data_cat['emoji']} **{nom_complet}** — {r['categorie']} > {metier} — Score : {r['scores']['total']}/100"):
                c1, c2 = st.columns([2, 1])
                with c1:
                    st.markdown("**📝 Résumé automatique**")
                    st.markdown(f"""
                    <div style="background: linear-gradient(135deg, #e0f0ff 0%, #d0f0e0 100%);
                                border-left: 5px solid #2B3F8C;
                                padding:1rem 1.2rem; border-radius:10px; margin-bottom:1rem;
                                color:#2B3F8C; font-weight:600;">
                        {r.get('resume', '')}
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.markdown("**👤 Informations extraites**")
                    for k, v in [
                        ('Nom', i.get('nom', 'RAS')), ('Prénom', i.get('prenom', 'RAS')),
                        ('Email', i.get('email', 'RAS')), ('Téléphone', i.get('telephone', 'RAS')),
                        ('Âge', i.get('age', 'RAS')), ('Sexe', i.get('sexe', 'RAS')),
                        ('Situation', i.get('situation', 'RAS')), ('Nationalité', i.get('nationalite', 'RAS')),
                        ('Ville', i.get('ville', 'RAS')), ('Niveau', i.get('niveau_etudes', 'RAS')),
                        ('Établissement', i.get('etablissement', 'RAS')),
                        ('Expérience', f"{i.get('annees_experience', 0)} ans"),
                        ('Poste actuel', i.get('poste_actuel', 'RAS')),
                        ('Langues', i.get('langues', 'RAS')),
                        ('Certifications', i.get('certifications', 'RAS')),
                    ]:
                        if v == "RAS":
                            st.markdown(f"- **{k}** : <span class='badge-ras'>RAS</span>", unsafe_allow_html=True)
                        else:
                            st.markdown(f"- **{k}** : {v}")
                    
                    comps = i.get('competences', [])
                    if comps:
                        st.markdown("**🛠️ Compétences**")
                        st.markdown(" ".join([f'<span class="skill-tag">{c}</span>' for c in comps[:15]]),
                                    unsafe_allow_html=True)
                    
                    # Bouton corriger
                    if st.button("✏️ Corriger ce CV", key=f"corr_res_{idx}"):
                        st.session_state.cv_a_valider = list(set(st.session_state.cv_a_valider + [idx]))
                        st.session_state.page_active = "Validation"
                        st.rerun()
                
                with c2:
                    st.markdown(f"""
                    <div style="text-align:center; padding:1rem;">
                        <div style="color:#2d4a7c; font-weight:700; font-size:0.9rem;">Classification</div>
                        <div style="margin:0.7rem 0;">
                            <span class="badge-cat" style="background:{data_cat['couleur']};">
                                {data_cat['emoji']} {r['categorie']}
                            </span>
                        </div>
                        <div style="margin:0.5rem 0;">
                            <span class="badge-metier">{metier}</span>
                        </div>
                        <div style="color:#2d4a7c; font-weight:600; font-size:0.85rem;">
                            Confiance : {r['confiance']*100:.0f}%
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
    
    st.markdown("---")
    if st.button("📥 Exporter en Excel", type="primary", key="btn_export_res"):
        chemin = exporter_excel(st.session_state.resultats)
        if chemin:
            st.success(f"✅ Exporté : {chemin}")
            with open(chemin, 'rb') as f:
                st.download_button("⬇️ Télécharger", data=f,
                                   file_name=os.path.basename(chemin),
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ============================================================
# MODALE DE CORRECTION
# ============================================================
@st.dialog("✏️ Correction des informations", width="large")
def modale_correction(idx_cv):
    r = st.session_state.resultats[idx_cv]
    infos = r['infos']
    
    st.markdown(f"### 📄 {r['fichier']}")
    st.markdown("---")
    
    # ========== IDENTITÉ ==========
    st.markdown("#### 👤 Identité")
    col1, col2 = st.columns(2)
    with col1:
        new_nom = st.text_input("Nom", value=infos.get('nom', 'RAS'), key=f"cn_{idx_cv}")
    with col2:
        new_prenom = st.text_input("Prénom", value=infos.get('prenom', 'RAS'), key=f"cp_{idx_cv}")
    
    col1, col2 = st.columns(2)
    with col1:
        new_email = st.text_input("Email", value=infos.get('email', 'RAS'), key=f"ce_{idx_cv}")
    with col2:
        new_tel = st.text_input("Téléphone", value=infos.get('telephone', 'RAS'), key=f"ct_{idx_cv}")
    
    # Sexe
    sexe_actuel = infos.get('sexe', 'RAS')
    options_sexe = ['RAS', 'Homme', 'Femme']
    idx_s = options_sexe.index(sexe_actuel) if sexe_actuel in options_sexe else 0
    new_sexe = st.radio("Sexe", options_sexe, index=idx_s, horizontal=True, key=f"cs_{idx_cv}")
    
    # Âge + Situation
    col1, col2 = st.columns(2)
    with col1:
        age_actuel = infos.get('age', 'RAS')
        age_def = 30 if age_actuel == 'RAS' else int(age_actuel)
        new_age = st.number_input("Âge", min_value=16, max_value=80, value=age_def, key=f"ca_{idx_cv}")
    with col2:
        situations = ['RAS', 'Célibataire', 'Marié(e)', 'Divorcé(e)', 'Veuf(ve)']
        sit_actuelle = infos.get('situation', 'RAS')
        idx_sit = situations.index(sit_actuelle) if sit_actuelle in situations else 0
        new_situation = st.selectbox("Situation", situations, index=idx_sit, key=f"csit_{idx_cv}")
    
    # Nationalité + Ville
    col1, col2 = st.columns(2)
    with col1:
        new_nationalite = st.text_input("Nationalité", value=infos.get('nationalite', 'RAS'), key=f"cnat_{idx_cv}")
    with col2:
        new_ville = st.text_input("Ville", value=infos.get('ville', 'RAS'), key=f"cv_{idx_cv}")
    
    st.markdown("---")
    
    # ========== FORMATION ==========
    st.markdown("#### 🎓 Formation")
    niveaux = ['RAS', 'Bac', 'BTS/DUT', 'Licence', 'Master/Ingénieur', 'Doctorat']
    niveau_actuel = infos.get('niveau_etudes', 'RAS')
    idx_n = niveaux.index(niveau_actuel) if niveau_actuel in niveaux else 0
    new_niveau = st.selectbox("Niveau d'études", niveaux, index=idx_n, key=f"cniv_{idx_cv}")
    
    new_etab = st.text_input("Établissement", value=infos.get('etablissement', 'RAS'), key=f"cetab_{idx_cv}")
    
    st.markdown("---")
    
    # ========== EXPÉRIENCE (NOUVEAU) ==========
    st.markdown("#### 💼 Expérience professionnelle")
    
    col1, col2 = st.columns(2)
    with col1:
        exp_actuelle = infos.get('annees_experience', 0)
        new_experience = st.number_input(
            "Années d'expérience",
            min_value=0,
            max_value=60,
            value=int(exp_actuelle),
            step=1,
            key=f"cexp_{idx_cv}",
            help="Corrigez le nombre d'années d'expérience si nécessaire"
        )
    with col2:
        new_poste = st.text_input("Poste actuel", value=infos.get('poste_actuel', 'RAS'), key=f"cposte_{idx_cv}")
    
    st.markdown("---")
    
    # ========== COMPÉTENCES ET LANGUES ==========
    st.markdown("#### 🛠️ Compétences et Langues")
    
    comps_actuelles = infos.get('competences', [])
    comps_str = ', '.join(comps_actuelles) if comps_actuelles else ''
    new_comps = st.text_area(
        "Compétences (séparées par virgule)",
        value=comps_str,
        height=80,
        key=f"ccomp_{idx_cv}"
    )
    
    col1, col2 = st.columns(2)
    with col1:
        new_langues = st.text_input("Langues", value=infos.get('langues', 'RAS'), key=f"clang_{idx_cv}")
    with col2:
        new_certs = st.text_input("Certifications", value=infos.get('certifications', 'RAS'), key=f"ccert_{idx_cv}")
    
    st.markdown("---")
    
    # ========== CLASSIFICATION ==========
    st.markdown("#### 🎯 Classification")
    
    col_d, col_m = st.columns(2)
    with col_d:
        domaine_actuel = r['categorie']
        domaines = list(CATEGORIES.keys())
        if domaine_actuel not in domaines:
            domaines.insert(0, domaine_actuel)
        idx_d = domaines.index(domaine_actuel) if domaine_actuel in domaines else 0
        new_domaine = st.selectbox("Domaine", domaines, index=idx_d, key=f"cd_{idx_cv}")
    
    with col_m:
        metiers_dispo = ['RAS'] + list(CATEGORIES.get(new_domaine, {}).get('metiers', {}).keys())
        metier_actuel = r.get('metier', 'RAS')
        if metier_actuel not in metiers_dispo:
            metiers_dispo.insert(1, metier_actuel)
        idx_m = metiers_dispo.index(metier_actuel) if metier_actuel in metiers_dispo else 0
        new_metier = st.selectbox("Métier", metiers_dispo, index=idx_m, key=f"cm_{idx_cv}")
    
    st.markdown("---")
    
    # ========== BOUTONS ==========
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("💾 Enregistrer + Apprendre", type="primary",
                     use_container_width=True, key=f"save_{idx_cv}"):
            
            # Sauvegarder TOUTES les corrections
            corrections = {
                'nom': new_nom,
                'prenom': new_prenom,
                'email': new_email,
                'telephone': new_tel,
                'sexe': new_sexe,
                'age': str(new_age),
                'situation': new_situation,
                'nationalite': new_nationalite,
                'ville': new_ville,
                'niveau_etudes': new_niveau,
                'etablissement': new_etab,
                'annees_experience': new_experience,
                'poste_actuel': new_poste,
                'langues': new_langues,
                'certifications': new_certs,
            }
            
            for champ, val in corrections.items():
                ancienne = infos.get(champ, 'RAS')
                if ancienne != val and ancienne != 'RAS' and ancienne != 0:
                    # Enregistrer pour apprentissage
                    enregistrer_correction(st.session_state.apprentissage, champ, ancienne, val)
                infos[champ] = val
            
            # Compétences
            if new_comps:
                infos['competences'] = [c.strip() for c in new_comps.split(',') if c.strip()]
            else:
                infos['competences'] = []
            
            # Classification
            r['categorie'] = new_domaine
            r['metier'] = new_metier
            
            # Recalculer
            r['scores'] = calculer_score(infos)
            r['resume'] = generer_resume(infos, new_domaine, new_metier)
            st.session_state.resultats[idx_cv] = r
            
            # Retirer de la liste à valider
            if idx_cv in st.session_state.cv_a_valider:
                st.session_state.cv_a_valider.remove(idx_cv)
            
            st.success("✅ Corrections enregistrées + apprises !")
            st.rerun()
    
    with col2:
        if st.button("⏭️ Plus tard", use_container_width=True, key=f"skip_{idx_cv}"):
            st.rerun()
    
    with col3:
        if st.button("🗑️ Fermer", use_container_width=True, key=f"close_{idx_cv}"):
            st.rerun()


# ============================================================
# PAGE VALIDATION
# ============================================================
def page_validation():
    st.markdown('<div class="main-title">⚠️ Validation requise</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-subtitle">Ces CV ont des champs à faible confiance. Corrigez-les pour améliorer l\'apprentissage.</div>', unsafe_allow_html=True)
    
    if not st.session_state.cv_a_valider:
        st.success("✅ Aucun CV en attente de validation !")
        st.markdown("---")
        st.subheader("📊 Statistiques d'apprentissage")
        stats = get_stats_apprentissage(st.session_state.apprentissage)
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("🎯 Corrections totales", stats['total_corrections'])
        with c2: st.metric("👤 Prénoms appris", stats['sexe_appris'])
        with c3: st.metric("🏷️ Domaines appris", stats['domaines_appris'])
        return
    
    st.warning(f"⚠️ **{len(st.session_state.cv_a_valider)} CV** nécessitent votre validation")
    st.markdown("---")
    
    for i, idx_cv in enumerate(st.session_state.cv_a_valider.copy()):
        if idx_cv >= len(st.session_state.resultats): continue
        r = st.session_state.resultats[idx_cv]
        infos = r['infos']
        nom = f"{infos.get('prenom', 'RAS')} {infos.get('nom', 'RAS')}".strip()
        
        with st.expander(f"⚠️ **{nom}** — {r['categorie']} > {r.get('metier', 'RAS')}", expanded=(i == 0)):
            st.markdown(f"**Fichier** : {r['fichier']}")
            st.markdown("")
            
            champs_a_valider = verifier_champs_a_valider(infos, r['categorie'], r.get('metier', 'RAS'), st.session_state.seuil_confiance)
            
            if champs_a_valider:
                st.markdown("**Champs à vérifier :**")
                for c in champs_a_valider:
                    conf_pct = c['confiance'] * 100
                    emoji = "🔴" if conf_pct < 30 else ("🟠" if conf_pct < 60 else "🟡")
                    st.markdown(f"{emoji} **{c['champ'].replace('_', ' ').title()}** : `{c['valeur_actuelle']}` — Confiance : {conf_pct:.0f}% (seuil : {c['seuil']}%)")
            
            st.markdown("")
            col1, col2 = st.columns(2)
            with col1:
                if st.button(f"✏️ Corriger", type="primary", key=f"corr_{idx_cv}"):
                    modale_correction(idx_cv)
            with col2:
                if st.button(f"✅ Marquer validé", key=f"valid_{idx_cv}"):
                    st.session_state.cv_a_valider.remove(idx_cv)
                    st.rerun()


# ============================================================
# PAGE APPRENTISSAGE
# ============================================================
def page_apprentissage():
    st.markdown('<div class="main-title">🧠 Apprentissage</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-subtitle">Voici ce que l\'application a appris de vos corrections.</div>', unsafe_allow_html=True)
    
    apprentissage = st.session_state.apprentissage
    stats = get_stats_apprentissage(apprentissage)
    
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("📊 Total corrections", stats['total_corrections'])
    with c2: st.metric("👤 Prénoms appris", stats['sexe_appris'])
    with c3: st.metric("🏷️ Domaines appris", stats['domaines_appris'])
    with c4: st.metric("💼 Métiers appris", stats['metiers_appris'])
    
    st.markdown("---")
    
    tabs = st.tabs(["⚧ Sexe", "🏷️ Domaines", "💼 Métiers"])
    
    with tabs[0]:
        st.subheader("Prénoms → Sexe")
        if apprentissage.get('sexe'):
            df = pd.DataFrame([{'Prénom': k.capitalize(), 'Sexe': v}
                              for k, v in apprentissage['sexe'].items()])
            st.dataframe(df, use_container_width=True)
        else:
            st.info("Aucune donnée apprise.")
    
    with tabs[1]:
        st.subheader("Mots-clés → Domaines")
        if apprentissage.get('domaines'):
            df = pd.DataFrame([{'Mot-clé': k, 'Domaine': v}
                              for k, v in apprentissage['domaines'].items()])
            st.dataframe(df, use_container_width=True)
        else:
            st.info("Aucune donnée apprise.")
    
    with tabs[2]:
        st.subheader("Mots-clés → Métiers")
        if apprentissage.get('metiers'):
            df = pd.DataFrame([{'Mot-clé': k, 'Métier': v}
                              for k, v in apprentissage['metiers'].items()])
            st.dataframe(df, use_container_width=True)
        else:
            st.info("Aucune donnée apprise.")
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🗑️ Réinitialiser", type="secondary", use_container_width=True):
            st.session_state.apprentissage = {
                'sexe': {}, 'domaines': {}, 'metiers': {},
                'stats': {'total_corrections': 0, 'derniere_maj': None}
            }
            sauvegarder_apprentissage(st.session_state.apprentissage)
            st.success("✅ Réinitialisé")
            st.rerun()
    with col2:
        if st.button("💾 Sauvegarder", type="primary", use_container_width=True):
            sauvegarder_apprentissage(st.session_state.apprentissage)
            st.success("✅ Sauvegardé dans apprentissage.json")


# ============================================================
# PAGE FILTRER
# ============================================================
def page_filtrer():
    st.markdown('<div class="main-title">🎯 Filtrer les CV</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-subtitle">Combinez les critères pour trouver les candidats idéaux.</div>', unsafe_allow_html=True)
    
    if not st.session_state.resultats:
        st.info("📊 Aucun CV analysé. Allez dans **Analyser** d'abord.")
        return
    
    resultats = [r for r in st.session_state.resultats if r.get('succes')]
    
    col_filtres, col_resultats = st.columns([1, 3])
    
    with col_filtres:
        st.markdown('<div class="filter-container">', unsafe_allow_html=True)
        st.markdown('<div class="filter-header">🔍 FILTRES</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="filter-group"><div class="filter-group-title">📚 Niveau</div>', unsafe_allow_html=True)
        niveaux_possibles = ['Bac', 'BTS/DUT', 'Licence', 'Master/Ingénieur', 'Doctorat']
        niveaux_counts = {n: sum(1 for r in resultats if r['infos'].get('niveau_etudes') == n) for n in niveaux_possibles}
        niveaux_selection = st.multiselect("Niveau", options=niveaux_possibles,
            format_func=lambda x: f"{x} ({niveaux_counts[x]})", key="f_niveau", label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="filter-group"><div class="filter-group-title">👤 Sexe</div>', unsafe_allow_html=True)
        sexes_possibles = ['Homme', 'Femme', 'RAS']
        sexes_counts = {s: sum(1 for r in resultats if r['infos'].get('sexe') == s) for s in sexes_possibles}
        sexes_selection = st.multiselect("Sexe", options=sexes_possibles,
            format_func=lambda x: f"{x} ({sexes_counts[x]})", key="f_sexe", label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="filter-group"><div class="filter-group-title">📍 Ville</div>', unsafe_allow_html=True)
        villes_dispo = sorted(set(r['infos'].get('ville', 'RAS') for r in resultats))
        villes_counts = {v: sum(1 for r in resultats if r['infos'].get('ville') == v) for v in villes_dispo}
        villes_selection = st.multiselect("Ville", options=villes_dispo,
            format_func=lambda x: f"{x} ({villes_counts[x]})", key="f_ville", label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="filter-group"><div class="filter-group-title">🏷️ Domaine</div>', unsafe_allow_html=True)
        cats_dispo = sorted(set(r['categorie'] for r in resultats))
        cats_counts = {c: sum(1 for r in resultats if r['categorie'] == c) for c in cats_dispo}
        cats_selection = st.multiselect("Domaine", options=cats_dispo,
            format_func=lambda x: f"{x} ({cats_counts[x]})", key="f_cat", label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="filter-group"><div class="filter-group-title">💼 Métier</div>', unsafe_allow_html=True)
        metiers_dispo = sorted(set(r.get('metier', 'RAS') for r in resultats if r.get('metier', 'RAS') != 'RAS'))
        metiers_counts = {m: sum(1 for r in resultats if r.get('metier') == m) for m in metiers_dispo}
        metiers_selection = st.multiselect("Métier", options=metiers_dispo,
            format_func=lambda x: f"{x} ({metiers_counts[x]})", key="f_metier", label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="filter-group"><div class="filter-group-title">📅 Âge</div>', unsafe_allow_html=True)
        ages_valides = [int(r['infos']['age']) for r in resultats if str(r['infos'].get('age', 'RAS')).isdigit()]
        age_min, age_max = (min(ages_valides), max(ages_valides)) if ages_valides else (18, 60)
        age_range = st.slider("Âge", min_value=16, max_value=70, value=(age_min, age_max), key="f_age")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="filter-group"><div class="filter-group-title">💼 Expérience</div>', unsafe_allow_html=True)
        exp_range = st.slider("Expérience", min_value=0, max_value=40, value=(0, 40), key="f_exp")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="filter-group"><div class="filter-group-title">⭐ Score min</div>', unsafe_allow_html=True)
        score_min = st.slider("Score", min_value=0, max_value=100, value=0, key="f_score")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="filter-group"><div class="filter-group-title">🔍 Recherche</div>', unsafe_allow_html=True)
        recherche = st.text_input("Recherche", placeholder="Ex: marketing...", key="f_recherche", label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    resultats_filtres = resultats.copy()
    
    if niveaux_selection:
        resultats_filtres = [r for r in resultats_filtres if r['infos'].get('niveau_etudes') in niveaux_selection]
    if sexes_selection:
        resultats_filtres = [r for r in resultats_filtres if r['infos'].get('sexe') in sexes_selection]
    if villes_selection:
        resultats_filtres = [r for r in resultats_filtres if r['infos'].get('ville') in villes_selection]
    if cats_selection:
        resultats_filtres = [r for r in resultats_filtres if r['categorie'] in cats_selection]
    if metiers_selection:
        resultats_filtres = [r for r in resultats_filtres if r.get('metier') in metiers_selection]
    
    resultats_filtres = [r for r in resultats_filtres 
                         if (str(r['infos'].get('age', 'RAS')).isdigit() 
                             and age_range[0] <= int(r['infos']['age']) <= age_range[1])
                         or r['infos'].get('age') == 'RAS']
    
    resultats_filtres = [r for r in resultats_filtres if exp_range[0] <= r['infos'].get('annees_experience', 0) <= exp_range[1]]
    resultats_filtres = [r for r in resultats_filtres if r['scores']['total'] >= score_min]
    
    if recherche:
        rech_lower = recherche.lower()
        resultats_filtres = [r for r in resultats_filtres if rech_lower in r['infos'].get('texte_brut', '').lower()]
    
    with col_resultats:
        nb_initial = len(resultats)
        nb_filtres = len(resultats_filtres)
        
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #2B3F8C 0%, #E63946 100%);
                    border-radius: 12px; padding: 1rem 1.5rem; color: white;
                    margin-bottom: 1rem; display: flex; justify-content: space-between;">
            <div style="font-weight: 800; font-size: 1.1rem;">🎯 {nb_filtres} CV correspondent</div>
            <div style="opacity: 0.9; font-size: 0.9rem;">sur {nb_initial} CV analysés</div>
        </div>
        """, unsafe_allow_html=True)
        
        if not resultats_filtres:
            st.warning("⚠️ Aucun CV ne correspond.")
        else:
            for r in sorted(resultats_filtres, key=lambda x: x['scores']['total'], reverse=True):
                i = r['infos']
                nom = f"{i.get('prenom', 'RAS')} {i.get('nom', 'RAS')}".strip()
                data_cat = CATEGORIES.get(r['categorie'], {'emoji': '❓', 'couleur': '#888'})
                metier = r.get('metier', 'RAS')
                
                st.markdown(f"""
                <div class="result-item">
                    <div style="display: flex; justify-content: space-between;">
                        <div style="flex: 1;">
                            <div style="font-weight: 800; color: #2B3F8C; font-size: 1.05rem;">{nom}</div>
                            <div style="color: #2d4a7c; font-size: 0.85rem; margin-top: 0.3rem;">
                                {i.get('niveau_etudes', 'RAS')} · {i.get('sexe', 'RAS')} · {i.get('age', 'RAS')} ans · 📍 {i.get('ville', 'RAS')}
                            </div>
                            <div style="color: #4a6fa5; font-size: 0.8rem; margin-top: 0.3rem;">
                                ✉️ {i.get('email', 'RAS')} · 📞 {i.get('telephone', 'RAS')}
                            </div>
                        </div>
                        <div style="text-align: right;">
                            <span class="badge-cat" style="background: {data_cat['couleur']}; font-size: 0.7rem;">
                                {data_cat['emoji']} {r['categorie']}
                            </span>
                            <span class="badge-metier" style="font-size: 0.7rem;">{metier}</span>
                            <div style="font-size: 1.5rem; font-weight: 900; color: #2B3F8C; margin-top: 0.5rem;">
                                {r['scores']['total']}
                            </div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
        
        if resultats_filtres:
            st.markdown("---")
            st.subheader("📤 Exporter la sélection")
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("📁 Créer un dossier", type="primary", use_container_width=True, key="btn_export_dossier"):
                    nom_parts = []
                    if niveaux_selection: nom_parts.append("_".join(niveaux_selection).replace("/", ""))
                    if sexes_selection: nom_parts.append("_".join(sexes_selection))
                    if villes_selection: nom_parts.append("_".join(villes_selection))
                    if cats_selection: nom_parts.append("_".join(cats_selection))
                    if metiers_selection: nom_parts.append("_".join(metiers_selection))
                    if not nom_parts: nom_parts.append("Selection")
                    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
                    nom_dossier = "_".join(nom_parts).replace(" ", "_") + f"_{horodatage}"
                    chemin_dossier = os.path.join(DOSSIER_FILTRES, nom_dossier)
                    os.makedirs(chemin_dossier, exist_ok=True)
                    for r in resultats_filtres:
                        fichier_orig = r.get('fichier', '')
                        for racine in [".", "cv_a_traiter", DOSSIER_TRI]:
                            chemin_rech = os.path.join(racine, fichier_orig)
                            if os.path.exists(chemin_rech):
                                shutil.copy2(chemin_rech, os.path.join(chemin_dossier, fichier_orig))
                                break
                    st.success(f"✅ Dossier créé : {chemin_dossier}")
            with col_b:
                if st.button("📥 Exporter Excel", type="primary", use_container_width=True, key="btn_export_filtre"):
                    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
                    chemin = exporter_excel(resultats_filtres, f"Selection_{horodatage}.xlsx")
                    if chemin:
                        st.success(f"✅ Exporté : {chemin}")
                        with open(chemin, 'rb') as f:
                            st.download_button("⬇️ Télécharger", data=f,
                                file_name=os.path.basename(chemin),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ============================================================
# PAGE MATCHING
# ============================================================
def page_matching():
    st.markdown('<div class="main-title">🤝 Matching avec une offre</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-subtitle">Collez une offre et voyez les CV qui correspondent.</div>', unsafe_allow_html=True)
    
    offre = st.text_area("📝 Description de l'offre", height=200,
        placeholder="Ex: Développeur Python Senior\n5 ans d'expérience minimum\nPython, Django, AWS",
        key="offre_text")
    
    if st.button("🎯 Lancer le matching", type="primary", key="btn_matching"):
        if not offre.strip():
            st.warning("⚠️ Collez une offre")
            return
        if not st.session_state.resultats:
            st.warning("⚠️ Aucun CV analysé.")
            return
        
        resultats = [r for r in st.session_state.resultats if r.get('succes')]
        matchings = []
        for r in resultats:
            m = calculer_matching(r['infos'], offre)
            matchings.append({'CV': r, **m})
        matchings.sort(key=lambda x: x['score'], reverse=True)
        
        st.markdown("---")
        st.subheader("🏆 Résultats du matching")
        
        for i, m in enumerate(matchings[:10]):
            r = m['CV']
            i_info = r['infos']
            nom = f"{i_info.get('prenom', 'RAS')} {i_info.get('nom', 'RAS')}".strip()
            
            if m['score'] >= 70: badge = "✅ MATCH FORT"
            elif m['score'] >= 40: badge = "⚠️ MATCH PARTIEL"
            else: badge = "❌ FAIBLE"
            
            with st.expander(f"**{nom}** — {badge} : {m['score']}%", expanded=(i < 3)):
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"**Score : {m['score']}%**")
                    st.progress(m['score'] / 100)
                    st.markdown(f"**Domaine :** {r['categorie']}")
                    st.markdown(f"**Métier :** {r.get('metier', 'RAS')}")
                    st.markdown(f"**Expérience :** {i_info.get('annees_experience', 0)} ans")
                with c2:
                    if m['communs']:
                        st.markdown("**✅ Points forts :**")
                        for mot in m['communs'][:6]: st.markdown(f"- {mot}")
                    if m['manquants']:
                        st.markdown("**⚠️ À améliorer :**")
                        for mot in m['manquants'][:4]: st.markdown(f"- {mot}")


# ============================================================
# PAGE PARAMÈTRES
# ============================================================
# ============================================================
# PAGE EMAILS
# ============================================================
def page_emails():
    """Page de gestion des emails - Récupération automatique des CV"""
    st.markdown('<div class="main-title">📧 Boîte Mail RH</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-subtitle">Récupérez automatiquement les CV envoyés par email.</div>', unsafe_allow_html=True)
    
    # Charger les identifiants depuis secrets.toml
    try:
        config = st.secrets["email"]
        serveur = config["imap_server"]
        port = config["imap_port"]
        email_address = config["email_address"]
        mot_de_passe = config["app_password"]
        dossier_traites = config.get("dossier_traites", "CV traités")
    except Exception as e:
        st.error(f"❌ Impossible de lire les identifiants : {e}")
        st.info("💡 Vérifiez que le fichier `.streamlit/secrets.toml` existe et est bien configuré.")
        return
    
    # Affichage de la config
    st.markdown(f"""
    <div style="background: white; border-radius: 12px; padding: 1rem 1.5rem;
                border: 2px solid #e0e7ff; margin-bottom: 1rem;">
        <div style="color: #2d4a7c; font-weight: 700; font-size: 0.9rem;">
            📬 Boîte connectée : <b style="color: #2B3F8C;">{email_address}</b>
        </div>
        <div style="color: #64748b; font-size: 0.8rem; margin-top: 0.3rem;">
            Serveur : {serveur}:{port} · Dossier traités : "{dossier_traites}"
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Boutons
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        verifier = st.button("🔄 Vérifier les emails", type="primary", 
                             use_container_width=True, key="btn_verifier_emails")
    with col2:
        if st.button("🗑️ Vider", use_container_width=True, 
                     type="secondary", key="btn_vider_emails"):
            if 'emails_recuperes' in st.session_state:
                del st.session_state['emails_recuperes']
            st.rerun()
    
    # Vérification
    if verifier:
        with st.spinner("📥 Connexion à la boîte mail..."):
            gestionnaire = GestionnaireEmail(serveur, port, email_address, mot_de_passe, dossier_traites)
            emails, message = gestionnaire.recuperer_emails_non_lus(max_emails=30)
            gestionnaire.se_deconnecter()
            
            st.session_state['emails_recuperes'] = emails
            st.session_state['config_email'] = {
                'serveur': serveur, 'port': port,
                'email': email_address, 'mdp': mot_de_passe,
                'dossier': dossier_traites
            }
            
            if emails:
                st.success(f"✅ {message}")
            else:
                st.info(f"📭 {message}")
    
    # Affichage des emails
    if 'emails_recuperes' in st.session_state:
        emails = st.session_state['emails_recuperes']
        
        if not emails:
            st.info("📭 Aucun email avec CV trouvé.")
            return
        
        st.markdown("---")
        
        # Statistiques
        nb_emails = len(emails)
        nb_cv = sum(e['nb_cv'] for e in emails)
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"""
            <div class="stat-card">
                <div>
                    <div class="stat-label">📧 Emails</div>
                    <div class="stat-value">{nb_emails}</div>
                </div>
                <div class="stat-icon">📧</div>
            </div>
            """, unsafe_allow_html=True)
        with c2:
            st.markdown(f"""
            <div class="stat-card">
                <div>
                    <div class="stat-label">📎 CV</div>
                    <div class="stat-value" style="color:#E63946;">{nb_cv}</div>
                </div>
                <div class="stat-icon">📎</div>
            </div>
            """, unsafe_allow_html=True)
        with c3:
            st.markdown(f"""
            <div class="stat-card">
                <div>
                    <div class="stat-label">📨 À traiter</div>
                    <div class="stat-value" style="color:#FF6B00;">{nb_emails}</div>
                </div>
                <div class="stat-icon">📨</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.subheader(f"📬 {nb_emails} email(s) avec CV")
        
        # Liste des emails
        for i, em in enumerate(emails):
            with st.expander(f"📧 **{em['email_expediteur']}** — {em['sujet'][:60]}", expanded=(i == 0)):
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    st.markdown(f"**De :** {em['expediteur']}")
                    st.markdown(f"**Sujet :** {em['sujet']}")
                    st.markdown(f"**Date :** {em['date']}")
                    st.markdown(f"**Pièces jointes ({em['nb_cv']}) :**")
                    
                    for pj in em['pieces_jointes']:
                        taille_ko = pj['taille'] / 1024
                        st.markdown(f"- 📎 `{pj['nom_original']}` ({taille_ko:.0f} Ko)")
                
                with col2:
                    st.markdown("**Actions :**")
                    
                    # Bouton ANALYSER
                    if st.button("✅ Analyser les CV", type="primary",
                                 use_container_width=True, key=f"btn_analyser_{i}"):
                        with st.spinner(f"📄 Analyse de {em['nb_cv']} CV..."):
                            cv_analyses = []
                            
                            for pj in em['pieces_jointes']:
                                chemin = pj['chemin']
                                texte = lire_cv(chemin)
                                if not texte:
                                    continue
                                
                                # Extraction
                                infos = st.session_state.extract.extraire_tout(texte)
                                
                                # Âge amélioré
                                age_calc, _ = calculer_age_depuis_annee(texte)
                                if age_calc != "RAS":
                                    infos['age'] = age_calc
                                
                                # Sexe amélioré
                                sexe_calc, _ = determiner_sexe_avec_confiance(texte, infos.get('prenom', 'RAS'))
                                infos['sexe'] = sexe_calc
                                
                                # Apprentissage
                                infos = appliquer_apprentissage(infos, st.session_state.apprentissage)
                                
                                # Classification
                                cat, metier, conf = st.session_state.classif.classifier(texte)
                                scores = calculer_score(infos)
                                resume = generer_resume(infos, cat, metier)
                                
                                # Ajouter aux résultats
                                st.session_state.resultats.append({
                                    'fichier': pj['nom_original'],
                                    'succes': True,
                                    'categorie': cat,
                                    'metier': metier,
                                    'confiance': conf,
                                    'infos': infos,
                                    'scores': scores,
                                    'resume': resume,
                                    'source': f"📧 {em['email_expediteur']}"
                                })
                                
                                # Validation si nécessaire
                                seuil = st.session_state.seuil_confiance / 100
                                champs_faibles = verifier_champs_a_valider(infos, cat, metier, st.session_state.seuil_confiance)
                                if conf < seuil or len(champs_faibles) >= 2:
                                    st.session_state.cv_a_valider.append(len(st.session_state.resultats) - 1)
                                
                                cv_analyses.append({
                                    'nom': pj['nom_original'],
                                    'domaine': cat,
                                    'metier': metier
                                })
                                
                                # Copie dans dossier
                                domaine_clean = cat.replace(" ", "_").replace("/", "-")
                                if metier and metier != "RAS":
                                    metier_clean = metier.replace(" ", "_").replace("/", "-")
                                    dest = os.path.join(DOSSIER_TRI, domaine_clean, metier_clean)
                                else:
                                    dest = os.path.join(DOSSIER_TRI, domaine_clean, "_Non_renseigné")
                                os.makedirs(dest, exist_ok=True)
                                try:
                                    shutil.copy2(chemin, os.path.join(dest, pj['nom_sauvegarde']))
                                except:
                                    pass
                            
                            # Déplacer l'email dans "CV traités"
                            try:
                                config = st.session_state['config_email']
                                gest = GestionnaireEmail(
                                    config['serveur'], config['port'],
                                    config['email'], config['mdp'], config['dossier']
                                )
                                gest.se_connecter()
                                gest.deplacer_email(em['id'])
                                gest.se_deconnecter()
                            except:
                                pass
                            
                            st.success(f"✅ {len(cv_analyses)} CV analysé(s) et classé(s) !")
                            st.info(f"📧 Email déplacé dans '{st.session_state['config_email']['dossier']}'")
                            st.balloons()
                            
                            # Retirer de la liste
                            st.session_state['emails_recuperes'] = [
                                e for e in emails if e['id'] != em['id']
                            ]
                            st.rerun()
                    
                    # Bouton IGNORER
                    if st.button("❌ Ignorer", use_container_width=True,
                                 key=f"btn_ignorer_{i}"):
                        st.session_state['emails_recuperes'] = [
                            e for e in emails if e['id'] != em['id']
                        ]
                        st.rerun()
        
        # Export global
        if emails:
            st.markdown("---")
            if st.button("📥 Analyser TOUS les emails d'un coup", type="primary",
                         use_container_width=True, key="btn_analyser_tous"):
                total_cv = 0
                for em in emails:
                    for pj in em['pieces_jointes']:
                        chemin = pj['chemin']
                        texte = lire_cv(chemin)
                        if not texte:
                            continue
                        
                        infos = st.session_state.extract.extraire_tout(texte)
                        age_calc, _ = calculer_age_depuis_annee(texte)
                        if age_calc != "RAS":
                            infos['age'] = age_calc
                        sexe_calc, _ = determiner_sexe_avec_confiance(texte, infos.get('prenom', 'RAS'))
                        infos['sexe'] = sexe_calc
                        infos = appliquer_apprentissage(infos, st.session_state.apprentissage)
                        cat, metier, conf = st.session_state.classif.classifier(texte)
                        scores = calculer_score(infos)
                        resume = generer_resume(infos, cat, metier)
                        
                        st.session_state.resultats.append({
                            'fichier': pj['nom_original'], 'succes': True,
                            'categorie': cat, 'metier': metier, 'confiance': conf,
                            'infos': infos, 'scores': scores, 'resume': resume,
                            'source': f"📧 {em['email_expediteur']}"
                        })
                        
                        seuil = st.session_state.seuil_confiance / 100
                        champs_faibles = verifier_champs_a_valider(infos, cat, metier, st.session_state.seuil_confiance)
                        if conf < seuil or len(champs_faibles) >= 2:
                            st.session_state.cv_a_valider.append(len(st.session_state.resultats) - 1)
                        
                        domaine_clean = cat.replace(" ", "_").replace("/", "-")
                        metier_clean = (metier if metier != "RAS" else "_Non_renseigné").replace(" ", "_").replace("/", "-")
                        dest = os.path.join(DOSSIER_TRI, domaine_clean, metier_clean)
                        os.makedirs(dest, exist_ok=True)
                        try:
                            shutil.copy2(chemin, os.path.join(dest, pj['nom_sauvegarde']))
                        except:
                            pass
                        total_cv += 1
                
                st.success(f"🎉 {total_cv} CV analysés et classés !")
                st.session_state['emails_recuperes'] = []
                st.balloons()
                st.rerun()
    else:
        st.info("👆 Cliquez sur **'Vérifier les emails'** pour récupérer les CV reçus.")
def page_parametres():
    st.markdown('<div class="main-title">⚙️ Paramètres</div>', unsafe_allow_html=True)
    
    st.subheader("🎚️ Seuil de confiance")
    st.session_state.seuil_confiance = st.slider(
        "Seuil pour classement automatique (%)",
        min_value=30, max_value=95,
        value=st.session_state.seuil_confiance,
        help="En dessous, l'app vous demande de valider"
    )
    
    st.markdown("---")
    st.subheader("📁 Dossiers")
    c1, c2, c3 = st.columns(3)
    with c1: st.text_input("Classement", value=DOSSIER_TRI, disabled=True, key="p_d1")
    with c2: st.text_input("Export", value=DOSSIER_EXCEL, disabled=True, key="p_d2")
    with c3: st.text_input("Filtres", value=DOSSIER_FILTRES, disabled=True, key="p_d3")
    
    st.markdown("---")
    st.subheader(f"🏷️ {len(CATEGORIES)-1} domaines configurés")
    nb_metiers = sum(len(data.get('metiers', {})) for data in CATEGORIES.values())
    st.info(f"📊 **{len(CATEGORIES)-1}** domaines et **{nb_metiers}** métiers")
    
    for cat, data in CATEGORIES.items():
        if cat == 'Non classé': continue
        metiers = data.get('metiers', {})
        with st.expander(f"{data['emoji']} {cat} ({len(metiers)} métiers)"):
            if metiers:
                for metier, mots in metiers.items():
                    st.markdown(f"**{metier}** : {', '.join(mots[:5])}")
    
    st.markdown("---")
    st.subheader("ℹ️ À propos")
    st.markdown(f"""
    **CVSmart v11.0** — L'intelligence qui trie vos talents
    
    - 📄 Support : PDF, DOCX, Scans (OCR)
    - 🎯 **{len(CATEGORIES)-1}** domaines et **{nb_metiers}** métiers
    - 📊 Extraction de 25+ champs
    - 🎚️ Seuil de confiance configurable
    - 🧠 Apprentissage utilisateur (JSON)
    - ⚠️ Validation manuelle si confiance faible
    - 🔍 Filtrage multi-critères
    - 🤝 Matching avec offres
    - 📁 Export dossiers + Excel
    - ✅ "RAS" si info absente
    """)


# ============================================================
# MAIN
# ============================================================
def main():
    inject_css()
    init_session()
    page = afficher_menu()
    
    if page == "Accueil": page_accueil()
    elif page == "Emails": page_emails()
    elif page == "Analyser": page_analyser()
    elif page == "Résultats": page_resultats()
    elif page == "Filtrer": page_filtrer()
    elif page == "Validation": page_validation()
    elif page == "Apprentissage": page_apprentissage()
    elif page == "Matching": page_matching()
    elif page == "Paramètres": page_parametres()


if __name__ == "__main__":
    main()