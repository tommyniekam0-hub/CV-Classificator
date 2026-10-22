import easyocr
import fitz
import os
import time

# ============================================================
# CONFIGURATION
# ============================================================

# ⚠️ Dossier contenant vos CV (changez si besoin)
DOSSIER_CV = r"C:\Users\HP\Desktop\CV A TRAITER"

# ============================================================
# CHARGEMENT DU MODÈLE
# ============================================================
print("=" * 70)
print("CHARGEMENT DU MODÈLE EASYOCR...")
print("=" * 70)

debut = time.time()
reader = easyocr.Reader(['fr', 'en'], gpu=False)
print(f"✅ Modèle chargé en {time.time() - debut:.1f} secondes\n")

# ============================================================
# RECHERCHE AUTOMATIQUE DU 1er CV
# ============================================================
if not os.path.isdir(DOSSIER_CV):
    print(f"❌ Dossier introuvable : {DOSSIER_CV}")
    print(f"💡 Vérifiez le chemin. Dossiers disponibles sur le Bureau :")
    bureau = r"C:\Users\HP\Desktop"
    for f in os.listdir(bureau):
        print(f"   - {f}")
    exit(1)

# Lister les CV
fichiers = [f for f in os.listdir(DOSSIER_CV) 
            if f.lower().endswith(('.pdf', '.docx', '.png', '.jpg', '.jpeg'))]

if not fichiers:
    print(f"❌ Aucun CV trouvé dans {DOSSIER_CV}")
    print(f"💡 Formats acceptés : PDF, DOCX, PNG, JPG")
    exit(1)

print(f"📁 Dossier : {DOSSIER_CV}")
print(f"📄 {len(fichiers)} fichier(s) trouvé(s) :")
for f in fichiers:
    print(f"   - {f}")

# Prendre le premier
chemin_cv = os.path.join(DOSSIER_CV, fichiers[0])

# ============================================================
# ANALYSE
# ============================================================
print("\n" + "=" * 70)
print(f"ANALYSE : {fichiers[0]}")
print("=" * 70)

try:
    doc = fitz.open(chemin_cv)
    print(f"\n✅ Fichier ouvert - {len(doc)} page(s)\n")
    
    for i, page in enumerate(doc):
        print(f"{'='*70}")
        print(f"=== PAGE {i+1} - TEXTE OCR ===")
        print(f"{'='*70}\n")
        
        pix = page.get_pixmap(dpi=200)
        img_bytes = pix.tobytes("png")
        
        debut = time.time()
        resultats = reader.readtext(img_bytes, detail=0)
        duree = time.time() - debut
        
        texte = "\n".join(resultats)
        print(texte)
        print(f"\n⏱️ Temps : {duree:.1f}s")
        print(f"📊 {len(resultats)} blocs détectés\n")
    
    doc.close()
    print("=" * 70)
    print("✅ ANALYSE TERMINÉE")
    print("=" * 70)
    
except Exception as e:
    print(f"\n❌ ERREUR : {e}")
    # Sauvegarder le texte extrait dans un fichier
with open("resultat_easyocr.txt", "w", encoding="utf-8") as f:
    f.write(texte)

print("\n✅ Résultat sauvegardé dans : resultat_easyocr.txt")