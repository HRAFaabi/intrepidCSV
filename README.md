# Fiches de transfert — Desert Evasion

App Streamlit : upload le fichier de réservations (.xlsx), choisis une période,
télécharge un PDF avec les Départs/Arrivées triés par heure, groupés par jour.

## Tester en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Ouvre ensuite `http://localhost:8501`.

## Héberger gratuitement sur Streamlit Community Cloud

1. Crée un dépôt GitHub (public ou privé) et mets-y ces 4 fichiers :
   `app.py`, `transfer_logic.py`, `requirements.txt`, `README.md`.
2. Va sur https://share.streamlit.io et connecte-toi avec ton compte GitHub.
3. Clique "New app", choisis ton dépôt, la branche (`main`), et le fichier
   principal : `app.py`.
4. Clique "Deploy". Au bout d'une à deux minutes, tu obtiens une URL du type
   `https://ton-app.streamlit.app` que tu peux partager avec ton père / les
   personnes qui gèrent les transferts.
5. À chaque `git push` sur le dépôt, l'app se redéploie automatiquement.

Streamlit Community Cloud est gratuit pour les apps publiques (le dépôt GitHub
peut rester privé si tu veux, Streamlit peut y accéder via son intégration
GitHub — seule l'URL de l'app reste accessible à qui tu la donnes).

## Alternative : Render (free tier)

Si tu préfères Render :
1. Crée un compte sur https://render.com et connecte ton dépôt GitHub.
2. "New Web Service" → choisis le dépôt.
3. Build command : `pip install -r requirements.txt`
4. Start command : `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
5. Déploie. Le free tier se met en veille après inactivité (le premier accès
   après une pause prend ~30-60 secondes à réveiller le service).

## Structure attendue du fichier de réservations

Colonnes nécessaires (comme dans le xlsx d'origine) :
`Current Status`, `Start Date`, `Component Name`, `Flight Time`, `Flight No`,
`Txn ID`, `Title`, `First Name`, `Surname`.

Seules les lignes avec `Current Status` = "confirmed" (peu importe la casse)
sont prises en compte.
