# Triage_Outlook 📥🤖

**Triage_Outlook** est un complément (Add-in) pour **Microsoft Outlook** (Web, Mac et Windows) propulsé par l'IA (Google Gemini ou Anthropic Claude), conçu pour **trier et classer automatiquement des lots d'e-mails** dans les bons sous-dossiers selon vos propres règles métier.

---

## ✨ Fonctionnalités principales

1. **Sélection Multiple (*Multi-Select*) :**
   - Sélectionnez 5, 15, 50 e-mails d'un coup dans votre boîte de réception.
   - L'Add-in charge et prépare l'analyse du lot complet.

2. **Dossiers personnalisés & Règles en français naturel :**
   - L'Add-in peut synchroniser automatiquement l'arborescence réelle de vos dossiers Outlook.
   - Cochez les dossiers actifs pour le tri.
   - Pour chaque dossier, définissez simplement la consigne en français :
     - *Exemple : « Mails des clients qui souhaitent commander un nouveau lecteur, des accessoires ou un devis matériel. »*
     - *Exemple : « Erreurs de télétransmission, blocages logiciels, pannes de lecteur et urgences techniques. »*

3. **Classification intelligente par IA :**
   - L'IA analyse le sujet et le contenu de chaque message, attribue le meilleur dossier et donne une justification concise (5 à 10 mots).
   - Possibilité de réassigner manuellement un dossier en cas de doute avant de valider.

4. **Déplacement physique dans Outlook :**
   - Un clic sur **« Déplacer les messages »** et chaque e-mail quitte la boîte de réception pour atterrir dans son dossier de destination via l'API native Exchange (`MoveItem`).

5. **Partage d'équipe :**
   - Boutons **Exporter** / **Importer** pour partager la configuration des règles entre collègues en 1 clic (format JSON).

6. **Sécurité d'accès :**
   - Protection optionnelle par code secret (`AUTH_TOKEN`) pour éviter tout usage non autorisé de vos quotas d'IA.

---

## 🚀 Déploiement sur Railway

1. Rendez-vous sur votre tableau de bord [Railway](https://railway.app).
2. Cliquez sur **New Project** → **Deploy from GitHub repo** → sélectionnez **`Triage_Outlook`**.
3. Dans l'onglet **Variables** du service :
   - `GEMINI_API_KEY` : votre clé Google AI (ou `ANTHROPIC_API_KEY`).
   - `AUTH_TOKEN` *(Recommandé)* : le code secret de votre choix pour protéger l'accès.
4. Railway génère automatiquement votre domaine sécurisé HTTPS.

---

## 📥 Installation dans Outlook

1. Rendez-vous sur `https://<votre-app>.up.railway.app/manifest.xml?download=1` pour télécharger votre fichier de configuration.
2. Dans Outlook (sur le Web ou bureau) :
   - Ouvrez un e-mail → **Compléments** (ou `...` > *Applications*).
   - Allez dans **Mes compléments** (*My Add-ins*).
   - En bas, dans **Compléments personnalisés**, cliquez sur **« + Ajouter un complément personnalisé »** → **« Ajouter à partir d'un fichier... »**.
   - Sélectionnez le fichier `manifest.xml`.
3. Le bouton **Trier les e-mails** apparaît dans votre ruban !

---

## 🔒 Sécurité & Confidentialité

- **Aucun stockage en base de données :** Les messages ne sont ni enregistrés ni stockés sur le serveur.
- **Règles locales :** Les dossiers et leurs règles sont enregistrés dans votre navigateur (`localStorage`).
- **Standard Microsoft :** Utilise les permissions officielles `ReadWriteMailbox` pour le déplacement des éléments.
