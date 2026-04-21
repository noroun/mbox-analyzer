# 📧 Zoquer GMail

> ⚠️ **Cet outil est uniquement à titre informatif.** Le plus important est d'avoir **activé la redirection** et de **suivre les instructions de l'équipe gadz.org et de vos DDP**. Consultez ces instructions en priorité — cet outil vient en complément, pas en remplacement.

**Identifiez tous les sites où vous avez un compte avant de fermer votre boîte Gmail.**

> 💡 Cet outil vient en **complément des instructions de l'équipe gadz.org et de vos DDP**. Utilisez-le pour recenser vos comptes en ligne avant de procéder à la migration de votre adresse.

Cet outil analyse l'export de votre boîte mail Gmail et génère un **fichier Excel** (`resultats.xlsx`) avec trois onglets :
- ✅ **Comptes détectés** — sites où vous avez probablement créé un compte
- 📰 **Newsletters** — abonnements avec lien de désinscription cliquable
- 📋 **Tous les expéditeurs** — liste complète pour vérification manuelle

---

## 🔒 Confidentialité

✅ **Tout est traité localement sur votre ordinateur.** Aucune donnée n'est envoyée sur Internet.
✅ Le code est ouvert, vous pouvez le vérifier dans `mbox_analyzer.py`.
✅ L'outil ne lit que votre fichier MBOX, ne se connecte à rien.

---

## 🚀 Tutoriel pas à pas (aucune connaissance technique requise)

### Étape 1 — Téléchargez l'outil

Rendez-vous sur la page **Releases** de ce projet et téléchargez :
- **Windows** → `ZoquerGMail-Windows.zip`
- **Mac** → `ZoquerGMail-macOS.zip`

Décompressez le fichier ZIP en double-cliquant dessus.

### Étape 2 — Exportez votre Gmail

1. Allez sur **https://takeout.google.com**
2. Cliquez sur **« Tout désélectionner »** en haut
3. Faites défiler la liste et cochez uniquement **« Messagerie »** (Mail en anglais)
4. Tout en bas, cliquez sur **« Étape suivante »**
5. Choisissez :
   - Méthode de livraison : **Envoyer le lien de téléchargement par e-mail**
   - Type d'export : **Une seule exportation**
   - Type de fichier : **.zip**
   - Taille maximale : **2 Go** (suffisant dans la plupart des cas)
6. Cliquez sur **« Créer l'exportation »**

⏱️ Google met de quelques heures à plusieurs jours à préparer votre archive. Vous recevrez un email avec le lien de téléchargement quand c'est prêt.

### Étape 3 — Récupérez le fichier MBOX

1. Téléchargez le ZIP que Google vous a envoyé
2. Décompressez-le
3. Vous trouverez un fichier `.mbox` dans le dossier `Takeout/Messagerie/` — c'est lui qu'il faut analyser

### Étape 4 — Lancez l'analyse

#### Sur Mac
1. Double-cliquez sur **ZoquerGMail.app**
2. ⚠️ Si Mac affiche *« Apple ne peut pas vérifier que cette app ne contient pas de logiciel malveillant »* :
   - Faites **clic droit** sur l'app → **Ouvrir** → **Ouvrir** dans la fenêtre de confirmation
   - Vous n'aurez à le faire qu'une seule fois

#### Sur Windows
1. Double-cliquez sur **ZoquerGMail.exe**
2. ⚠️ Si Windows affiche *« Windows a protégé votre ordinateur »* :
   - Cliquez sur **« Informations complémentaires »**
   - Puis sur **« Exécuter quand même »**

### Étape 5 — Utilisez l'outil

1. Cliquez sur **« Parcourir »** à côté de *« Fichier MBOX »* et sélectionnez votre fichier `.mbox`
2. Le dossier de sortie est rempli automatiquement (vous pouvez le changer)
3. Cliquez sur **« Lancer l'analyse »**
4. ⏱️ Patientez : selon la taille de votre boîte, ça peut prendre de quelques minutes à 30 minutes

### Étape 6 — Exploitez les résultats

L'outil génère un fichier **`resultats.xlsx`** ouvrable dans Excel, Numbers, ou Google Sheets. Il contient 3 onglets :

| Onglet | Contenu | À quoi ça sert |
|---|---|---|
| **Comptes détectés** | Sites où vous avez un compte | **Changer l'email de connexion** sur chacun |
| **Newsletters** | Abonnements + lien de désinscription cliquable | **Vous désinscrire** d'un clic |
| **Tous les expéditeurs** | Tous les expéditeurs uniques | Vérification manuelle si besoin |

💡 **Astuce** : chaque onglet a des filtres activés — cliquez sur les flèches à côté des en-têtes pour trier ou filtrer.

### Conseils pour la migration

> **1. Activez la redirection en premier.** Avant de commencer à changer vos adresses sur les sites, configurez la redirection de votre ancienne vers votre nouvelle adresse : **Paramètres Gmail → Transfert et POP/IMAP**. Vous continuerez ainsi à recevoir les mails pendant toute la migration, même si vous oubliez un compte.

> **2. Priorisez par ordre d'importance.**
> Administratif (impôts, CAF…) → Banque / Finance → Pro (LinkedIn, GitHub, etc.) → Réseaux sociaux → Loisirs → Reste.

> **3. Acceptez de perdre certains comptes.** Les services morts ou sans intérêt n'ont pas besoin d'être traités — c'est tout à fait normal de les laisser tomber.

---

## 🔍 Compléter avec deux pages Google indispensables

L'analyse du MBOX rate par construction deux types de comptes : ceux dont vous avez supprimé les mails de bienvenue, et ceux ouverts via **« Se connecter avec Google »** (SSO/OAuth) qui ne laissent souvent aucune trace dans la boîte. Pour les retrouver, ouvrez ces deux pages dans votre navigateur (connecté à votre compte Google) :

- 🔑 **Mots de passe enregistrés** → [https://passwords.google.com](https://passwords.google.com)
  La liste de tous les sites pour lesquels Chrome / Android a stocké un mot de passe. Recoupez avec l'onglet *Comptes détectés* — tout site présent ici et absent du XLSX est un compte à traiter.

- 🔗 **Applications connectées via Google** → [https://myaccount.google.com/connections](https://myaccount.google.com/connections)
  La liste des services tiers connectés via *Sign in with Google* (Spotify, Notion, Canva, Figma, etc.). C'est la seule façon fiable de retrouver les comptes SSO, **invisibles dans les emails**.

Pour chaque entrée trouvée sur ces deux pages : changez l'email/identifiant de connexion vers votre nouvelle adresse, puis révoquez l'accès Google si vous fermez votre compte Gmail.

---

## ❓ FAQ

**Q : Combien de temps prend l'analyse ?**
R : Environ 1 minute par 10 000 mails. Une boîte de 50 000 mails ≈ 5 minutes.

**Q : Mes données sont-elles envoyées quelque part ?**
R : Non, jamais. Tout reste sur votre ordinateur.

**Q : Pourquoi certains comptes ne sont pas détectés ?**
R : L'outil détecte les comptes via les mails d'inscription/bienvenue/réinitialisation. Si vous avez supprimé ces mails, le compte n'apparaîtra que dans `tous_expediteurs.csv`.

**Q : Et pour Yahoo / Outlook / autre ?**
R : Cet outil fonctionne avec n'importe quel fichier MBOX. Pour Outlook, exportez d'abord en MBOX via Thunderbird.

**Q : À quoi ça sert ?**
R : À queud's.

---

## 🛠️ Pour les développeurs

### Lancer depuis les sources

```bash
git clone https://github.com/VOTRE_PSEUDO/mbox-analyzer
cd mbox-analyzer
pip install -r requirements.txt
python3 mbox_analyzer.py
```

Seule dépendance externe : `openpyxl` (pour générer le fichier Excel formaté).

### Compiler localement

```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --onefile --windowed --name ZoquerGMail mbox_analyzer.py
```

L'exécutable apparaît dans `dist/`.

### Compilation automatique via GitHub Actions

À chaque push d'un tag `v*` (ex: `v1.0.0`), GitHub compile automatiquement les versions Mac et Windows et crée une Release.

```bash
git tag v1.0.0
git push origin v1.0.0
```

Vous pouvez aussi lancer manuellement le workflow depuis l'onglet **Actions** de GitHub.

---

## 📄 Licence

MIT — utilisez, modifiez et partagez librement.
