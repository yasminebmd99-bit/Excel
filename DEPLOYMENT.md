# Guide de Déploiement - Excel Manager

## 📋 Prérequis

- Un compte GitHub (gratuit) : https://github.com
- Un compte Vercel (gratuit) : https://vercel.com
- Git installé sur votre ordinateur

---

## 🚀 Étape 1 : Préparer le projet

### 1.1 Créer un fichier `.gitignore`

Créez un fichier `.gitignore` dans `c:\Users\DELL\Desktop\EXCEL\` avec ce contenu :

```
# Fichiers système
.DS_Store
Thumbs.db

# Fichiers temporaires
*.tmp
*.log

# Dossiers inutiles
node_modules/
.vscode/
```

### 1.2 Vérifier que vous avez ces fichiers

Assurez-vous d'avoir :
- ✅ `index.html`
- ✅ `styles.css`
- ✅ `app.js`
- ✅ `README.md`
- ✅ `.gitignore` (nouveau)

---

## 📦 Étape 2 : Créer un dépôt GitHub

### 2.1 Initialiser Git localement

Ouvrez PowerShell dans le dossier `EXCEL` et exécutez :

```powershell
cd c:\Users\DELL\Desktop\EXCEL
git init
git add .
git commit -m "Initial commit - Excel Manager"
```

### 2.2 Créer le dépôt sur GitHub

1. Allez sur https://github.com
2. Cliquez sur le bouton **"+"** en haut à droite → **"New repository"**
3. Remplissez :
   - **Repository name** : `excel-manager` (ou le nom de votre choix)
   - **Description** : "Application web pour gérer et réorganiser des fichiers Excel"
   - **Public** ou **Private** : à votre choix
   - ⚠️ **NE COCHEZ PAS** "Add a README file" (vous en avez déjà un)
4. Cliquez sur **"Create repository"**

### 2.3 Lier votre projet local à GitHub

GitHub vous affichera des commandes. Copiez et exécutez dans PowerShell :

```powershell
git remote add origin https://github.com/VOTRE_USERNAME/excel-manager.git
git branch -M main
git push -u origin main
```

> Remplacez `VOTRE_USERNAME` par votre nom d'utilisateur GitHub

---

## 🌐 Étape 3 : Déployer sur Vercel

### 3.1 Créer un compte Vercel

1. Allez sur https://vercel.com
2. Cliquez sur **"Sign Up"**
3. Choisissez **"Continue with GitHub"**
4. Autorisez Vercel à accéder à votre compte GitHub

### 3.2 Importer votre projet

1. Sur le dashboard Vercel, cliquez sur **"Add New..."** → **"Project"**
2. Trouvez votre dépôt `excel-manager` dans la liste
3. Cliquez sur **"Import"**

### 3.3 Configurer le déploiement

Vercel détectera automatiquement que c'est un site statique. Vérifiez :

- **Framework Preset** : Other (ou None)
- **Root Directory** : `./` (laisser par défaut)
- **Build Command** : (laisser vide)
- **Output Directory** : (laisser vide)

Cliquez sur **"Deploy"** 🚀

### 3.4 Attendre le déploiement

Vercel va :
1. Cloner votre dépôt
2. Déployer les fichiers
3. Vous donner une URL (ex: `excel-manager.vercel.app`)

⏱️ Cela prend environ 30 secondes.

---

## ✅ Étape 4 : Tester votre site en ligne

Une fois le déploiement terminé :

1. Vercel affichera votre URL : `https://excel-manager-xxx.vercel.app`
2. Cliquez dessus pour ouvrir votre site
3. Testez toutes les fonctionnalités :
   - Import de fichiers
   - Fusion
   - Réorganisation
   - Export

---

## 🔄 Mettre à jour votre site

Chaque fois que vous modifiez votre code :

```powershell
cd c:\Users\DELL\Desktop\EXCEL
git add .
git commit -m "Description de vos modifications"
git push
```

Vercel redéploiera **automatiquement** votre site ! 🎉

---

## 🎨 Personnaliser le domaine (Optionnel)

### Option 1 : Domaine Vercel gratuit

Vercel vous donne un domaine gratuit : `votre-projet.vercel.app`

Vous pouvez le personnaliser dans les settings du projet.

### Option 2 : Votre propre domaine

Si vous avez un domaine (ex: `monsite.com`) :

1. Allez dans **Settings** → **Domains**
2. Ajoutez votre domaine
3. Suivez les instructions pour configurer les DNS

---

## 📝 Commandes Git utiles

```powershell
# Voir le statut de vos fichiers
git status

# Voir l'historique des commits
git log --oneline

# Annuler les modifications non commitées
git checkout .

# Créer une nouvelle branche
git checkout -b nouvelle-fonctionnalite

# Revenir à la branche principale
git checkout main
```

---

## 🆘 Problèmes courants

### Problème : "git: command not found"

**Solution** : Installez Git depuis https://git-scm.com/download/win

### Problème : Erreur d'authentification GitHub

**Solution** : Utilisez un Personal Access Token :
1. GitHub → Settings → Developer settings → Personal access tokens
2. Generate new token (classic)
3. Utilisez ce token comme mot de passe

### Problème : Le site ne se met pas à jour

**Solution** : 
1. Vérifiez que vous avez bien fait `git push`
2. Allez sur le dashboard Vercel → Deployments
3. Vérifiez que le dernier déploiement est réussi

---

## 🎉 Félicitations !

Votre application Excel Manager est maintenant en ligne et accessible partout dans le monde ! 🌍

**URL de votre site** : `https://excel-manager-xxx.vercel.app`

Partagez-le avec vos collègues et amis ! 🚀
