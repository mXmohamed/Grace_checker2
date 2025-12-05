# GRACE Analyzer 📊

**Analyseur d'Erreurs Multi-Périodes pour les rapports PM**

Une application web moderne pour analyser et comparer les erreurs entre différentes versions de rapports de maintenance préventive.

## 🚀 Fonctionnalités

### 📋 Analyse Multi-Périodes
- **Chargement de fichiers ZIP** multiples contenant des rapports CSV
- **Classification automatique** des erreurs par famille (BPE, CABLE, ARMOIRE, etc.)
- **Comparaison temporelle** entre anciennes et nouvelles versions
- **Extraction automatique** des DSP et informations PM

### 📊 Visualisation
- **Graphiques interactifs** avec Chart.js
- **Comparaison visuelle** ancienne vs nouvelle version
- **Filtrage avancé** par PM, famille d'erreur, code d'erreur
- **Évolution temporelle** des erreurs

### 📑 Export Excel
- **Rapport Excel complet** avec 3 feuilles :
  - Données brutes (colonnes A-O)
  - Comparatif temporel avec pourcentages d'évolution
  - Graphiques natifs Excel par PM
- **Graphiques Excel natifs** avec couleurs cohérentes
- **Formatage professionnel** prêt pour impression

## 🛠️ Technologies

- **Frontend** : HTML5, Bootstrap 5, JavaScript ES6
- **Graphiques** : Chart.js pour web, graphiques Excel natifs
- **Traitement fichiers** : JSZip, PapaParse, ExcelJS
- **Style** : Bootstrap Icons, CSS3 personnalisé

## 📱 Utilisation

### 1. Chargement des données
1. **Sélectionner** les fichiers ZIP contenant les rapports d'erreurs
2. **Optionnel** : Charger un fichier de référence pour classification
3. **Cliquer** sur "Analyser les fichiers"

### 2. Filtrage et sélection
1. **Choisir** les PM à analyser (Ctrl+Click pour sélection multiple)
2. **Sélectionner** les familles d'erreurs à inclure
3. **Cliquer** sur "Analyser et visualiser"

### 3. Visualisation des résultats
- **Onglet 1** : Données brutes avec vrais en-têtes de colonnes
- **Onglet 2** : Comparaison temporelle avec évolution (%)
- **Onglet 3** : Graphiques interactifs avec filtres avancés

### 4. Export Excel
- **Cliquer** sur "Exporter en Excel"
- **Télécharger** le rapport complet avec graphiques natifs

## 🎨 Caractéristiques des graphiques

### Couleurs cohérentes
- **🔲 Gris** : Version ancienne
- **🔵 Bleu** : Version nouvelle  
- **🟢 Vert, 🟡 Jaune, 🟣 Violet** : Versions intermédiaires

### Fonctionnalités avancées
- **Ordonnancement** par nombre d'erreurs décroissant
- **Affichage des zéros** (familles passant de 5→0 ou 0→5)
- **Dates réelles** dans les légendes
- **Filtrage dynamique** par PM/famille/code

## 📂 Structure des fichiers

```
GRACE_CHECKER/
├── index.html          # Interface principale
├── app.js             # Logique de l'application
├── style.css          # Styles personnalisés
└── README.md          # Documentation
```

## 🔧 Configuration

### Familles d'erreurs intégrées
L'application inclut **369 codes d'erreur** pré-configurés :
- ARMOIRE, BAIE, BPE, CABLE, CASSETTE
- CHAMBRE, IPE, JUMPER, MODULE, OPTRECEIVER
- PBO, PIGTAIL, PORTEE_AERIENNE, POTEAU, PRISE
- ROOM, SITE, TIROIR, TRANCHEE, ZAPM, ZNRO
- AUTRES (pour codes non référencés)

### Format des fichiers supportés
- **Fichiers ZIP** : Contenant des CSV et fichiers de synthèse
- **Fichiers CSV** : Avec en-têtes et au moins 15 colonnes
- **Fichiers de synthèse** : Format `RAPPORT_SYNTHESE_GV3_SRO-*`

## 🌐 Déploiement

### GitHub Pages
```bash
# Cloner et déployer
git clone https://github.com/[username]/grace-analyzer.git
cd grace-analyzer
# Les fichiers sont prêts pour GitHub Pages
```

### Serveur web local
```bash
# Python 3
python -m http.server 8000

# Node.js
npx serve .
```

## 📋 Prérequis navigateur

- **Chrome/Edge** : Version 80+
- **Firefox** : Version 75+
- **Safari** : Version 13+
- **JavaScript** activé
- **Support HTML5** File API

## 🤝 Contribution

Les contributions sont les bienvenues ! Merci de :
1. Fork le projet
2. Créer une branche feature
3. Commit les changements
4. Ouvrir une Pull Request

## 📄 Licence

Ce projet est sous licence MIT. Voir le fichier LICENSE pour plus de détails.

## 📞 Support

Pour toute question ou problème :
- Ouvrir une issue sur GitHub
- Documentation complète dans le code source

---

**GRACE Analyzer** - Simplifie l'analyse des erreurs multi-périodes ⚡