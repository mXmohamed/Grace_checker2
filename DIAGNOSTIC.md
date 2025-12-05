# Guide de Diagnostic GRACE Analyzer 🔍

## 🚨 Problèmes connus et solutions

### 1. Toutes les erreurs sont classées "AUTRES"

#### Diagnostic :
1. **Ouvrir la console** (F12 → Console)
2. **Charger vos fichiers** et analyser
3. **Chercher ces logs** :
   ```
   Classification du code: [CODE_ERREUR]
   Code en majuscules: [CODE_MAJUSCULES]
   Familles disponibles: [LISTE_FAMILLES]
   ```

#### Causes possibles :
- **Codes d'erreur différents** de ceux mappés (ex: `BPE_25_B` au lieu de `BPE_25_A`)
- **Format inattendu** (ex: espaces, caractères spéciaux)
- **Colonne d'erreur incorrecte** (pas la colonne D)

#### Solutions :
1. **Vérifier les codes** dans la console
2. **Comparer** avec la liste intégrée dans le code
3. **Adapter le mapping** si nécessaire

### 2. En-têtes non visibles dans l'interface web

#### Diagnostic :
1. **Console** → Chercher :
   ```
   En-têtes capturés: [LISTE_EN_TETES]
   csvHeaders disponibles: [HEADERS]
   En-têtes utilisés: [HEADERS_FINAUX]
   ```

#### Causes possibles :
- **Séparateur CSV incorrect** (`;` vs `,` vs `\t`)
- **En-tête non capturé** du premier fichier
- **Moins de 15 colonnes** dans les CSV

#### Solutions :
1. **Vérifier le séparateur** de vos CSV
2. **S'assurer** que le premier CSV a bien des en-têtes
3. **Contrôler** que vos CSV ont ≥15 colonnes

## 📊 Format des données attendu

### Structure CSV :
```
DSP;REF_ERREUR;NATURE_ERREUR;CODE_ERREUR;DESCRIPTION;...
WIGA;ERR_001;Erreur BPE;BPE_25_A;Description...
```

### Codes d'erreur supportés :
- **BPE** : BPE_24_A, BPE_25_A, BPE_77_A, etc.
- **CABLE** : CABLE_1_C, CABLE_71_B, OPTCABLE_113_A, etc.
- **ARMOIRE** : ARMOIRE_37_A, ARMOIRE_86_B, etc.
- **Etc.** : 369 codes total

## 🛠️ Actions de debug

### 1. Test rapide :
```javascript
// Dans la console du navigateur
console.log(window.graceApp.errorCodeToFamily['BPE_25_A']);
console.log(window.graceApp.csvHeaders);
```

### 2. Lister vos codes d'erreur :
```javascript
// Voir tous les codes uniques trouvés
const codes = [...new Set(window.graceApp.allData.map(r => r.errorCode))];
console.log('Codes trouvés:', codes.slice(0, 10));
```

### 3. Vérifier la classification :
```javascript
// Tester la classification d'un code
window.graceApp.classifyError('VOTRE_CODE_ICI');
```

## 📞 Support

Si les problèmes persistent :
1. **Copier les logs** de la console
2. **Fournir un exemple** de fichier CSV (anonymisé)
3. **Indiquer** le navigateur utilisé

## 🔧 Correction rapide

Si vos codes ne sont pas dans la liste, vous pouvez temporairement les ajouter :
```javascript
// Ajouter manuellement un code (dans la console)
window.graceApp.errorCodeToFamily['VOTRE_CODE'] = 'FAMILLE_SOUHAITEE';
```