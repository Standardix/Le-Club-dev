# README – Application de préparation des fichiers produits Shopify (Matrixify)

## 1. Objectif de l’application
Cette application permet de transformer un fichier produit fournisseur en un fichier prêt à importer dans Shopify via Matrixify, en appliquant automatiquement :

- des règles de standardisation
- des mappings (couleurs, tailles, genres, types de produits)
- des contrôles de cohérence
- des validations visuelles pour faciliter la vérification finale

L’objectif est de réduire les erreurs d’import, uniformiser les données et accélérer la mise en ligne.

---

## 2. Fichiers requis
Pour fonctionner correctement, l’application nécessite **trois fichiers** :

### 1) Fichier fournisseur
- Fichier source contenant les produits (Excel ou CSV)
- Les noms de colonnes peuvent varier selon le fournisseur

### 2) Help Data
- Fichier de référence utilisé pour :
  - la standardisation des couleurs
  - la standardisation des tailles
  - la logique de genre / non genré
  - les types de produits
- Ce fichier est utilisé en **lecture seulement** et n’est jamais modifié par l’application
- ⚠️ Ce fichier ne doit pas être modifié sans validation du responsable de **Le Club Boutique**

### 3) Export Shopify (produits existants)
- Export des produits Shopify existants
- Utilisé pour :
  - détecter les produits déjà présents
  - éviter les doublons à l’import
  - appliquer les règles **do not import** lorsque requis
- Ce fichier est utilisé en **lecture seulement** et n’est jamais modifié par l’application

---

## 3. Fonctionnement général
1. Le fichier fournisseur est analysé.
2. Pour chaque colonne Shopify :
   - la meilleure source est détectée automatiquement
   - les règles définies sont appliquées
   - les données sont nettoyées et standardisées
3. Un fichier final Shopify est généré, prêt à l’import.

---

## 4. Comprendre les validations visuelles

### 🔴 Cellules en rouge
Une cellule en rouge indique :
- une donnée ambigüe, non conforme ou à valider manuellement
- exemples :
  - couleur non reconnue
  - produit possiblement unisexe sans indication claire
  - titre contenant des caractères problématiques (`?`, `/`)

👉 **Action requise avant import Shopify.**

---

### 🟡 Cellules en jaune
Une cellule en jaune indique :
- une donnée optionnelle mais recommandée
- exemples :
  - description marketing manquante
  - product features absentes

👉 **Non bloquant, mais conseillé de compléter.**

---

## 5. Points clés à vérifier avant l’import Shopify
Avant d’importer le fichier dans Shopify, assurez-vous que :

- il n’y a **aucune cellule rouge restante**
- les titres produits sont cohérents et lisibles
- les couleurs Google Shopping sont valides
- les produits NON genrés sont correctement identifiés
- les variantes sont correctement regroupées (même Handle)
- les informations critiques (prix, SKU, stock, etc.) sont présentes

---

## 6. À propos des variantes et du Handle
- le Handle est l’identifiant principal du produit
- toutes les variantes d’un même produit partagent le même Handle
- cela permet à Shopify de regrouper automatiquement les variantes sous un seul produit

---

## 7. Bonnes pratiques
- toujours utiliser la dernière version du fichier Help Data
- ne pas renommer les colonnes Shopify dans le fichier final
- corriger les cellules rouges avant l’import
- en cas de doute, se référer au document **Règles des colonnes**, qui fait foi et décrit, colonne par colonne :
  - la source des données
  - les règles appliquées
  - les validations

---

## 8. Support et évolution
Cette application repose sur des règles documentées et évolutives.  
Toute modification de logique (nouvelle règle, nouveau fournisseur, nouveau mapping) doit être validée afin de garantir la cohérence des imports.
