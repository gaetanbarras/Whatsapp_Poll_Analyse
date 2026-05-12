# Suivi de sondages WhatsApp

Application Streamlit locale pour rapprocher une liste de participants avec un export de votes WhatsApp, normaliser les numéros suisses, suivre les relances manuelles et produire un rapport PDF simple.

## Fonctionnalités

- import d'un fichier de référence en Excel ou CSV
- import d'un export WhatsApp au format CSV
- détection automatique des colonnes WhatsApp fixes (`Name`, `Phone`) et des options du sondage
- mapping manuel des colonnes du fichier de référence
- normalisation robuste des numéros suisses vers le format `41791234567`
- sauvegarde locale de la dernière session pour éviter de réimporter les fichiers à chaque ouverture
- base SQLite locale pour mémoriser la liste des membres et les actions de suivi
- gestion des doublons de téléphone côté référence et côté WhatsApp
- colonne `Invité` persistante pour le suivi des personnes à compter
- mise à jour manuelle d'un participant en `Présent`, `Absent`, `Autre` ou `Pas de réponse`
- ajout des colonnes `Sexe` et `Fonction / Instrument` dans les résultats
- deux modes d'analyse :
  - `Analyse libre` : conserve les réponses brutes sans interprétation métier
  - `Analyse catégorisée` : classe les réponses en `Présent`, `Absent`, `Autre` ou `Conflit`
- exports Excel multi-onglets, CSV principal et rapport PDF simple
- interface en français avec messages d'aide, d'alerte et d'erreur

## Installation

Prérequis :

- Python 3.11 ou plus récent

Installation des dépendances :

```bash
pip install -r requirements.txt
```

## Lancement

Depuis le dossier du projet :

```bash
streamlit run app.py
```

L'application s'ouvre ensuite dans votre navigateur local.

## Protection par mot de passe

L'application attend un secret Streamlit nommé `app_password`.

En local, créez le fichier `.streamlit/secrets.toml` :

```toml
app_password = "choisissez-un-mot-de-passe-solide"
```

Pour un déploiement sur Streamlit Community Cloud, ajoutez la même clé dans l'interface **App settings > Secrets** :

```toml
app_password = "choisissez-un-mot-de-passe-solide"
```

Le fichier `.streamlit/secrets.toml` ne doit pas être versionné.

## Formats attendus

### 1. Fichier de référence

Format recommandé :

- fichier Excel `.xlsx`

Formats aussi acceptés :

- `.xls`
- `.xlsm`
- `.csv`

Colonnes possibles dans le fichier :

- prénom
- nom
- nom complet
- téléphone / portable / mobile
- sexe
- fonction / instrument

Le mapping est fait dans l'interface. La colonne téléphone est obligatoire.

### 2. Export WhatsApp

Format attendu :

- fichier `.csv`

Structure typique :

- une colonne `Name`
- une colonne `Phone`
- une ou plusieurs colonnes d'options de sondage

Dans les colonnes de réponse, un `X` ou `x` indique une option sélectionnée. Les espaces autour du marqueur sont tolérés.

## Fonctionnement

1. importer le fichier de référence et le CSV WhatsApp, ou laisser l'application recharger la dernière session locale
2. vérifier les aperçus de fichiers
3. choisir le mapping des colonnes de la référence, y compris `Sexe` et `Fonction` si disponibles
4. vérifier ou corriger le mapping WhatsApp
5. choisir le mode d'analyse
6. si besoin, configurer quelles options signifient `Présent`, `Absent` ou `Autre`
7. lancer le traitement
8. utiliser la section `Suivi manuel et relance` pour :
   - saisir une réponse brute manuelle
   - marquer une décision manuelle
   - cocher la colonne `Invité`
   - saisir une note de suivi
9. exporter le résultat en Excel, CSV ou PDF

## Persistance locale

L'application crée une base SQLite locale dans `.app_data/whatsapp_poll.db`.

Cette base conserve :

- la dernière session importée, avec les fichiers et réglages principaux
- la liste des membres synchronisée depuis le fichier de référence
- les décisions manuelles de suivi
- la colonne `Invité`
- les commentaires de suivi

Sans nouveau fichier importé, la dernière session est rechargée automatiquement au prochain lancement.

## Résultats produits

Le tableau principal contient notamment :

- prénom
- nom
- nom complet
- sexe
- fonction
- invité
- téléphone original et normalisé côté référence
- nom WhatsApp
- téléphone original et normalisé côté WhatsApp
- réponses brutes
- catégorie automatique
- catégorie manuelle
- catégorie finale
- statut de réponse
- commentaire de suivi
- commentaire d'anomalie

Les exports Excel contiennent les onglets :

- `Tous`
- `Presents`
- `Absents`
- `Autres`
- `Sans_reponse`
- `Inconnus`
- `Anomalies`
- `Synthese_options`

Le rapport PDF contient :

- une synthèse par option de réponse cochée avec le nombre de votes
- la liste des présents avec nom, prénom, réponse choisie, fonction et indicateur invité
- la liste des absents avec la réponse choisie
- la liste des non-répondants
- une synthèse par fonction avec le nombre de présents et le nombre d'invités cochés

## Règles de normalisation des téléphones

Les numéros sont normalisés ainsi :

- suppression des espaces, tirets, parenthèses et caractères spéciaux
- conversion de `+41`, `0041`, `41`, `0` ou d'un numéro local à 9 chiffres vers un format unique
- format interne final : `41791234567`

Les numéros invalides ne bloquent pas l'application. Ils sont signalés dans les résultats et dans l'onglet des anomalies.

## Limites connues

- seule la première feuille d'un fichier Excel est lue
- la persistance est locale au poste sur lequel l'application tourne
- les numéros sont normalisés pour les cas suisses courants uniquement
- les modifications manuelles sont rattachées à un membre via son identité et son téléphone ; si l'identité change fortement entre deux imports, une nouvelle fiche membre peut être créée
- si le CSV WhatsApp contient des colonnes supplémentaires non prévues, elles seront considérées comme des options de sondage tant qu'elles ne sont pas choisies comme nom ou téléphone

## Idées d'amélioration

- ajout de sessions nommées au lieu d'une seule session locale
- import de plusieurs sondages dans une même base locale
- prise en charge de plusieurs feuilles Excel
- mémorisation de modèles de mapping par type de fichier
- règles de normalisation pour d'autres pays
- tableaux de bord graphiques pour les relances et les présences
