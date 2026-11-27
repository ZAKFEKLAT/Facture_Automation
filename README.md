# Automatisation des factures

Ce projet automatise la préparation et la génération des factures à partir de :

- un fichier **TRACKING EXPORT** ;
- un fichier **état des prix** ;
- un modèle de facture Excel ;
- des macros VBA intégrées dans le classeur `.xlsm`.

Le facturier n'a plus besoin de copier manuellement les MSN, les VU ou les prix. Il utilise le bouton du ruban, contrôle le résultat, puis récupère les factures générées.

## 1. Structure recommandée des fichiers

Les fichiers suivants doivent être placés dans le même dossier :

```text
Facture/
├── FACTURE.xlsm
├── TRACKING EXPORT 10-11-2026.xlsx
├── état des prix export a jour 12-2024.xlsx
└── Factures Generated/
```

Le nom exact des fichiers peut changer, mais les préfixes doivent rester compatibles avec les recherches automatiques :

```text
TRACKING EXPORT*.xlsx
prix*.xlsx
```

Le classeur facture doit contenir au minimum les feuilles :

- une feuille dont le nom contient `remplir` ;
- `Facture finale`.

Dans le modèle actuel, la feuille de saisie est `Infos à remplir`.

## 2. Installation initiale

### Importer les modules VBA

Dans Excel :

1. Ouvrir le fichier `.xlsm`.
2. Appuyer sur `Alt + F11`.
3. Dans l'éditeur VBA, choisir **File > Import File**.
4. Importer les modules :
   - `ModuleControleFactures.bas`
   - `ModuleGenerationFactures.bas`
   - `ModulePrixFacture.bas` si la génération spécifique LOT2 est conservée.
5. Enregistrer le classeur au format **Excel Macro-Enabled Workbook (`.xlsm`)**.

Un fichier `.xlsx` ne conserve pas les macros VBA.

### Configurer le bouton du ruban

Le bouton du ruban doit appeler la macro publique suivante :

```vb
InstallerControleFactures
```

Dans le ruban actuel :

```text
Facture > Invoices Controle
```

Lorsque l'utilisateur clique sur **Invoices Controle**, la macro crée ou actualise automatiquement la feuille :

```text
Controle Factures
```

## 3. Utilisation quotidienne

### Étape 1 — Préparer les fichiers

Placer dans le même dossier :

1. le fichier facture `.xlsm` ;
2. le nouveau `TRACKING EXPORT` ;
3. le nouvel état des prix.

Le tracking doit contenir une feuille dont le nom commence par `TRACKING EXPORT`.

Les colonnes attendues dans le tracking sont :

| Colonne | Donnée |
|---|---|
| A | programme |
| B | LOT |
| C | MSN |
| D | VU |
| F | Numéro article / code IS*M |

L'état des prix doit contenir une feuille dont le nom contient `prix de vente`.

Les colonnes utilisées sont :

| Colonne | Donnée |
|---|---|
| A | code IS*M |
| H | prix par VU |
| I | Prix de Facturation |

### Étape 2 — Actualiser le contrôle

Dans Excel :

```text
Facture > Invoices Controle
```

La macro :

- lit le tracking ;
- détecte les programmes et LOT présents ;
- crée une ligne par groupe de facture ;
- affiche le nombre de lignes concernées ;
- crée un bouton `Generate` pour chaque groupe ;
- crée un bouton `Generate all invoices`.

Si la feuille `Controle Factures` existe déjà, elle est reconstruite afin de refléter le dernier tracking.

### Étape 3 — Générer une facture

Dans `Controle Factures`, cliquer sur `Generate` sur la ligne souhaitée.

Une fenêtre demande le **coisage / packing list**, par exemple :

```text
CAS26-11-025ID
```

La valeur est écrite automatiquement dans `Facture finale!C4`. Le numéro de
facture de `B3` est ensuite construit avec la règle :

```text
LAT21 + jour de la date d'expédition en N4 + séquence numérique du coisage
```

Exemple :

```text
CAS26-11-025ID + date N4 = 10/11/2026
→ N° Facture : LAT211025
```

La macro :

1. copie le modèle facture ;
2. sélectionne les lignes du programme et du LOT ;
3. remplit les MSN et les VU ;
4. regroupe les lignes par MSN ;
5. construit les références ;
6. recherche les prix avec le code IS*M ;
7. remplit la colonne `Code Article` avec les IS*M uniques de chaque MSN ;
8. met à jour `Facture finale` ;
9. supprime la feuille de contrôle du fichier généré ;
10. sauvegarde la facture dans `Factures Generated`.

### Étape 4 — Générer toutes les factures

Cliquer sur :

```text
Generate all invoices
```

Une facture séparée est créée pour chaque groupe visible dans `Controle Factures`.

## 4. Résultats générés

Les fichiers sont créés dans :

```text
Factures Generated/
```

Exemples :

```text
A320-90VU_90VU.xlsx
A320-CCP_reliquat lot2.xlsx
WINGS_900.xlsx
WINGS_1000.xlsx
```

Si un fichier porte déjà le même nom, il est remplacé lors de la nouvelle génération.

## 5. Règle spéciale WINGS

Dans le tracking, WINGS peut contenir des LOT comme :

```text
1/900
2/900
3/900
1/1000
2/1000
42/1000
412/1000
```

Le nombre avant `/` correspond à une sous-variation ou à un groupe de VU. La facture doit cependant être regroupée par variation principale :

- tous les `X/900` donnent une seule facture `WINGS_900.xlsx` ;
- tous les `X/1000` donnent une seule facture `WINGS_1000.xlsx`.

Les VU individuels restent conservés dans le détail de la facture.

## 6. Règle de prix

Pour chaque code IS*M :

1. la macro utilise d'abord `Prix de Facturation` en colonne I ;
2. si ce prix est vide ou non numérique, elle utilise `prix par VU` en colonne H ;
3. si aucun prix n'est disponible, le résultat doit être contrôlé avant envoi.

### Produits vendus par lot complet

La macro analyse également les colonnes `Nature` (F) et `Remarque` (J).
Lorsqu'un MSN contient un article identifié comme `lot Complet`, les autres
articles du même MSN marqués `Vu non Facturable` ne sont pas additionnés.
Le MSN est facturé avec le prix de facturation du produit vendu par lot.

Exemple A320-90VU :

- VU 91 à 96 sont présents dans le tracking ;
- le VU 91 correspond au produit vendu par lot complet ;
- les autres VU sont non facturables individuellement ;
- le MSN est donc facturé une seule fois à **17 641 €**.

Cette règle est détectée automatiquement à partir de la nature et de la
remarque du fichier de prix, et ne dépend pas d'un code ou d'un numéro de VU
écrit en dur dans la macro.

L'état des prix est donc la source de vérité. Les anciennes grilles auxiliaires du modèle ne sont plus nécessaires pour calculer les nouveaux montants.

## 7. Contrôles à effectuer avant envoi

Après génération, vérifier dans `Facture finale` :

- le programme et le LOT ;
- les MSN ;
- les VU ;
- les codes article ;
- les quantités ;
- les prix unitaires ;
- le montant total ;
- la date d'expédition ;
- les informations commerciales et client.

La macro automatise la préparation, mais le contrôle métier final reste obligatoire.

## 8. Mise à jour des fichiers source

À chaque nouveau cycle :

1. conserver le même dossier de travail ;
2. remplacer ou ajouter le nouveau `TRACKING EXPORT` ;
3. remplacer ou ajouter le nouvel état des prix ;
4. éviter de laisser plusieurs fichiers correspondant aux mêmes motifs dans le dossier ;
5. ouvrir le fichier `.xlsm` ;
6. cliquer sur `Invoices Controle` ;
7. générer une facture ou toutes les factures.

Il est recommandé de déplacer les anciens exports dans un sous-dossier d'archive afin d'éviter qu'un mauvais fichier soit sélectionné automatiquement.

## 9. Dépannage

### La feuille `Controle Factures` n'apparaît pas

Vérifier que :

- les macros sont activées ;
- le fichier est bien `.xlsm` ;
- le bouton du ruban appelle `InstallerControleFactures` ;
- le tracking est dans le même dossier que le fichier facture.

### Aucun tracking n'est trouvé

Vérifier que le fichier commence par :

```text
TRACKING EXPORT
```

et possède une feuille dont le nom commence par `TRACKING EXPORT`.

### Aucun état des prix n'est trouvé

Vérifier que le fichier commence par :

```text
prix
```

et possède une feuille contenant `prix de vente`.

### Une facture est vide

Vérifier que le programme et le LOT affichés dans `Controle Factures` existent réellement dans les colonnes A et B du tracking.

### Le résultat contient un montant inattendu

Vérifier dans l'état des prix :

- le code IS*M ;
- la valeur de la colonne I ;
- la valeur de remplacement de la colonne H ;
- la nature du produit et sa règle de facturation.

## 10. Modules VBA

### `ModuleControleFactures.bas`

Responsable de :

- créer la feuille `Controle Factures` ;
- détecter les groupes programme/LOT ;
- afficher les boutons ;
- lancer une facture individuelle ;
- lancer toutes les factures.

### `ModuleGenerationFactures.bas`

Responsable de :

- copier le modèle ;
- lire le tracking ;
- regrouper les MSN et les VU ;
- calculer les prix ;
- remplir `Infos à remplir` ;
- mettre à jour `Facture finale` ;
- sauvegarder les fichiers générés.

### `ModulePrixFacture.bas`

Module conservé pour la génération/recalcul spécifique de la facture LOT2 lorsque cette fonction est utilisée séparément.

## 11. Bonnes pratiques

- Toujours travailler sur une copie du modèle.
- Ne pas modifier manuellement les fichiers générés comme s'ils étaient le modèle.
- Archiver les anciens exports.
- Vérifier les factures avant envoi.
- Enregistrer les modifications VBA dans le `.xlsm`.
- Ne pas renommer les feuilles principales sans adapter les constantes du code.
- Ne pas supprimer les colonnes source du tracking ou de l'état des prix.

