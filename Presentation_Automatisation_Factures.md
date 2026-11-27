# Automatisation de la création des factures

## Message clé

Nous avons transformé une préparation de factures manuelle, longue et dépendante de plusieurs tableaux intermédiaires en un processus automatisé et reproductible :

> **Le responsable sélectionne le programme/LOT, lance la génération, puis récupère directement la facture prête à contrôler.**

---

## 1. Situation avant l'automatisation

### Processus manuel

Pour chaque programme et produit, le facturier devait :

1. Ouvrir le fichier **TRACKING EXPORT**.
2. Identifier les lignes du programme et du LOT concerné.
3. Copier manuellement les colonnes **MSN** et **VU** dans un nouveau tableau.
4. Créer un deuxième tableau pour supprimer les MSN en double.
5. Copier les MSN/VU dans le fichier facture.
6. Reconstituer manuellement les références.
7. Utiliser une grille auxiliaire de nombreuses colonnes pour retrouver les VU.
8. Rechercher les prix dans l'ancien fichier de prix.
9. Vérifier les doublons, les VU manquants et les montants.
10. Répéter le travail pour chaque programme et chaque LOT.

### Risques de l'ancien processus

- Copier-coller manuel et risque d'erreur humaine.
- Doublons ou lignes oubliées.
- Références mal reconstituées.
- Prix dépendants d'une grille complexe et difficile à maintenir.
- Dépendance à une ancienne grille auxiliaire dans les colonnes **I:IO**.
- Certaines formules I:IO appelaient un fichier situé sur le réseau de l'usine,
  créé par une ancienne personne en stage et désormais supprimé.
- Ouverture de la facture avec des liens cassés, des erreurs de recherche ou des
  résultats impossibles à expliquer.
- Forte dépendance à la connaissance du facturier.
- Difficulté à reproduire ou auditer le résultat.
- Temps de préparation élevé lorsque plusieurs factures doivent être créées.

### Pourquoi une seule facture était stressante

Créer une seule facture ne consistait pas simplement à cliquer sur un bouton.
Le facturier devait reconstruire manuellement la facture et vérifier plusieurs
sources en même temps :

- retrouver les bons MSN et VU ;
- supprimer les doublons sans supprimer une vraie ligne ;
- reconstruire la référence exacte ;
- retrouver les prix dans les formules I:IO ;
- comprendre les liens vers l'ancien fichier réseau ;
- vérifier que le fichier externe existait encore ;
- contrôler le total final malgré les erreurs éventuelles.

Une erreur sur une seule ligne pouvait modifier la référence, le prix ou le
montant de la facture. Cette situation créait une pression importante : le
facturier devait être à la fois opérateur de saisie, contrôleur des données et
mainteneur d'un ancien système de formules qu'il n'avait pas créé.

Le problème n'était donc pas uniquement le temps passé. C'était aussi
l'impossibilité de garantir rapidement que le résultat était fiable lorsque la
source historique appelée par I:IO n'existait plus.

---

## 2. Situation après l'automatisation

### Nouveau processus

1. Le fichier de contrôle détecte automatiquement les combinaisons présentes dans le tracking.
2. Une ligne et un bouton sont créés pour chaque programme/LOT.
3. Le bouton récupère automatiquement :
   - le programme ;
   - le LOT ;
   - les MSN ;
   - les VU ;
   - les codes article IS*M ;
   - les prix de l'état des prix.
4. Les MSN/VU sont regroupés automatiquement.
5. Les références sont générées automatiquement.
6. Les prix sont recherchés à partir de l'**état des prix export à jour**.
7. Une facture séparée est créée pour chaque groupe.
8. Un bouton permet également de générer toutes les factures en une seule opération.

### Cas particulier WINGS

Le tracking contient des LOT comme :

- `1/900`
- `2/900`
- `3/900`
- `1/1000`
- `2/1000`
- `42/1000`
- `412/1000`

Le système comprend que `900` et `1000` sont les deux variations principales :

- tous les LOT `X/900` sont regroupés dans la facture **WINGS 900** ;
- tous les LOT `X/1000` sont regroupés dans la facture **WINGS 1000** ;
- les VU détaillés restent conservés dans le contenu de la facture.

Cela évite de créer une facture séparée pour chaque sous-variation.

---

## 3. Résultats mesurés sur le fichier de démonstration

Le fichier TRACKING EXPORT fourni contient :

| Indicateur | Résultat |
|---|---:|
| Lignes tracking analysées | **213** |
| Combinaisons programme/LOT détectées | **19** |
| Programmes concernés | **8** |
| Lignes WINGS | **33** |
| Factures WINGS après regroupement | **2** |
| Lignes WINGS 900 | **13** |
| Lignes WINGS 1000 | **20** |

### Programmes détectés

- A320-90VU
- A320-CCP
- A350-RACK
- A400 M
- ATR
- A350-CCP
- WINGS
- AH

---

## 4. Comparaison avant / après

| Étape | Avant | Après |
|---|---|---|
| Lecture du tracking | Manuelle | Automatique |
| Sélection des lignes | Recherche/copie | Filtrage automatique |
| Tableau MSN/VU | Créé manuellement | Généré automatiquement |
| Déduplication MSN | Deuxième tableau manuel | Regroupement automatique |
| Références | Construites avec formules complexes | Générées directement |
| Recherche des prix | Grille auxiliaire I:AN/longue table | Recherche IS*M dans l'état des prix |
| Cas WINGS | Plusieurs factures par sous-LOT | Une facture par variation 900/1000 |
| Création de plusieurs factures | Répétitive | Bouton individuel ou “Tout générer” |
| Traçabilité | Dépend du fichier de travail | Résultat reproductible |
| Risque d'erreur | Élevé | Réduit et contrôlable |

---

## 5. Gain de temps

Les chiffres exacts dépendent du temps réel mesuré auprès du facturier. Le modèle de calcul est :

```text
Temps économisé par facture
= temps manuel moyen - temps automatisé moyen
```

```text
Temps économisé total
= nombre de factures × temps économisé par facture
```

### Exemple de scénario à présenter

> **Hypothèse à confirmer sur le terrain :**
>
> - préparation manuelle : 20 minutes par facture ;
> - contrôle et lancement automatisés : 20 seconds par 20 Factures ;
> - 19 combinaisons détectées.

| Indicateur | Calcul | Résultat |
|---|---:|---:|
| Temps manuel estimé | 19 × 20 min | 380 min |
| Temps automatisé estimé | 19 × 3 min | 57 min |
| Gain estimé | 380 - 57 min | **323 min** |
| Gain en heures | 323 / 60 | **5 h 23 min** |
| Réduction du temps | 323 / 380 | **85 %** |

Ces valeurs sont un scénario de présentation, pas une mesure comptable définitive. Elles doivent être remplacées par les temps observés dans l'équipe.

---

## 6. Valorisation financière

Le gain financier peut être calculé sans inventer de montant :

```text
Économie financière
= heures économisées × coût horaire chargé
```

### Exemple avec un coût horaire de 35 €/h

| Indicateur | Calcul | Résultat |
|---|---:|---:|
| Heures économisées dans le scénario | 5,38 h | 5,38 h |
| Coût horaire chargé | Hypothèse | 35 €/h |
| Économie estimée | 5,38 × 35 | **188,30 €** |

Si le volume est mensuel :

```text
Économie annuelle
= économie par cycle × nombre de cycles par an
```

Exemple :

```text
188,30 € × 12 cycles = 2 259,60 € par an
```

Le montant final doit être recalculé avec :

- le nombre réel de factures par mois ;
- le temps manuel réellement observé ;
- le temps de contrôle après automatisation ;
- le coût horaire chargé validé par l'entreprise.

---

## 7. Gains qualitatifs

### Fiabilité

- Moins de manipulations manuelles.
- Moins de risque de doublons MSN.
- Moins de risque d'oublier un VU.
- Prix centralisés dans une source de vérité.

### Standardisation

- Même méthode pour chaque programme.
- Même structure de facture.
- Même logique de regroupement.
- Résultats reproductibles par plusieurs personnes.

### Maintenabilité

- Les fichiers de prix peuvent être remplacés sans reconstruire les tableaux.
- Les nouveaux programmes/LOT sont détectés automatiquement.
- Les règles spécifiques, comme WINGS 900/1000, sont explicites.

### Continuité opérationnelle

- Le processus ne dépend plus d'une seule personne.
- La connaissance est intégrée dans les macros.
- Le remplacement ou l'absence du facturier a moins d'impact.

---

## 8. Architecture de la solution

### Fichiers d'entrée

- **TRACKING EXPORT** : source des programmes, LOT, MSN, VU et codes IS*M.
- **État des prix** : source des prix par IS*M.
- **Modèle facture** : structure et présentation de la facture.

### Modules VBA

- Module de contrôle : création de la feuille de commandes et des boutons.
- Module de génération : création des factures séparées.
- Module de calcul historique : conservé pour compatibilité avec la facture LOT2.

### Sortie

Les factures sont générées dans le dossier :

```text
Factures Generated
```

Chaque fichier est nommé selon son programme et sa variation, par exemple :

```text
WINGS_900.xlsx
WINGS_1000.xlsx
```

---

## 9. Contrôle humain conservé

L'automatisation ne supprime pas le contrôle métier. Elle supprime surtout les tâches répétitives.

Le facturier conserve la responsabilité de vérifier :

- les données du tracking ;
- la cohérence des MSN/VU ;
- les montants exceptionnels ;
- la date et les informations commerciales ;
- le PDF ou document final avant envoi.

La différence est que le contrôle porte désormais sur le résultat, et non sur la reconstruction manuelle de toutes les données.

---

## 10. Conclusion pour la direction

### Avant

Un processus manuel prondre 5-10 minute pour 1 LOT/Program , fragmenté et difficile à contrôler, avec plusieurs tableaux intermédiaires et une forte possibilité d'erreur.

### Après

Un processus automatisé, standardisé et extensible :

- 213 lignes traitées automatiquement ;
- 19 groupes détectés ;
- regroupement métier WINGS 900/1000 ;
- génération individuelle ou globale ;
- prix calculés depuis une source de vérité ;
- génération de toutes les factures des programmes et de leurs LOT en environ
  **20 secondes** sur le fichier de démonstration ;
- passage d'une construction manuelle stressante à une validation du résultat ;
- meilleure traçabilité et moindre dépendance à une personne.

### Résumé de la valeur créée

Avant, une seule facture pouvait demander de nombreuses manipulations et
générer du stress à cause des liens I:IO vers un fichier réseau supprimé.
Chaque nouveau programme ou LOT augmentait le risque et le temps de préparation.

Après, le responsable peut :

1. cliquer sur **Invoices Controle** dans le ruban ;
2. créer automatiquement la feuille **Controle Factures** ;
3. cliquer sur **Generate all invoices** ;
4. obtenir toutes les factures séparées en environ **20 secondes** ;
5. effectuer uniquement le contrôle métier final.

> **La valeur principale du projet est de remplacer une reconstruction manuelle,
> lente et stressante par une génération centralisée, rapide et contrôlable.**

> **Le gain n'est pas seulement quelques minutes par facture : c'est la
> suppression d'une dépendance critique à un fichier réseau supprimé et à une
> logique I:IO devenue fragile.**
