# Suivi des modifications - JAM SaaS

## Rôle du fichier

Ce fichier sert de repère unique pour suivre :

- les versions importantes du projet
- les modifications locales en cours
- les points validés avec les utilisateurs
- les prochaines actions à faire

Il doit être mis à jour à chaque évolution importante de l'application.

---

## Identité du projet

- Dossier local : `/Users/samymettidji/Desktop/jam-saas-v12-test-package`
- Dépôt GitHub : `https://github.com/belsamis01-rgb/jam-saas`
- Branche principale : `main`
- Site Netlify : `https://jam-saas-test.netlify.app`
- Projet Supabase : `jam-saas-test`

---

## Jalons Git connus

| Date | Commit | Intitulé |
|---|---|---|
| 2026-05-17 | `2b2ed71` | `first commit` |
| 2026-07-03 | `2bf6d42` | `Update JAM SaaS dashboard and project setup` |
| 2026-07-21 | `158e0bc` | `Update project UX and attachment interactions` |
| 2026-07-21 | `1161dcf` | `Create netlify.toml` |
| 2026-07-21 | `afcad3c` | `Update App.jsx` |

Repère utile :

- le rendez-vous utilisateurs de juillet repose très probablement sur la lignée de version de juillet 2026
- la base Git distante actuelle pointe sur `afcad3c`

---

## État actuel constaté le 2026-08-24

### Version Git

- `HEAD` local = `origin/main`
- dernier commit connu : `afcad3c`

### Modifications locales non encore cadrées proprement

- fichier modifié : `.claude/settings.local.json`
- fichier modifié : `src/App.jsx`
- fichier non suivi : `.claude/settings.json`
- fichier non suivi : `RESTAURER.txt`
- fichier non suivi : `SAUVEGARDES.txt`

### Lecture simple

Cela veut dire :

- la base officielle distante est bien sauvegardée sur GitHub
- mais le poste local contient encore des changements en plus
- donc l'état actuel visible en local n'est pas forcément exactement identique au dernier état poussé

---

## Ce que l'on sait sur la version de travail

### Partie déjà présente dans la continuité du projet

- gestion des projets
- gestion des tâches
- axes analytiques
- athlètes
- sections
- tableau de bord
- responsive mobile v1
- Netlify connecté
- Supabase réactivé et sain côté infrastructure

### Point important

À ce stade, Supabase est disponible côté plateforme, mais l'application n'est pas encore entièrement branchée dessus dans le code si les données sont encore pilotées localement dans `src/App.jsx`.

---

## Journal de suivi

### 2026-08-24

#### Vérification de cohérence

- confirmation que le projet actuel est bien la continuité de la version travaillée en juillet
- confirmation que GitHub et la branche `main` pointent sur `afcad3c`
- constat de modifications locales supplémentaires dans `src/App.jsx`

#### Besoin exprimé

- créer un vrai suivi des modifications pour éviter de se perdre

#### Décision

- création de ce fichier `SUIVI_MODIFICATIONS.md` comme registre principal du projet
- préparation d'un push GitHub propre en excluant les réglages purement locaux de `.claude`
- nettoyage du dépôt après constat qu'un fichier de configuration locale `.claude/settings.local.json` a été poussé par erreur

#### Modifications poussées / à pousser

- mise à jour de `src/App.jsx`
- ajout de `SUIVI_MODIFICATIONS.md`
- ajout de `RESTAURER.txt`
- ajout de `SAUVEGARDES.txt`
- mise à jour de `.gitignore` pour ignorer `.claude/settings.json`

#### Ménage dépôt

- retrait prévu de `.claude/settings.local.json` du suivi Git
- conservation du fichier uniquement en local sur la machine
- conservation de `.claude/launch.json` qui peut rester utile au projet

---

## Règle de fonctionnement pour la suite

À partir de maintenant, quand une modification importante est faite, ajouter ici :

1. la date
2. ce qui a changé
3. si c'est seulement en local ou poussé sur GitHub
4. si c'est aussi mis en ligne sur Netlify
5. si c'est validé par les utilisateurs ou encore en test

Format recommandé :

```md
### AAAA-MM-JJ
- Modification :
- Fichiers principaux :
- Statut Git : local / commit / push
- Statut Netlify : non déployé / déployé
- Validation : en test / validé
- Notes :
```

---

## Prochaines étapes suggérées

- clarifier les changements locaux actuels dans `src/App.jsx`
- décider ce qu'on garde officiellement
- pousser une version propre sur GitHub
- déployer une version cohérente sur Netlify
- préparer ensuite la vraie couche utilisateurs / administrateurs / confidentialité
