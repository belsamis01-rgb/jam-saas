# JAM SaaS - Contexte pour audit externe

## Objectif de l'application

JAM SaaS est une application web locale de pilotage pour gérer des projets, tâches, athlètes, clubs, documents, convocations et notifications dans un contexte sportif/événementiel.

L'application est actuellement un prototype React/Vite côté navigateur. Les données sont stockées localement dans le navigateur via `localStorage`.

## Fonctionnalités principales

- Tableau de bord avec indicateurs de projets, tâches, athlètes, clubs et budget.
- Gestion des projets avec projet principal / sous-projet.
- Gestion des tâches liées aux projets.
- Gestion des athlètes et clubs.
- Gestion des documents, convocations et notifications.
- Vues tableau, liste, arborescence et Kanban selon les modules.
- Ajout de colonnes personnalisées sur les tableaux.
- Types de colonnes : texte, nombre, pourcentage, date, financier, statut, case à cocher, personne.
- Import Excel avec gabarit et prévisualisation avant validation.
- Export Excel des données.
- Animation de validation dans la vue Kanban quand un projet passe en terminé.

## Points importants déjà travaillés

- Correction d'un bug de focus dans les formulaires.
- Clarification de la logique projet parent / sous-projet.
- Ajout d'une vue Kanban.
- Remplacement de la dépendance `xlsx` par `exceljs` pour éviter une dépendance vulnérable.
- Ajout d'un gabarit Excel plus exploitable.
- Ajout d'une prévisualisation d'import Excel.
- Ajout de compteurs et repères dans la navigation latérale.
- Vérification récente : build OK et `npm audit --omit=dev` sans vulnérabilité connue.

## Angles d'analyse souhaités

Merci d'analyser le projet sur ces axes :

1. Sécurité applicative :
   - risques liés au stockage local ;
   - risques XSS ou injection ;
   - validation des données importées ;
   - séparation future frontend/backend ;
   - authentification email/mot de passe à prévoir.

2. Architecture :
   - maintenabilité du fichier principal `src/App.jsx` ;
   - découpage recommandé en composants/services/hooks ;
   - structure de données ;
   - évolutivité vers une base de données.

3. Ergonomie :
   - clarté des parcours projet/tâche/athlète ;
   - lisibilité du tableau de bord ;
   - cohérence des vues tableau/Kanban ;
   - priorités d'amélioration UI.

4. Robustesse :
   - gestion des imports Excel volumineux, notamment 1080 athlètes ;
   - erreurs utilisateur ;
   - sauvegarde/restauration ;
   - tests à ajouter.

5. Plan d'amélioration conseillé :
   - quick wins ;
   - risques critiques ;
   - étapes vers une version production sécurisée.

## Commandes utiles

```bash
npm install
npm run dev -- --host 127.0.0.1
npm run build
npm audit --omit=dev
```

URL locale habituelle :

```text
http://localhost:5173/
```
