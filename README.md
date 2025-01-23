# Pricer d'Obligation en VBA

## Description
Ce projet est un pricer d'obligations développé en VBA. Il récupère les données depuis une base Microsoft Access (alimentée par Bloomberg) et permet de calculer les prix des obligations en fonction des taux et des échéances.

## Structure du projet
- **Classes :**
  - `Class_Bond.cls` : Gestion et pricing des obligations.
  - `Class_Curve.cls` : Gestion des courbes de taux pour l'évaluation des obligations.
- **Modules VBA :**
  - `Insertion_Access.bas` : Module d'insertion et de récupération des données depuis Access.
  - `Interface.bas` : Interface utilisateur et interactions pour faciliter l'utilisation du pricer.
  - `mod_toolbox.bas` : Fonctions utilitaires pour le pricing et la gestion des données.
  - `RecuperationData.bas` : Récupération et traitement des données de marché.
  - `Toolbox.bas` : Module contenant des outils supplémentaires.

## Prérequis
- Microsoft Excel avec support VBA activé
- Microsoft Access (Base alimentée par Bloomberg)
- Connexion à la base de données Access
