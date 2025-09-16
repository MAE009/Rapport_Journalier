# Rapport_Journalier
C'est une amélioration du système de rapport journalier de LCC

Configuration de la feuille "RJ"

Assurez-vous que votre feuille "RJ" contient un tableau structuré nommé "suivi_intervention" avec les colonnes suivantes :

1. ID (Numéro unique)
2. Nom du bien
3. Localisation
4. État actuel
5. Défaillance constatée
6. Cause de la défaillance
7. Action menée
8. Date de la défaillance
9. Date actuelle
10. Temps
11. Résultat
12. Intervenant
13. Observation

Fonctionnalités implémentées

1. Affichage en temps réel : La ListBox montre tous les rapports existants
2. Filtrage avancé : Par année, mois et jour grâce aux trois ComboBox
3. Ajout sécurisé : Les nouveaux rapports sont ajoutés sans écraser les existants
4. Modification et suppression : Possibilité de modifier ou supprimer un rapport sélectionné
5. Validation des données : Vérification des champs obligatoires avant enregistrement
6. Envoi contrôlé : Le rapport ne peut être envoyé qu'après 17h30 avec demande de confirmation

Personnalisation possible

· Adaptez les valeurs des ComboBox (état actuel, résultat, observation) selon vos besoins
· Modifiez la procédure EnvoyerRapport() selon votre méthode d'envoi préférée (email, impression, etc.)
· Ajustez les largeurs de colonnes dans la ListBox pour une meilleure visibilité
