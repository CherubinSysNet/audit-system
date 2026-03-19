# Audit Système PowerShell

## Description

Script d'audit système complet pour Windows qui collecte des informations détaillées sur le matériel, le réseau et les logiciels, puis génère un rapport HTML professionnel avec une interface à onglets et des thèmes personnalisables.

**Auteur:** Tshienda Cherubin  
**Version:** 4.0  
**Date:** Mars 2026  

## Fonctionnalités

- **Collecte exhaustive d'informations :**
  - Configuration matérielle (processeur, RAM, disques, BIOS)
  - Informations réseau (IP, masque, passerelle, DNS)
  - Mises à jour installées (10 dernières)
  - Temps de fonctionnement du système
  - Version détaillée du système d'exploitation

- **Rapport HTML professionnel :**
  - Interface avec 4 onglets navigables
  - Design responsive (mobile/desktop)
  - Badges de statut colorés (succès/avertissement/critique)
  - Cartes d'information stylisées
  - Optimisé pour l'impression

- **Personnalisation avancée :**
  - 3 thèmes prédéfinis (Light, Dark, Hacker)
  - Configuration manuelle des couleurs
  - Mode interactif ou paramétrage en ligne de commande

## Prérequis

- Windows 7/8/10/11 ou Windows Server 2008 R2+
- PowerShell 5.1 ou supérieur
- Droits administrateur (recommandé pour un accès complet)

## Installation

1. **Téléchargement :**
   # Depuis GitHub
   Invoke-WebRequest -Uri "https://github.com/CherubinSysNet/audit-system/main/Audit-System.ps1" -OutFile "Audit-System.ps1"
2. **Permissions d'exécution :**
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

## Sortie



<img width="1351" height="676" alt="audit2" src="https://github.com/user-attachments/assets/95d6860d-61c3-42ff-b85c-c29140b036fd" /><br /><br />
<img width="1346" height="676" alt="audit1" src="https://github.com/user-attachments/assets/7c5905c7-a523-469e-8e88-d7c8f8fbfba3" /><br /><br />
<img width="1348" height="676" alt="audit" src="https://github.com/user-attachments/assets/dfeb6632-27dc-43c8-9b1c-6f98a305a671" />

