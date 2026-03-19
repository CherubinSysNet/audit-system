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

