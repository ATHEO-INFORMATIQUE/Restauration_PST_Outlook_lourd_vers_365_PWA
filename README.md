## 📦 Script PowerShell – Migration Automatisés de Fichiers PST Outlook client lourd vers Outlook 365 PWA.

### 🧭 Objectif

Ce script PowerShell a pour but d'automatiser la **gestion des fichiers PST**  pour les importer dans Outlook PWA. Il effectue les opérations suivantes :

1. **Détache tous les fichiers PST** actuellement montés dans le profil Outlook de l'utilisateur.
2. **Monte automatiquement les fichiers PST** présents dans un dossier spécifique de l'utilisateur.
3. **Copie le contenu des fichiers PST** (mails uniquement) dans un dossier nommé `OldArchives` dans la boîte aux lettres principale.
4. **Détache les fichiers PST** une fois la migration terminée.
5. **Génère un fichier de log** détaillant toutes les opérations effectuées.

---

### 🛠️ Prérequis

- Microsoft Outlook client lourd installé sur le poste (version classique et non New Outlook).
- Droits d'exécution de scripts PowerShell (ex. : `ExecutionPolicy Bypass`).
- Les fichiers PST doivent être stockés dans un dossier nommé `pst` situé dans `D:\Utilisateurs\<NomUtilisateur>\pst`. (Possible de modifier)
- Le script doit être exécuté **dans le contexte de l'utilisateur**, sans élévation de privilèges.

---

### 📁 Structure dans ce script

```
D:\
└── Utilisateurs\
|   └── <NomUtilisateur>\
|      └── pst\
|        ├── archive1.pst
|        ├── archive2.pst
|        └── ...
└── logs\
    └── Log_migration_archives_YYYY-MM-DD.log
```

---

### 📝 Fichier de log

Un fichier de log est généré automatiquement dans `D:\logs\` avec un nom basé sur la date du jour :

```
Log_migration_archives_YYYY-MM-DD.log
```

Il contient toutes les étapes du traitement, les erreurs éventuelles, et les fichiers PST traités.

---

### 🔧 Personnalisation

Voici comment adapter le script à votre infrastructure :

| Élément à modifier | Description | Exemple |
|--------------------|-------------|---------|
| `$pstBasePath`     | Chemin de base des fichiers PST | `D:\Utilisateurs\$user` |
| `$logDir`          | Dossier de stockage des logs | `D:\logs` |
| `\\serveur_distant\BackupPC\Migration_archives_script`          | Dossier de stockage des logs distant| `\\Nas01\Migration_archives_script` |
| `"OldArchives"`    | Nom du dossier de destination dans Outlook | Peut être changé selon vos besoins |
| `$minSize`         | Taille minimale d’un fichier PST à traiter | `265KB` par défaut |

---
---

## 📝 Note de mise à jour – Version 1.1 (2025-07-15)

### ✅ Ajouts

- **Vérification automatique du lancement d’Outlook** : le script vérifie si Outlook est lancé, le démarre si nécessaire, et le place au premier plan.
- **Copie automatique du fichier de log** vers le serveur `\\serveur_distant\BackupPC\Migration_archives_script`.
  - Le nom du fichier inclut désormais le nom de l'utilisateur.
  - En cas de doublon, un suffixe numérique est ajouté automatiquement (`_1`, `_2`, etc.).
- **Amélioration des logs** :
  - Ajout d’un message confirmant la copie du fichier de log sur le serveur.
  - Meilleure gestion des erreurs et des éléments non traités.

---

## 📄 Exemple de sortie de log

```
[2025-07-15 08:58:16] Script démarré pour l'utilisateur : tcoutard
[2025-07-15 08:58:16] Chemin PST détecté : D:\Utilisateurs\tcoutard\PST
[2025-07-15 08:58:17] Dossier 'OldArchives' créé.
[2025-07-15 08:58:17] Montage de : D:\Utilisateurs\tcoutard\PST\archive.pst
[2025-07-15 09:02:01] Ignoré (non-mail) : Calendrier
[2025-07-15 09:02:01] Ignoré (non-mail) : Tâches
[2025-07-15 09:02:01] Ignoré (non-mail) : Journal
[2025-07-15 09:02:38] Ignoré (non-mail) : Notes
[2025-07-15 09:02:39] PST 'archive.pst' traité et détaché.
[2025-07-15 09:02:39] Script terminé.
```

### ▶️ Exécution

Lancer le fichier Restauration_PST_365.bat dans le context utilisateur ou

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\migration-pst.ps1
```

---

### 🧑‍💻 Auteur

- Conversion PowerShell : Théo COUTARD

---
