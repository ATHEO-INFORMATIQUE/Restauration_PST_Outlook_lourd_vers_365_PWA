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
| `"OldArchives"`    | Nom du dossier de destination dans Outlook | Peut être changé selon vos besoins |
| `$minSize`         | Taille minimale d’un fichier PST à traiter | `265KB` par défaut |

---

### ▶️ Exécution

Lancer le fichier Restauration_PST_365.bat dans le context utilisateur ou

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\migration-pst.ps1
```

---

### 🧑‍💻 Auteur

- Conversion PowerShell : Théo COUTARD

---
