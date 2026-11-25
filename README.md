# Inventaire Microsoft 365 – Almyga

Script PowerShell en **lecture seule** qui inventorie un tenant Microsoft 365 : personnes, départements, groupes et applications connectées via Microsoft. Almyga s'en sert pour adapter sa formation à la protection des données aux outils réellement utilisés par chaque métier.

## Ce que fait le script

- Il ne modifie rien : il ne demande que des permissions Microsoft Graph en lecture.
- Il ne contacte que Microsoft. Les résultats restent sur votre poste.
- Le jeton de connexion reste en mémoire. La session est fermée à la fin, même en cas d'erreur.
- Il affiche son plan et demande votre confirmation avant toute lecture.
- Il ne lit aucun mot de passe.

| Étape | Ce qui est lu | Permission |
|---|---|---|
| 1. Annuaire (obligatoire) | Nom, e-mail, poste, département et manager des personnes. Le statut, le type et la licence du compte servent seulement à écarter les invités, les salles et les comptes techniques. | User.Read.All |
| 2. Groupes (facultatif) | Équipes Teams, groupes et leurs membres | GroupMember.Read.All |
| 3. Applications (facultatif) | Applications connectées via Microsoft (SSO), et les personnes et groupes qui y ont accès | Application.Read.All, Directory.Read.All |

## Fichiers produits

Dans un dossier `inventaire_<date>`, en CSV (séparateur `;`, UTF-8, s'ouvrent dans Excel) :

| Fichier | Contenu |
|---|---|
| annuaire.csv | Personnes : nom, e-mail, poste, département, manager |
| departements.csv | Départements, effectif, postes fréquents |
| groupes.csv | Groupes Microsoft 365, équipes Teams, groupes de sécurité, listes de distribution |
| groupes_membres.csv | Groupes dont chaque personne est membre |
| applications.csv | Applications détectées et nombre de personnes |
| applications_affectations.csv | Personnes et groupes qui ont accès à chaque application |
| matrice_departement_application.csv | Applications par département (totaux) |
| matrice_groupe_application.csv | Applications par groupe (totaux) |
| resume.txt | Synthèse et qualité des données |

Ces fichiers contiennent des données personnelles. Transmettez-les par un canal sécurisé, jamais en pièce jointe.

## Prérequis

- PowerShell 5.1 ou 7 sous Windows, PowerShell 7 sous macOS.
- Le module Microsoft Graph, à installer une seule fois :
  ```powershell
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  ```
- Le rôle « Lecteur général ». Au premier lancement, un Administrateur général accepte les permissions.

## Utilisation

1. Téléchargez `Export-InventaireAlmyga.ps1`.
2. Vérifiez son empreinte SHA-256 :
   ```powershell
   Get-FileHash .\Export-InventaireAlmyga.ps1 -Algorithm SHA256
   ```
   Résultat attendu : `CB3828E09B9857ED4C99BFE68C84472F1A4F21E00AE83DD98975FE6CD905EC02`
3. Lancez-le :
   - Windows : `powershell -ExecutionPolicy Bypass -File .\Export-InventaireAlmyga.ps1`
   - macOS : `pwsh ./Export-InventaireAlmyga.ps1`
4. Transmettez le dossier `inventaire_<date>` par le lien sécurisé envoyé par Almyga.

Durée : quelques minutes pour une PME.

## Options

| Option | Effet |
|---|---|
| `-SkipGroups` | Ne lit pas les groupes (étape 2) |
| `-SkipApplications` | Ne lit pas les applications (étape 3) |
| `-IncludeSignIns` | Lit aussi les journaux de connexion. Seuls des totaux par application, département et groupe sont exportés, jamais le détail par personne. Nécessite Entra ID P1. |
| `-OutputDir <chemin>` | Choisit le dossier de sortie |

Toutes les options : `Get-Help .\Export-InventaireAlmyga.ps1 -Full`

## Après le lancement

Le consentement donné au premier lancement reste dans le tenant. Si Microsoft Graph PowerShell n'était pas utilisé avant chez vous, vous pouvez le retirer dans le centre d'administration Microsoft Entra, sous Applications d'entreprise > « Microsoft Graph Command Line Tools ».

Questions : contactez votre interlocuteur Almyga.
