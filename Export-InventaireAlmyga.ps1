<#
.SYNOPSIS
    Inventaire d'un tenant Microsoft 365 / Entra ID : annuaire, départements, groupes et applications SSO.

.DESCRIPTION
    Script en LECTURE SEULE, à lancer par un administrateur du client.
    - Rôle suffisant pour lire : "Lecteur général" (Global Reader).
    - La première fois, le consentement aux permissions Graph peut demander un "Administrateur général".

    Fichiers produits (CSV séparés par ';', UTF-8, s'ouvrent directement dans Excel) :
      annuaire.csv                        Personnes seulement (comptes actifs, internes, licenciés) :
                                          nom, e-mail, poste, département, manager
      departements.csv                    Départements (champ 'department'), effectif, postes fréquents
      groupes.csv                         Groupes M365 / équipes Teams / sécurité / listes de distribution
      groupes_membres.csv                 Appartenance des personnes aux groupes (membres imbriqués inclus)
      applications.csv                    Applications détectées + indicateurs d'usage
      applications_affectations.csv       Qui (personne ou groupe) est affecté à quelle application, et quelle
                                          personne l'a connectée elle-même (« Se connecter avec Microsoft »)
      matrice_departement_application.csv Usage des applications par département (agrégé)
      matrice_groupe_application.csv      Usage des applications par groupe (agrégé)
      resume.txt                          Synthèse et indicateurs de qualité des données

    Protection des données (nLPD) : les journaux de connexion sont lus en mémoire et seulement
    AGRÉGÉS (par application, département, groupe). Aucun fichier ne contient l'historique de
    connexion d'une personne.

.PARAMETER OutputDir
    Dossier de sortie. Par défaut : .\inventaire_<date>

.PARAMETER SignInDays
    Nombre de jours de journaux de connexion à analyser (1 à 30). Par défaut : 30.

.PARAMETER IncludeMicrosoftApps
    Inclut les applications Microsoft (Teams, SharePoint…). Exclues par défaut, car présentes partout.

.PARAMETER MinGroupSize
    Taille minimale d'un groupe pour qu'il soit marqué "pertinent". Par défaut : 2.

.PARAMETER SkipGroups
    Ne lit pas les groupes (étape 2).

.PARAMETER SkipApplications
    Ne lit pas les applications (étape 3), ni les journaux de connexion.

.PARAMETER IncludeSignIns
    Lit aussi les journaux de connexion (étape 4, désactivée par défaut). Seuls des totaux par
    application, département et groupe sont exportés. Nécessite une licence Entra ID P1 ou P2.

.PARAMETER Yes
    Lance sans demander de confirmation (utile pour une exécution automatisée).

    Sécurité :
    - Seules des permissions en LECTURE sont demandées, et uniquement celles des étapes choisies.
    - Le jeton de connexion n'est gardé qu'en mémoire (-ContextScope Process) : rien n'est écrit sur
      le disque, et la session est fermée à la fin, même en cas d'erreur.
    - Le script ne contacte que Microsoft Graph. Les fichiers restent sur ce poste.

.EXAMPLE
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser   # une seule fois
    .\Export-InventaireAlmyga.ps1

.EXAMPLE
    .\Export-InventaireAlmyga.ps1 -IncludeSignIns   # avec les journaux de connexion (Entra ID P1)
#>
[CmdletBinding()]
param(
    [string]$OutputDir = (Join-Path (Get-Location) ("inventaire_" + (Get-Date -Format "yyyyMMdd_HHmm"))),
    [ValidateRange(1, 30)][int]$SignInDays = 30,
    [switch]$IncludeMicrosoftApps,
    [int]$MinGroupSize = 2,
    [switch]$SkipGroups,
    [switch]$SkipApplications,
    [switch]$IncludeSignIns,
    [switch]$Yes
)

$ErrorActionPreference = 'Stop'
$Graph   = 'https://graph.microsoft.com/v1.0'
$NoDept  = '(non renseigné)'
# Tenants propriétaires des applications Microsoft "first-party"
$MicrosoftTenants = @('f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47')
$Encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
$Warnings = [System.Collections.Generic.List[string]]::new()

# ---------------------------------------------------------------------------
# Fonctions utilitaires
# ---------------------------------------------------------------------------
function Invoke-GraphPaged {
    # Parcourt toutes les pages d'une requête Graph.
    # Sans -OnPage : renvoie les éléments. Avec -OnPage : traite chaque page sans tout garder en mémoire.
    param([Parameter(Mandatory)][string]$Uri, [scriptblock]$OnPage)
    $next = $Uri
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType HashTable
        if ($OnPage) { & $OnPage $resp['value'] } else { $resp['value'] }
        $next = $resp['@odata.nextLink']
    }
}

function Export-Data {
    param($Rows, [string]$Name)
    $path = Join-Path $OutputDir $Name
    $list = @($Rows | Where-Object { $null -ne $_ })
    if ($list.Count -gt 0) {
        $list | Export-Csv -Path $path -NoTypeInformation -Delimiter ';' -Encoding $Encoding
    } else {
        Set-Content -Path $path -Value '' -Encoding $Encoding
    }
    Write-Host ("   -> {0} ({1} lignes)" -f $Name, $list.Count)
}

function Get-DeptBreakdown {
    # "Ventes 60% · Marketing 20% · ..." pour un ensemble d'identifiants utilisateurs
    param($UserIds, [int]$Top = 3)
    $ids = @($UserIds | Where-Object { $_ -and $usersById.ContainsKey($_) })
    if ($ids.Count -eq 0) { return '' }
    ($ids | ForEach-Object { $usersById[$_].Departement } |
        Group-Object | Sort-Object Count -Descending | Select-Object -First $Top |
        ForEach-Object { '{0} {1}%' -f $_.Name, [math]::Round(100 * $_.Count / $ids.Count) }) -join ' · '
}

function New-Set { , [System.Collections.Generic.HashSet[string]]::new() }

function Get-Set {
    # Renvoie l'ensemble stocké sous $Key, ou un ensemble vide (jamais $null, jamais déplié)
    param([hashtable]$Table, [string]$Key)
    if ($Table.ContainsKey($Key)) { , $Table[$Key] } else { , [System.Collections.Generic.HashSet[string]]::new() }
}

function Read-YesNo {
    param([string]$Question, [bool]$Default = $true)
    $hint = if ($Default) { '[O/n]' } else { '[o/N]' }
    while ($true) {
        $r = (Read-Host "$Question $hint").Trim().ToLower()
        if ($r -eq '') { return $Default }
        if ($r -in 'o', 'oui', 'y', 'yes') { return $true }
        if ($r -in 'n', 'non', 'no') { return $false }
    }
}

# ---------------------------------------------------------------------------
# Plan d'exécution et confirmation
# ---------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw "Module manquant. Lancez d'abord : Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
}

$Steps = [ordered]@{
    Annuaire = @{
        Titre   = 'Annuaire et départements'
        Lit     = 'Comptes utilisateurs : nom, e-mail, poste, département, manager ; statut, type et licence du compte, seulement pour ne garder que les personnes (aucun mot de passe)'
        Droits  = @('User.Read.All')
        Fichiers = 'annuaire.csv, departements.csv'
        Actif   = $true; Optionnel = $false
    }
    Groupes = @{
        Titre   = 'Groupes'
        Lit     = 'Groupes Microsoft 365, équipes Teams, groupes de sécurité, listes de distribution, et leurs membres'
        Droits  = @('GroupMember.Read.All')
        Fichiers = 'groupes.csv, groupes_membres.csv'
        Actif   = -not $SkipGroups; Optionnel = $true
    }
    Applications = @{
        Titre   = 'Applications connectées via Microsoft (SSO)'
        Lit     = "Applications d'entreprise, personnes et groupes qui y sont affectés, consentements 'Se connecter avec Microsoft'"
        Droits  = @('Application.Read.All', 'Directory.Read.All')
        Fichiers = 'applications.csv, applications_affectations.csv, matrices *_application.csv'
        Actif   = -not $SkipApplications; Optionnel = $true
    }
    Connexions = @{
        Titre   = "Journaux de connexion ($SignInDays derniers jours), facultatif"
        Lit     = 'Connexions réussies aux applications. Seuls des TOTAUX par application, département et groupe sont exportés, jamais le détail par personne'
        Droits  = @('AuditLog.Read.All')
        Fichiers = '(enrichit les fichiers applications et matrices)'
        Actif   = $IncludeSignIns -and -not $SkipApplications; Optionnel = $true
    }
}

function Show-Plan {
    Write-Host ''
    Write-Host '=====================================================================' -ForegroundColor Cyan
    Write-Host ' Inventaire Microsoft 365 - Almyga      (LECTURE SEULE)' -ForegroundColor Cyan
    Write-Host '=====================================================================' -ForegroundColor Cyan
    Write-Host ' Ce script ne modifie rien dans votre tenant. Il ne contacte que'
    Write-Host ' Microsoft Graph et enregistre les résultats uniquement sur ce poste :'
    Write-Host "   $OutputDir"
    Write-Host ' Le jeton de connexion reste en mémoire et est supprimé à la fin.'
    Write-Host ''
    $n = 0
    foreach ($key in $Steps.Keys) {
        $s = $Steps[$key]; $n++
        $etat  = if ($s.Actif) { 'OUI ' } else { 'NON ' }
        $color = if ($s.Actif) { 'Green' } else { 'DarkGray' }
        Write-Host (" [{0}] Étape {1} - {2}" -f $etat, $n, $s.Titre) -ForegroundColor $color
        Write-Host ("        Lit       : {0}" -f $s.Lit)
        Write-Host ("        Droits    : {0}" -f ($s.Droits -join ', '))
        Write-Host ("        Produit   : {0}" -f $s.Fichiers)
    }
    Write-Host ''
}

Show-Plan
if (-not $Yes) {
    $choice = ''
    while ($choice -notin 'o', 'p', 'n') {
        $choice = (Read-Host "Lancer ces étapes ? [O] Oui   [P] Personnaliser   [N] Annuler").Trim().ToLower()
        if ($choice -eq '') { $choice = 'o' }
    }
    if ($choice -eq 'n') { Write-Host 'Annulé. Aucune donnée lue.' -ForegroundColor Yellow; return }
    if ($choice -eq 'p') {
        foreach ($key in $Steps.Keys) {
            $s = $Steps[$key]
            if (-not $s.Optionnel) { continue }
            if ($key -eq 'Connexions' -and -not $Steps.Applications.Actif) { $s.Actif = $false; continue }
            $s.Actif = Read-YesNo ("Inclure l'étape « {0} » ?" -f $s.Titre) $s.Actif
        }
        Show-Plan
        if (-not (Read-YesNo 'Confirmer et lancer ?' $true)) { Write-Host 'Annulé. Aucune donnée lue.' -ForegroundColor Yellow; return }
    }
}

$DoGroups  = $Steps.Groupes.Actif
$DoApps    = $Steps.Applications.Actif
$DoSignIns = $Steps.Connexions.Actif -and $DoApps
$Scopes = @($Steps.Values | Where-Object { $_.Actif } | ForEach-Object { $_.Droits }) | Select-Object -Unique
# Les affectations d'applications à des groupes ne peuvent être dépliées qu'avec la lecture des membres
if ($DoApps -and -not $DoGroups) { $Scopes = @($Scopes) + 'GroupMember.Read.All' | Select-Object -Unique }

# ---------------------------------------------------------------------------
# Connexion (jeton conservé en mémoire uniquement)
# ---------------------------------------------------------------------------
Import-Module Microsoft.Graph.Authentication
Write-Host "Connexion à Microsoft Graph. Droits demandés (lecture seule) : $($Scopes -join ', ')" -ForegroundColor Cyan
Connect-MgGraph -NoWelcome -ContextScope Process -Scopes $Scopes
try {
$ctx = Get-MgContext
$TenantId = $ctx.TenantId
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
Write-Host "Tenant : $TenantId  |  Compte : $($ctx.Account)"

# Valeurs par défaut pour les étapes non lancées
$groupes        = @()
$groupMembers   = @{}
$groupNames     = @{}
$membershipRows = [System.Collections.Generic.List[object]]::new()
$applications   = @()
$assignRows     = [System.Collections.Generic.List[object]]::new()
$accessByApp    = @{}
$consentByApp   = @{}
$adminConsent   = New-Set
$activeByApp    = @{}
$signinStats    = @{}
$counter        = @{ n = 0 }
$signInsAvailable = $false

# ---------------------------------------------------------------------------
# 1. Annuaire
# ---------------------------------------------------------------------------
Write-Host "`n[1/4] Annuaire" -ForegroundColor Cyan
# Seuls les champs utiles sont lus ; statut, type et licence servent à reconnaître les personnes et ne sont pas exportés
$userSelect = 'id,displayName,givenName,surname,userPrincipalName,mail,jobTitle,department,' +
              'accountEnabled,userType,assignedLicenses'
$rawUsers = @(Invoke-GraphPaged "$Graph/users?`$select=$userSelect&`$expand=manager(`$select=id,displayName,userPrincipalName)&`$top=999")

$usersById = @{}
$annuaire = foreach ($u in $rawUsers) {
    $dept     = if ([string]::IsNullOrWhiteSpace($u.department)) { $NoDept } else { $u.department.Trim() }
    $licensed = [bool]($u.assignedLicenses -and @($u.assignedLicenses).Count -gt 0)
    # "Personne" = compte actif, interne (pas invité) et licencié : exclut salles, boîtes partagées, comptes techniques
    # (la condition de licence est levée plus bas si le tenant n'a aucune licence attribuée)
    $isPerson = [bool]($u.accountEnabled -and $u.userType -ne 'Guest' -and $licensed)
    $row = [pscustomobject]@{
        Id             = $u.id
        Nom            = $u.displayName
        Prenom         = $u.givenName
        NomFamille     = $u.surname
        UPN            = $u.userPrincipalName
        Email          = $u.mail
        Poste          = $u.jobTitle
        Manager        = if ($u.manager) { $u.manager.displayName } else { '' }
        ManagerUPN     = if ($u.manager) { $u.manager.userPrincipalName } else { '' }
        Departement    = $dept
        CompteActif    = $u.accountEnabled
        TypeCompte     = $u.userType
        ComptePersonne = $isPerson
    }
    $usersById[$u.id] = $row
    $row
}
if (-not ($annuaire | Where-Object ComptePersonne)) {
    $msg = "Aucun compte licencié trouvé : tous les comptes actifs et internes sont considérés comme des personnes."
    $Warnings.Add($msg); Write-Warning $msg
    foreach ($row in $annuaire) { $row.ComptePersonne = [bool]($row.CompteActif -and $row.TypeCompte -ne 'Guest') }
}
$people   = @($annuaire | Where-Object ComptePersonne)
$personIds = New-Set
$people | ForEach-Object { [void]$personIds.Add($_.Id) }
# Seules les personnes sont exportées : ni invités, ni salles, ni comptes techniques
Export-Data ($people | Select-Object Id, Nom, Prenom, NomFamille, UPN, Email, Poste, Manager, ManagerUPN, Departement) 'annuaire.csv'

$departements = $people | Group-Object Departement | Sort-Object Count -Descending | ForEach-Object {
    [pscustomobject]@{
        Departement     = $_.Name
        Effectif        = $_.Count
        PostesFrequents = ($_.Group | Where-Object Poste | Group-Object Poste |
                           Sort-Object Count -Descending | Select-Object -First 5 |
                           ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', '
    }
}
Export-Data $departements 'departements.csv'

# ---------------------------------------------------------------------------
# 2. Groupes
# ---------------------------------------------------------------------------
Write-Host "`n[2/4] Groupes" -ForegroundColor Cyan
if (-not $DoGroups) { Write-Host '   étape ignorée' -ForegroundColor DarkGray }
else {
$groupSelect = 'id,displayName,description,groupTypes,securityEnabled,mailEnabled,mail,' +
               'membershipRule,resourceProvisioningOptions,createdDateTime'
$rawGroups = @(Invoke-GraphPaged "$Graph/groups?`$select=$groupSelect&`$top=999")

$i = 0
$groupes = foreach ($g in $rawGroups) {
    $i++
    Write-Progress -Activity 'Membres des groupes' -Status $g.displayName -PercentComplete (100 * $i / [math]::Max(1, $rawGroups.Count))

    $members = New-Set
    Invoke-GraphPaged "$Graph/groups/$($g.id)/transitiveMembers?`$select=id&`$top=999" |
        ForEach-Object { if ($personIds.Contains($_.id)) { [void]$members.Add($_.id) } }
    $groupMembers[$g.id] = $members
    $groupNames[$g.id]   = $g.displayName

    foreach ($uid in $members) {
        $membershipRows.Add([pscustomobject]@{
            GroupeId = $g.id; Groupe = $g.displayName; UtilisateurId = $uid
            UPN = $usersById[$uid].UPN; Departement = $usersById[$uid].Departement
        })
    }

    $type = if ($g.groupTypes -contains 'Unified') {
                if ($g.resourceProvisioningOptions -contains 'Team') { 'Équipe Teams' } else { 'Groupe Microsoft 365' }
            } elseif ($g.securityEnabled -and $g.mailEnabled) { 'Sécurité (avec e-mail)' }
            elseif ($g.securityEnabled) { 'Sécurité' }
            else { 'Liste de distribution' }

    $toutLeMonde = $people.Count -gt 0 -and $members.Count -ge 0.9 * $people.Count
    [pscustomobject]@{
        Id                    = $g.id
        Groupe                = $g.displayName
        Type                  = $type
        Description           = $g.description
        Email                 = $g.mail
        Dynamique             = ($g.groupTypes -contains 'DynamicMembership')
        RegleDynamique        = $g.membershipRule
        NbPersonnes           = $members.Count
        TousLesEmployes       = $toutLeMonde
        Pertinent             = ($members.Count -ge $MinGroupSize -and -not $toutLeMonde)
        DepartementsMembres   = Get-DeptBreakdown $members
        CreeLe                = $g.createdDateTime
    }
}
Write-Progress -Activity 'Membres des groupes' -Completed
$groupes = @($groupes | Sort-Object NbPersonnes -Descending)
Export-Data $groupes 'groupes.csv'
Export-Data $membershipRows 'groupes_membres.csv'
}

# ---------------------------------------------------------------------------
# 3. Applications : catalogue, affectations, consentements
# ---------------------------------------------------------------------------
Write-Host "`n[3/4] Applications" -ForegroundColor Cyan
if (-not $DoApps) { Write-Host '   étape ignorée' -ForegroundColor DarkGray }
else {
# Si l'étape Groupes a été sautée, on lit quand même les membres des groupes affectés à une application
$spSelect = 'id,appId,displayName,publisherName,appOwnerOrganizationId,accountEnabled,' +
            'appRoleAssignmentRequired,preferredSingleSignOnMode,homepage,loginUrl,tags,servicePrincipalType'
$rawSps = @(Invoke-GraphPaged "$Graph/servicePrincipals?`$select=$spSelect&`$top=999")

function Get-Origin($sp) {
    if ($MicrosoftTenants -contains $sp.appOwnerOrganizationId) { 'Microsoft' }
    elseif ($sp.appOwnerOrganizationId -eq $TenantId)          { 'Configurée dans le tenant (galerie ou interne)' }
    else                                                       { 'Éditeur tiers (multi-tenant)' }
}

$spById = @{}; $spByAppId = @{}
foreach ($sp in $rawSps) { $spById[$sp.id] = $sp; $spByAppId[$sp.appId] = $sp }

$candidates = @($rawSps | Where-Object {
    $_.servicePrincipalType -eq 'Application' -and ($IncludeMicrosoftApps -or (Get-Origin $_) -ne 'Microsoft')
})

# Affectations (utilisateurs et groupes) -> accès effectifs par application
$accessByApp = @{}   # appId -> HashSet personnes ayant accès (directement ou via un groupe)
$assignRows  = [System.Collections.Generic.List[object]]::new()
$groupAppRows = [System.Collections.Generic.List[object]]::new()
$i = 0
foreach ($sp in $candidates) {
    $i++
    Write-Progress -Activity 'Affectations des applications' -Status $sp.displayName -PercentComplete (100 * $i / [math]::Max(1, $candidates.Count))
    $set = New-Set
    foreach ($a in (Invoke-GraphPaged "$Graph/servicePrincipals/$($sp.id)/appRoleAssignedTo?`$top=999")) {
        # Invités et comptes techniques ignorés : seules les personnes et les groupes sont exportés
        if ($a.principalType -eq 'User' -and -not $personIds.Contains($a.principalId)) { continue }
        $assignRows.Add([pscustomobject]@{
            Application = $sp.displayName; AppId = $sp.appId
            TypeAffectation = $a.principalType; Affecte = $a.principalDisplayName; AffecteId = $a.principalId
        })
        if ($a.principalType -eq 'User' -and $personIds.Contains($a.principalId)) { [void]$set.Add($a.principalId) }
        elseif ($a.principalType -eq 'Group') {
            if (-not $groupMembers.ContainsKey($a.principalId)) {
                $gm = New-Set
                Invoke-GraphPaged "$Graph/groups/$($a.principalId)/transitiveMembers?`$select=id&`$top=999" |
                    ForEach-Object { if ($personIds.Contains($_.id)) { [void]$gm.Add($_.id) } }
                $groupMembers[$a.principalId] = $gm
            }
            foreach ($uid in $groupMembers[$a.principalId]) { [void]$set.Add($uid) }
        }
    }
    $accessByApp[$sp.appId] = $set
}
Write-Progress -Activity 'Affectations des applications' -Completed

# Consentements OAuth ("Se connecter avec Microsoft")
foreach ($gr in (Invoke-GraphPaged "$Graph/oauth2PermissionGrants")) {
    $sp = $spById[$gr.clientId]
    if (-not $sp) { continue }
    if ($gr.consentType -eq 'AllPrincipals') { [void]$adminConsent.Add($sp.appId) }
    elseif ($personIds.Contains($gr.principalId)) {
        if (-not $consentByApp.ContainsKey($sp.appId)) { $consentByApp[$sp.appId] = New-Set }
        # Add() renvoie $false si la personne est déjà comptée (un consentement par API) : une seule ligne par personne
        $isNew = $consentByApp[$sp.appId].Add($gr.principalId)
        if ($isNew -and $sp.servicePrincipalType -eq 'Application' -and ($IncludeMicrosoftApps -or (Get-Origin $sp) -ne 'Microsoft')) {
            $assignRows.Add([pscustomobject]@{
                Application = $sp.displayName; AppId = $sp.appId
                TypeAffectation = 'Consentement'; Affecte = $usersById[$gr.principalId].Nom; AffecteId = $gr.principalId
            })
        }
    }
}

# ---------------------------------------------------------------------------
# 4. Journaux de connexion (agrégés, jamais exportés par personne)
# ---------------------------------------------------------------------------
Write-Host "`n[4/4] Connexions des $SignInDays derniers jours" -ForegroundColor Cyan
if (-not $DoSignIns) { Write-Host '   étape ignorée' -ForegroundColor DarkGray }
else {
$since = (Get-Date).ToUniversalTime().AddDays(-$SignInDays).ToString('yyyy-MM-ddTHH:mm:ssZ')
$signInsAvailable = $true
try {
    Invoke-GraphPaged "$Graph/auditLogs/signIns?`$filter=createdDateTime ge $since&`$top=999" -OnPage {
        param($items)
        foreach ($s in $items) {
            $counter.n++
            if ($s.status -and $s.status.errorCode -ne 0) { continue }           # échecs ignorés
            if (-not $personIds.Contains([string]$s.userId)) { continue }
            $sp = $spByAppId[$s.appId]
            if ($sp -and -not $IncludeMicrosoftApps -and (Get-Origin $sp) -eq 'Microsoft') { continue }
            if (-not $sp -and -not $IncludeMicrosoftApps -and $s.appDisplayName -like 'Microsoft*') { continue }

            if (-not $activeByApp.ContainsKey($s.appId)) {
                $activeByApp[$s.appId] = New-Set
                $signinStats[$s.appId] = @{ Hits = 0; Last = $s.createdDateTime; Name = $s.appDisplayName }
            }
            [void]$activeByApp[$s.appId].Add([string]$s.userId)
            $st = $signinStats[$s.appId]
            $st.Hits++
            if ([datetime]$s.createdDateTime -gt [datetime]$st.Last) { $st.Last = $s.createdDateTime }
        }
        Write-Host ("   {0} connexions lues..." -f $counter.n)
    }
} catch {
    $signInsAvailable = $false
    $msg = "Journaux de connexion indisponibles (licence Entra ID P1/P2 absente ou droits insuffisants) : $($_.Exception.Message)"
    $Warnings.Add($msg); Write-Warning $msg
}
}

# ---------------------------------------------------------------------------
# Synthèse applications
# ---------------------------------------------------------------------------
$appIds = New-Set
$candidates | Where-Object {
    $_.tags -contains 'WindowsAzureActiveDirectoryIntegratedApp' -or
    $accessByApp[$_.appId].Count -gt 0 -or $consentByApp.ContainsKey($_.appId) -or $adminConsent.Contains($_.appId)
} | ForEach-Object { [void]$appIds.Add($_.appId) }
$activeByApp.Keys | ForEach-Object { [void]$appIds.Add($_) }

$applications = foreach ($appId in $appIds) {
    $sp      = $spByAppId[$appId]
    $access  = Get-Set $accessByApp  $appId
    $consent = Get-Set $consentByApp $appId
    $active  = Get-Set $activeByApp  $appId
    $stats   = $signinStats[$appId]

    $users = New-Set
    foreach ($s in @($access, $consent, $active)) { foreach ($u in $s) { [void]$users.Add($u) } }

    [pscustomobject]@{
        Application              = if ($sp) { $sp.displayName } else { $stats.Name }
        Editeur                  = if ($sp) { $sp.publisherName } else { '' }
        Origine                  = if ($sp) { Get-Origin $sp } else { 'Inconnue' }
        ModeSSO                  = if ($sp) { $sp.preferredSingleSignOnMode } else { '' }
        AffectationRequise       = if ($sp) { $sp.appRoleAssignmentRequired } else { '' }
        Active                   = if ($sp) { $sp.accountEnabled } else { '' }
        URL                      = if ($sp) { (@($sp.homepage, $sp.loginUrl) | Where-Object { $_ } | Select-Object -First 1) } else { '' }
        NbPersonnesAvecAcces     = $access.Count
        ConsentementAdmin        = $adminConsent.Contains($appId)
        NbConsentementsPersonnes = $consent.Count
        NbUtilisateursActifs     = if ($signInsAvailable) { $active.Count } else { 'n/d' }
        NbConnexions             = if ($stats) { $stats.Hits } elseif ($signInsAvailable) { 0 } else { 'n/d' }
        DerniereConnexion        = if ($stats) { $stats.Last } else { '' }
        DepartementsPrincipaux   = $(if ($active.Count -gt 0) { Get-DeptBreakdown -UserIds $active } else { Get-DeptBreakdown -UserIds $users })
        AppId                    = $appId
    }
}
$applications = @($applications | Sort-Object { [int]("0" + ($_.NbUtilisateursActifs -replace '\D')) }, NbPersonnesAvecAcces -Descending)
Export-Data $applications 'applications.csv'
Export-Data $assignRows 'applications_affectations.csv'

# Matrice département x application
$deptSize = @{}
$people | Group-Object Departement | ForEach-Object { $deptSize[$_.Name] = $_.Count }
$matriceDept = foreach ($app in $applications) {
    $appId  = $app.AppId
    $access  = Get-Set $accessByApp  $appId
    $active  = Get-Set $activeByApp  $appId
    $consent = Get-Set $consentByApp $appId
    $all = New-Set
    foreach ($s in @($access, $active, $consent)) { foreach ($u in $s) { [void]$all.Add($u) } }

    $all | Group-Object { $usersById[$_].Departement } | ForEach-Object {
        $dName = $_.Name
        $ids = $_.Group
        $nActive = @($ids | Where-Object { $active.Contains($_) }).Count
        [pscustomobject]@{
            Departement          = $dName
            Application          = $app.Application
            EffectifDepartement  = $deptSize[$dName]
            AvecAcces            = @($ids | Where-Object { $access.Contains($_) }).Count
            UtilisateursActifs   = if ($signInsAvailable) { $nActive } else { 'n/d' }
            PartActifsPct        = if ($signInsAvailable -and $deptSize[$dName]) { [math]::Round(100 * $nActive / $deptSize[$dName]) } else { '' }
        }
    }
}
Export-Data ($matriceDept | Sort-Object Departement, Application) 'matrice_departement_application.csv'

# Matrice groupe x application (utile quand le champ département est mal rempli)
if ($DoGroups) {
$pertinentGroups = $groupes | Where-Object { $_.Pertinent -and $_.Type -ne 'Liste de distribution' }
$matriceGroupe = foreach ($g in $pertinentGroups) {
    $members = $groupMembers[$g.Id]
    foreach ($app in $applications) {
        $appId  = $app.AppId
        $source = if ($signInsAvailable) { $activeByApp } else { $accessByApp }
        $use = Get-Set $source $appId
        if ($null -eq $use -or $use.Count -eq 0) { continue }
        $n = @($members | Where-Object { $use.Contains($_) }).Count
        if ($n -eq 0) { continue }
        [pscustomobject]@{
            Groupe       = $g.Groupe
            TypeGroupe   = $g.Type
            NbMembres    = $g.NbPersonnes
            Application  = $app.Application
            Utilisateurs = $n
            PartPct      = [math]::Round(100 * $n / [math]::Max(1, $g.NbPersonnes))
            Base         = if ($signInsAvailable) { 'connexions' } else { 'affectations' }
        }
    }
}
Export-Data ($matriceGroupe | Sort-Object Groupe, PartPct -Descending) 'matrice_groupe_application.csv'
}
}

# ---------------------------------------------------------------------------
# Résumé
# ---------------------------------------------------------------------------
$withDept  = @($people | Where-Object { $_.Departement -ne $NoDept }).Count
$withTitle = @($people | Where-Object Poste).Count
$pct = { param($n) if ($people.Count) { [math]::Round(100 * $n / $people.Count) } else { 0 } }

$resume = @"
INVENTAIRE MICROSOFT 365 - $(Get-Date -Format 'yyyy-MM-dd HH:mm')
Tenant : $TenantId
Exécuté par : $($ctx.Account)

ANNUAIRE
  Comptes au total              : $($annuaire.Count)
  Personnes (actives, internes, licenciées) : $($people.Count)
  Département renseigné         : $withDept ($(& $pct $withDept) %)
  Poste renseigné               : $withTitle ($(& $pct $withTitle) %)
  Départements distincts        : $(@($departements | Where-Object Departement -ne $NoDept).Count)

ÉTAPES LANCÉES
  Annuaire : oui  |  Groupes : $(if ($DoGroups) {'oui'} else {'non'})  |  Applications : $(if ($DoApps) {'oui'} else {'non'})  |  Connexions : $(if ($DoSignIns) {'oui'} else {'non'})
  Droits utilisés (lecture seule) : $($Scopes -join ', ')

GROUPES
  Groupes au total              : $($groupes.Count)
  Équipes Teams                 : $(@($groupes | Where-Object Type -eq 'Équipe Teams').Count)
  Groupes pertinents (>= $MinGroupSize personnes, hors "tout le monde") : $(@($groupes | Where-Object Pertinent).Count)

APPLICATIONS (hors Microsoft : $(-not $IncludeMicrosoftApps))
  Applications détectées        : $($applications.Count)
  Journaux de connexion         : $(if ($signInsAvailable) { "oui, $SignInDays jours, $($counter.n) connexions lues" } elseif ($DoSignIns) { 'NON DISPONIBLES' } else { 'non demandés' })

LIMITES
  - Les outils sans SSO Microsoft (mot de passe propre, logiciels installés, SSO Google) n'apparaissent pas.
  - Un outil peu utilisé peut ne pas apparaître dans $SignInDays jours de connexions.
  - Les connexions non interactives (synchronisations, jetons rafraîchis) ne sont pas comptées.

AVERTISSEMENTS
$(if ($Warnings.Count) { ($Warnings | ForEach-Object { "  - $_" }) -join "`n" } else { '  aucun' })
"@
Set-Content -Path (Join-Path $OutputDir 'resume.txt') -Value $resume -Encoding $Encoding
Write-Host "`n$resume"
Write-Host "Fichiers enregistrés dans : $OutputDir" -ForegroundColor Green

} finally {
    # Ferme la session Graph et efface le jeton, même en cas d'erreur
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Host 'Session Microsoft Graph fermée.' -ForegroundColor DarkGray
}
