<#
.SYNOPSIS
    Génère Edge-SaveAs-Check.PAD-standalone.ps1 : le script complet, prêt à être collé dans
    l'action « Exécuter un script PowerShell » de Power Automate Desktop.

.DESCRIPTION
    Le fichier généré contient les paramètres PAD (variables BrowserPageURL et BrowserPageTitle)
    suivis du script Edge-SaveAs-Check.ps1 enveloppé dans un bloc exécuté avec ces paramètres.
    Il voyage avec le flux PAD : rien à installer ni à partager sur les postes.

    À relancer après toute modification de Edge-SaveAs-Check.ps1 (le fichier généré est versionné
    à côté de la source). Vérifie le résultat : syntaxe valide, aucun signe pourcent hors des
    deux variables PAD, pas de BOM.

.EXAMPLE
    .\Build-PadStandalone.ps1
#>
[CmdletBinding()]
param(
    [string]$Source = (Join-Path $PSScriptRoot 'Edge-SaveAs-Check.ps1'),
    [string]$Output = (Join-Path $PSScriptRoot 'Edge-SaveAs-Check.PAD-standalone.ps1')
)
$ErrorActionPreference = 'Stop'

# En-tête : variables PAD et dossier de journaux propre à chaque utilisateur. Chaîne entre
# guillemets (et non here-string simple) car elle contient elle-même des here-strings @' '@.
$header = @"
# =============================================================================
# Edge-SaveAs-Check : version autonome pour Power Automate Desktop (fichier GENERE par
# Build-PadStandalone.ps1, ne pas modifier a la main).
#
# A coller tel quel dans l'action " Executer un script PowerShell ".
#  - Creer avant cette action les variables PAD BrowserPageURL et BrowserPageTitle (action
#    " Definir la variable "). Jokers acceptes (ex. *monappli/labo*), apostrophes autorisees.
#  - Action PAD : desactiver " Echec apres expiration du delai " ou mettre au moins 120 s.
#  - Resultat : PowershellOutput contient le code (0 = fenetre " Enregistrer sous " ouverte),
#    ScriptError contient le message en cas d'echec. Codes : voir l'en-tete du script ci-dessous.
#  - Journaux : dossier EdgeSaveAsCheck\Logs du profil de chaque utilisateur.
#  - Aucun signe pourcent ne doit apparaitre ici, sauf autour des variables PAD.
# =============================================================================

`$parameters = @{
    UrlPattern      = @'
%BrowserPageURL%
'@
    TitlePattern    = @'
%BrowserPageTitle%
'@
    FileNamePattern = '*'                  # ex. 'Instruments*.xlsx'
    LogPath         = `$(if (`$env:LOCALAPPDATA) { Join-Path `$env:LOCALAPPDATA 'EdgeSaveAsCheck\Logs' } else { '' })
    CloseDialog     = `$false               # `$true : ferme la fenetre apres verification
    # AllowRealMouse = `$true              # dernier recours : vrai clic souris (curseur deplace un instant)
}

"@

$lines = [System.IO.File]::ReadAllText($Source) -split "`r?`n"
# La ligne #Requires n'a de sens qu'en tête d'un fichier : retirée du bloc.
$body = ($lines | Where-Object { $_ -notmatch '^\s*#Requires\b' }) -join "`n"

$content = $header + "& {`n" + $body.TrimEnd() + "`n} @parameters`n"
[System.IO.File]::WriteAllText($Output, $content, (New-Object System.Text.UTF8Encoding($false)))

# --- Contrôles du fichier généré ---------------------------------------------------------
$tokens = $null; $errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($Output, [ref]$tokens, [ref]$errors)
if ($errors.Count -gt 0) { throw ('Le fichier genere contient des erreurs de syntaxe : ' + $errors[0].Message) }
$percentCount = ([regex]::Matches($content, '%')).Count
$pad = [regex]::Matches($content, '%(BrowserPageURL|BrowserPageTitle)%').Count
if ($percentCount -ne 4 -or $pad -ne 2) { throw "Signes pourcent inattendus : $percentCount au total (attendu 4, soit 2 variables PAD), $pad variables PAD." }
$bytes = [System.IO.File]::ReadAllBytes($Output)
if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { throw 'Le fichier genere commence par un BOM.' }
$hash = (Get-FileHash -LiteralPath $Output -Algorithm SHA256).Hash
Write-Output ('Genere : {0} ({1:N0} octets, {2} lignes) - SHA-256 {3}' -f $Output, $bytes.Length, ($content -split "`n").Count, $hash)
