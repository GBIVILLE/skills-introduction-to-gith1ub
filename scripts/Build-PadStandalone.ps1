<#
.SYNOPSIS
    Génère Edge-SaveAs-Check.PAD-standalone.ps1 : le script complet, prêt à être collé dans
    l'action « Exécuter un script PowerShell » de Power Automate Desktop.

.DESCRIPTION
    Le fichier généré contient les paramètres PAD (variables BrowserPageURL et BrowserPageTitle)
    suivis du script Edge-SaveAs-Check.ps1 enveloppé dans un bloc exécuté avec ces paramètres.
    Il voyage avec le flux PAD : rien à installer ni à partager sur les postes.

    Les commentaires du script sont retirés (le fichier passe d'environ 63 Ko à 45 Ko et ne
    contient plus que de l'ASCII, donc aucun risque d'encodage au collage dans PAD). Le retrait
    est vérifié : la suite des éléments du langage (jetons) hors commentaires doit être
    strictement identique avant et après, sinon la génération échoue.

    À relancer après toute modification de Edge-SaveAs-Check.ps1 (le fichier généré est versionné
    à côté de la source). Vérifie le résultat : syntaxe valide, jetons identiques, ASCII pur,
    aucun signe pourcent hors des deux variables PAD, pas de BOM.

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
#  - Resultat : PowershellOutput contient le code, ScriptError le message en cas d'echec.
#      0 fenetre " Enregistrer sous " ouverte      1 erreur inattendue (voir le journal)
#      2 Edge n'est pas lance                      3 aucune page Edge ne correspond
#      4 bouton Telechargements introuvable        5 bouton " Enregistrer sous " introuvable
#      6 la fenetre Windows ne s'est pas ouverte   8 PowerShell en mode restreint
#      9 variable PAD vide
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

# --- Lecture du source et retrait de la ligne #Requires ---------------------------------
$sourceText = [System.IO.File]::ReadAllText($Source) -replace "`r`n", "`n"
$lines = $sourceText -split "`n"
# La ligne #Requires n'a de sens qu'en tête d'un fichier : retirée du bloc.
$bodyWithComments = (($lines | Where-Object { $_ -notmatch '^\s*#Requires\b' }) -join "`n").TrimEnd()

# --- Retrait des commentaires -------------------------------------------------------------
function Get-CodeTokens {
    param([string]$Text)
    $tokens = $null; $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$tokens, [ref]$errors)
    if ($errors.Count -gt 0) { throw ('Erreur de syntaxe : ' + $errors[0].Message + ' (ligne ' + $errors[0].Extent.StartLineNumber + ')') }
    return $tokens
}

function Remove-Comments {
    param([string]$Text)
    # Du dernier commentaire au premier, pour que les positions restent valables. Un commentaire
    # n'est jamais à l'intérieur d'une chaîne ou d'un here-string : l'analyseur le sait.
    $comments = @(Get-CodeTokens $Text | Where-Object { $_.Kind -eq [System.Management.Automation.Language.TokenKind]::Comment })
    for ($i = $comments.Count - 1; $i -ge 0; $i--) {
        $start = $comments[$i].Extent.StartOffset
        $end = $comments[$i].Extent.EndOffset
        # Espaces qui précèdent le commentaire sur sa ligne.
        while ($start -gt 0 -and ($Text[$start - 1] -eq ' ' -or $Text[$start - 1] -eq "`t")) { $start-- }
        $lineStart = ($start -eq 0) -or ($Text[$start - 1] -eq "`n")
        $lineEnd = ($end -ge $Text.Length) -or ($Text[$end] -eq "`n")
        # Ligne qui ne contenait que le commentaire : on retire aussi son saut de ligne.
        if ($lineStart -and $lineEnd -and $end -lt $Text.Length) { $end++ }
        $Text = $Text.Remove($start, $end - $start)
    }
    return $Text
}

$body = (Remove-Comments $bodyWithComments).TrimEnd()

# Contrôle : mêmes jetons (type et texte) hors commentaires avant et après le retrait. Une suite
# de sauts de ligne, ou un saut de ligne en tête, compte pour un seul (les lignes de commentaire retirées laissent des
# lignes vides en moins, sans effet sur le langage).
function Get-ComparableTokens {
    param([string]$Text)
    $kinds = [System.Management.Automation.Language.TokenKind]
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($token in (Get-CodeTokens $Text)) {
        if ($token.Kind -eq $kinds::Comment) { continue }
        if ($token.Kind -eq $kinds::NewLine) {
            if ($result.Count -eq 0 -or $result[$result.Count - 1].Kind -eq $kinds::NewLine) { continue }
        }
        $result.Add($token)
    }
    return , $result.ToArray()
}
$before = Get-ComparableTokens $bodyWithComments
$after = Get-ComparableTokens $body
if ($before.Count -ne $after.Count) { throw ('Le retrait des commentaires a change le nombre de jetons : {0} -> {1}.' -f $before.Count, $after.Count) }
for ($i = 0; $i -lt $before.Count; $i++) {
    if ($before[$i].Kind -ne $after[$i].Kind -or $before[$i].Text -ne $after[$i].Text) {
        throw ('Le retrait des commentaires a modifie le jeton {0} : [{1}] -> [{2}].' -f $i, $before[$i].Text, $after[$i].Text)
    }
}

$content = $header + "& {`n" + $body + "`n} @parameters`n"
[System.IO.File]::WriteAllText($Output, $content, (New-Object System.Text.UTF8Encoding($false)))

# --- Contrôles du fichier généré ---------------------------------------------------------
$tokens = $null; $errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($Output, [ref]$tokens, [ref]$errors)
if ($errors.Count -gt 0) { throw ('Le fichier genere contient des erreurs de syntaxe : ' + $errors[0].Message) }
$percentCount = ([regex]::Matches($content, '%')).Count
$pad = [regex]::Matches($content, '%(BrowserPageURL|BrowserPageTitle)%').Count
if ($percentCount -ne 4 -or $pad -ne 2) { throw "Signes pourcent inattendus : $percentCount au total (attendu 4, soit 2 variables PAD), $pad variables PAD." }
$bytes = [System.IO.File]::ReadAllBytes($Output)
$nonAscii = @($bytes | Where-Object { $_ -gt 127 }).Count
if ($nonAscii -gt 0) { throw "Le fichier genere contient $nonAscii octets non ASCII (hors commentaires du script source ?)." }
if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { throw 'Le fichier genere commence par un BOM.' }
$hash = (Get-FileHash -LiteralPath $Output -Algorithm SHA256).Hash
Write-Output ('Genere : {0} ({1:N0} octets, {2} lignes, {3} jetons identiques au source) - SHA-256 {4}' -f $Output, $bytes.Length, ($content -split "`n").Count, $after.Count, $hash)
