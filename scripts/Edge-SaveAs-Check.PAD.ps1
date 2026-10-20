# =============================================================================
# Lanceur Power Automate Desktop pour Edge-SaveAs-Check.ps1
#
# À coller tel quel dans l'action « Exécuter un script PowerShell ».
#  - Créer avant cette action les variables PAD UrlCible et TitreCible (action
#    « Définir la variable »). Jokers acceptés, apostrophes autorisées.
#  - Action PAD : désactiver « Échec après expiration du délai » ou mettre au moins 120 s.
#  - Résultat : PowershellOutput contient le code (0 = fenêtre « Enregistrer sous »
#    ouverte), ScriptError contient le message en cas d'échec.
#    Codes 1 à 9 : voir Edge-SaveAs-Check.ps1. Codes du lanceur :
#      10 script introuvable (ScriptError liste le contenu du dossier)
#      11 script trouvé mais illisible (droits, fichier verrouillé)
#      12 script invalide ou abîmé (ScriptError donne la ligne en erreur)
#      13 script incomplet : il s'est terminé sans renvoyer de code
#  - Aucun signe pourcent ne doit apparaître ici, sauf autour des variables PAD.
# =============================================================================

$ScriptPath = 'C:\Scripts\Edge-SaveAs-Check.ps1'   # <-- emplacement du script

$parameters = @{
    UrlPattern      = @'
%UrlCible%
'@
    TitlePattern    = @'
%TitreCible%
'@
    FileNamePattern = '*'                  # ex. 'Instruments*.xlsx'
    LogPath         = 'C:\Scripts\Logs'    # dossier : un journal par exécution
    CloseDialog     = $false               # $true : ferme la fenêtre après vérification
}

# Mode linguistique restreint (AppLocker / WDAC) : le script ne peut pas y fonctionner.
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-Output 8
    Write-Error -ErrorAction Continue ('PowerShell est en mode {0} (strategie de securite) : script inutilisable.' -f $ExecutionContext.SessionState.LanguageMode)
    exit 8
}

# Le code est écrit avant le message : PAD le reçoit même si l'écriture du message échoue.
function Stop-Launcher {
    param([int]$Code, [string]$Message)
    Write-Output $Code
    try { [Console]::Error.WriteLine($Message) } catch { }
    exit $Code
}

# Windows masque souvent les extensions connues : un fichier « renommé » en .ps1 peut
# s'appeler en réalité Edge-SaveAs-Check.ps1.txt. On accepte donc les deux noms.
$resolvedPath = @($ScriptPath, ($ScriptPath + '.txt')) |
    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
    Select-Object -First 1
if (-not $resolvedPath) {
    $folder = Split-Path -Parent $ScriptPath
    $listing = if ($ScriptPath.Contains('"')) {
        'le chemin contient des guillemets, a retirer'
    }
    elseif (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        'le dossier n existe pas'
    }
    else {
        # Noms entre crochets pour voir les espaces ; dossiers signalés.
        $entries = @(Get-ChildItem -LiteralPath $folder -Force | Select-Object -First 20 | ForEach-Object {
                if ($_.PSIsContainer) { '[' + $_.Name + '] (dossier)' } else { '[' + $_.Name + ']' }
            })
        if ($entries.Count -gt 0) { 'contenu du dossier : ' + ($entries -join ' ') } else { 'le dossier est vide' }
    }
    Stop-Launcher 10 ('Script introuvable pour {0}\{1} : {2} ({3}).' -f $env:COMPUTERNAME, $env:USERNAME, $ScriptPath, $listing)
}

# Exécution en mémoire : ni la stratégie d'exécution ni le marquage « fichier venant
# d'Internet » ne bloquent le script.
try { $code = [System.IO.File]::ReadAllText($resolvedPath) }
catch { Stop-Launcher 11 ('Script illisible : {0} - {1}' -f $resolvedPath, $_.Exception.Message) }

# Caractères parasites avant #Requires (BOM recopié depuis GitHub, BOM mal décodé, espace
# invisible...) : ils empêchent l'analyse du bloc param (erreur « ligne 105 »). On les retire.
$start = $code.IndexOf('#Requires')
if ($start -gt 0 -and $start -le 16) { $code = $code.Substring($start) }

try { $scriptBlock = [scriptblock]::Create($code) }
catch { Stop-Launcher 12 ('Script invalide ou abime (le retelecharger depuis GitHub) : {0} - {1}' -f $resolvedPath, $_.Exception.Message) }

& $scriptBlock @parameters

# Le script se termine toujours par « exit <code> », qui arrête aussi ce lanceur :
# arriver ici signifie qu'il est incomplet (fichier tronqué ou autre contenu).
Stop-Launcher 13 ('Le script s est termine sans renvoyer de code (fichier tronque ou incomplet ?) : {0}' -f $resolvedPath)
