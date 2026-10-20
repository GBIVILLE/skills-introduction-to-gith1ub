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
#      10 script introuvable (ScriptError liste les fichiers Edge-SaveAs-Check* du dossier)
#      11 script trouvé mais illisible (droits, fichier verrouillé)
#      12 script invalide ou abîmé (vérifier son empreinte SHA-256)
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

function Stop-Launcher {
    param([int]$Code, [string]$Message)
    [Console]::Error.WriteLine($Message)
    Write-Output $Code
    exit $Code
}

# Windows masque souvent les extensions connues : un fichier « renommé » en .ps1 peut
# s'appeler en réalité Edge-SaveAs-Check.ps1.txt. On accepte donc les deux noms.
$resolvedPath = @($ScriptPath, ($ScriptPath + '.txt')) |
    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
    Select-Object -First 1
if (-not $resolvedPath) {
    $folder = Split-Path -Parent $ScriptPath
    $listing = if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        'le dossier n existe pas'
    }
    else {
        $names = @(Get-ChildItem -LiteralPath $folder -Filter 'Edge-SaveAs-Check*' -Name)
        if ($names.Count -gt 0) { 'fichiers presents : ' + ($names -join ', ') } else { 'aucun fichier Edge-SaveAs-Check* dans le dossier' }
    }
    Stop-Launcher 10 ('Script introuvable : {0} ({1}).' -f $ScriptPath, $listing)
}

# Exécution en mémoire : ni la stratégie d'exécution ni le marquage « fichier venant
# d'Internet » ne bloquent le script.
try { $code = [System.IO.File]::ReadAllText($resolvedPath) }
catch { Stop-Launcher 11 ('Script illisible : {0} - {1}' -f $resolvedPath, $_.Exception.Message) }

try { $scriptBlock = [scriptblock]::Create($code) }
catch { Stop-Launcher 12 ('Script invalide ou abime (verifier son empreinte SHA-256) : {0} - {1}' -f $resolvedPath, $_.Exception.Message) }

& $scriptBlock @parameters
