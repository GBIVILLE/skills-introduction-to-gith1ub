# =============================================================================
# Lanceur Power Automate Desktop pour Edge-SaveAs-Check.ps1
#
# À coller tel quel dans l'action « Exécuter un script PowerShell ».
#  - Créer avant cette action les variables PAD UrlCible et TitreCible (action
#    « Définir la variable »). Jokers acceptés, apostrophes autorisées.
#  - Action PAD : désactiver « Échec après expiration du délai » ou mettre au moins 120 s.
#  - Résultat : PowershellOutput contient le code (0 = fenêtre « Enregistrer sous »
#    ouverte), ScriptError contient le message en cas d'échec.
#    Codes 1 à 9 : voir Edge-SaveAs-Check.ps1. Code 10 : script introuvable ou illisible.
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

# Exécution en mémoire : ni la stratégie d'exécution ni le marquage « fichier venant
# d'Internet » ne bloquent le script.
try {
    $scriptBlock = [scriptblock]::Create([System.IO.File]::ReadAllText($ScriptPath))
}
catch {
    [Console]::Error.WriteLine(('Script introuvable ou illisible : {0} - {1}' -f $ScriptPath, $_.Exception.Message))
    Write-Output 10
    exit 10
}
& $scriptBlock @parameters
