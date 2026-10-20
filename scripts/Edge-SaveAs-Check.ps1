#Requires -Version 5.1
<#
.SYNOPSIS
    Repère une page Edge (URL + titre), ouvre le panneau Téléchargements, clique sur
    « Enregistrer sous » et vérifie que la fenêtre Windows d'enregistrement s'ouvre.

.DESCRIPTION
    Fonctionne en arrière-plan : tout passe par UI Automation (InvokePattern,
    SelectionItemPattern...). Aucun clic souris, aucune frappe clavier, Edge n'est pas
    mis au premier plan : l'utilisateur peut continuer à travailler pendant l'exécution.

    Étapes :
      1. Recherche, dans les fenêtres Edge, de l'onglet dont le titre correspond à
         -TitlePattern et l'URL à -UrlPattern. Si c'est un onglet inactif, il est
         sélectionné (sans activer la fenêtre).
      2. Ouverture du panneau Téléchargements de cette fenêtre via le bouton de la barre
         d'outils (sauf si l'invite « Que voulez-vous faire avec ... ? » est déjà visible).
      3. Clic sur le bouton « Enregistrer sous » de l'invite de téléchargement.
      4. Attente d'une nouvelle boîte de dialogue Windows (classe #32770)
         « Enregistrer sous » appartenant à Edge.

    Le journal est ouvert dès le début du script et se termine toujours par une ligne
    « RESULTAT ». Codes de sortie :
      0  OK : la fenêtre « Enregistrer sous » s'est ouverte
      1  erreur inattendue (détail dans le journal)
      2  Edge n'est pas lancé
      3  aucune page Edge ne correspond à l'URL et au titre
      4  bouton « Téléchargements » introuvable dans la barre d'outils
      5  bouton « Enregistrer sous » introuvable dans le panneau
      6  la fenêtre Windows « Enregistrer sous » ne s'est pas ouverte
      7  -RunHidden utilisé alors que le script n'est pas lancé depuis un fichier .ps1

.PARAMETER UrlPattern
    URL de la page, jokers acceptés (opérateur -like). Si Edge masque « https:// » ou
    « www. » dans la barre d'adresse, la comparaison se fait aussi sans ce préfixe.

.PARAMETER TitlePattern
    Titre de l'onglet, jokers acceptés (-like). Exemple : '*Copilot*'.

.PARAMETER FileNamePattern
    Facultatif : nom du fichier attendu dans l'invite de téléchargement, jokers acceptés.
    Exemple : 'Instruments*.xlsx'. Par défaut '*' = première invite trouvée.

.PARAMETER TimeoutSeconds
    Délai d'attente du bouton « Enregistrer sous » après ouverture du panneau.

.PARAMETER DialogTimeoutSeconds
    Délai d'attente de la fenêtre Windows « Enregistrer sous » après le clic.

.PARAMETER LogPath
    Fichier journal. Par défaut : dossier Logs à côté du script (ou dossier TEMP de
    l'utilisateur si impossible), un fichier par exécution.

.PARAMETER CloseDialog
    Ferme la fenêtre « Enregistrer sous » après vérification (équivaut à « Annuler » :
    aucun fichier n'est enregistré). Sans ce switch, la fenêtre reste ouverte.

.PARAMETER RunHidden
    Relance le script dans un processus PowerShell masqué (aucune console) et rend la
    main immédiatement. Le résultat est à lire dans le journal.

.PARAMETER DownloadsButtonNames
    Noms possibles du bouton Téléchargements de la barre d'outils (jokers acceptés).

.PARAMETER SaveAsButtonNames
    Noms possibles du bouton « Enregistrer sous » de l'invite (jokers acceptés).

.PARAMETER AddressBarNames
    Noms possibles de la barre d'adresse Edge (jokers acceptés).

.PARAMETER SaveDialogTitles
    Titres possibles de la fenêtre Windows d'enregistrement (jokers acceptés).

.EXAMPLE
    .\Edge-SaveAs-Check.ps1 -UrlPattern '*monappli.local/labo*' -TitlePattern '*Instruments*' -FileNamePattern 'Instruments*.xlsx'

.EXAMPLE
    # Totalement invisible : rend la main tout de suite, résultat dans le journal.
    .\Edge-SaveAs-Check.ps1 -UrlPattern '*monappli.local/labo*' -TitlePattern '*Instruments*' -RunHidden

.EXAMPLE
    # Depuis une tâche planifiée « Exécuter seulement si l'utilisateur est connecté » :
    powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Scripts\Edge-SaveAs-Check.ps1" -UrlPattern "*monappli.local/labo*" -TitlePattern "*Instruments*" -CloseDialog

.NOTES
    - Doit tourner dans la session de l'utilisateur connecté (pas en SYSTEM ni en
      « Exécuter même si l'utilisateur n'est pas connecté ») : UI Automation ne voit que
      le bureau de la session courante.
    - Fichier enregistré en UTF-8 avec BOM ; les libellés accentués sont en plus construits
      avec [char] pour fonctionner même si le fichier est réenregistré dans un autre encodage.
    - Compatible Windows PowerShell 5.1 et PowerShell 7.
    - Power Automate Desktop (action « Exécuter un script PowerShell ») : le script ne contient
      volontairement aucun signe pourcent, que PAD interprète comme une variable. Pour passer
      une variable PAD, remplacer une valeur par défaut ci-dessous par le nom de la variable
      PAD entouré de signes pourcent, entre apostrophes. -RunHidden est inutile dans PAD
      (PowerShell y tourne déjà sans fenêtre).
    - Quand la sortie est redirigée (PAD, « Exécuter la commande DOS »...), le détail va
      seulement dans le journal : la sortie standard contient uniquement le code de sortie
      et, en cas d'échec, le message d'erreur est écrit sur la sortie d'erreur.
#>
[CmdletBinding()]
param(
    # --- À adapter -------------------------------------------------------------
    [string]$UrlPattern = 'https://mon-application.example/*',
    [string]$TitlePattern = '*Titre de ma page*',
    [string]$FileNamePattern = '*',
    # ---------------------------------------------------------------------------
    [ValidateRange(1, 300)][int]$TimeoutSeconds = 20,
    [ValidateRange(1, 300)][int]$DialogTimeoutSeconds = 15,
    [string]$LogPath = '',
    [switch]$CloseDialog,
    [switch]$RunHidden,
    [string[]]$DownloadsButtonNames = @(('T{0}l{0}chargements*' -f [char]0xE9), 'Downloads*'),
    [string[]]$SaveAsButtonNames = @('Enregistrer sous', 'Save as'),
    [string[]]$AddressBarNames = @("Barre d'adresse*", 'Address and search bar*'),
    [string[]]$SaveDialogTitles = @('Enregistrer sous*', 'Save As*')
)

$ErrorActionPreference = 'Stop'

$script:StartTime = Get-Date
$script:ExitCode = 1
$script:ResultText = 'non termine'
$script:LogFile = $null
$script:BoundParams = $PSBoundParameters
$script:ScriptPath = $PSCommandPath
$script:MinimizedHandle = [IntPtr]::Zero
$script:DialogLeftOpen = $false
# Sortie redirigée (Power Automate Desktop...) : pas d'écho console, seul le code est renvoyé.
$script:QuietConsole = ($Host.Name -eq 'ConsoleHost') -and [Console]::IsOutputRedirected

# =============================================================================
# Journal et gestion des erreurs
# =============================================================================

function Initialize-Log {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        $dir = if ($script:ScriptPath) { Join-Path (Split-Path -Parent $script:ScriptPath) 'Logs' } else { Join-Path $env:TEMP 'EdgeSaveAsCheck' }
        $Path = Join-Path $dir ('EdgeSaveAsCheck_{0:yyyyMMdd_HHmmss}.log' -f $script:StartTime)
    }
    # Si le dossier demandé n'est pas accessible en écriture, repli sur le dossier TEMP.
    foreach ($candidate in @($Path, (Join-Path $env:TEMP (Split-Path -Leaf $Path)))) {
        try {
            $dir = Split-Path -Parent $candidate
            if ($dir -and -not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
            Add-Content -LiteralPath $candidate -Value ('=' * 100) -Encoding UTF8
            return $candidate
        }
        catch { }
    }
    return $null
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERREUR')][string]$Level = 'INFO'
    )
    $line = '{0:yyyy-MM-dd HH:mm:ss.fff} [{1,-6}] {2}' -f (Get-Date), $Level, $Message
    if ($script:LogFile) {
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 } catch { }
    }
    if (-not $script:QuietConsole) { Write-Host $line }
}

# Interrompt le traitement avec un code de sortie précis (intercepté par le catch final).
function Stop-Check {
    param([int]$Code, [string]$Message)
    $exception = [System.Exception]::new($Message)
    $exception.Data['ExitCode'] = $Code
    throw $exception
}

function Get-ExitCodeFromError {
    param($ErrorRecord)
    $exception = $ErrorRecord.Exception
    while ($exception) {
        if ($exception.Data.Contains('ExitCode')) { return [int]$exception.Data['ExitCode'] }
        $exception = $exception.InnerException
    }
    return 1
}

function Test-LikeAny {
    param([string]$Value, [string[]]$Patterns)
    foreach ($pattern in $Patterns) {
        if ($Value -like $pattern) { return $true }
    }
    return $false
}

# =============================================================================
# Relance masquée (-RunHidden)
# =============================================================================

function ConvertTo-PsLiteral {
    param([string]$Value)
    # PowerShell traite aussi ‘ ’ ‚ ‛ comme des apostrophes : on les double toutes.
    return "'" + ($Value -replace '([''\u2018\u2019\u201A\u201B])', '$1$1') + "'"
}

function ConvertTo-PsArgument {
    param([string]$Name, $Value)
    if ($Value -is [System.Management.Automation.SwitchParameter] -or $Value -is [bool]) {
        return '-{0}:${1}' -f $Name, ([bool]$Value).ToString().ToLower()
    }
    if ($Value -is [array]) {
        return '-{0} @({1})' -f $Name, ((@($Value) | ForEach-Object { ConvertTo-PsLiteral ([string]$_) }) -join ', ')
    }
    return '-{0} {1}' -f $Name, (ConvertTo-PsLiteral ([string]$Value))
}

function Start-HiddenInstance {
    if (-not $script:ScriptPath) {
        Stop-Check 7 '-RunHidden necessite que le script soit lance depuis un fichier .ps1.'
    }
    $arguments = New-Object System.Collections.Generic.List[string]
    foreach ($entry in $script:BoundParams.GetEnumerator()) {
        if ($entry.Key -in @('RunHidden', 'LogPath')) { continue }
        $arguments.Add((ConvertTo-PsArgument -Name $entry.Key -Value $entry.Value))
    }
    # Même journal pour le processus parent et le processus masqué.
    if ($script:LogFile) { $arguments.Add((ConvertTo-PsArgument -Name 'LogPath' -Value $script:LogFile)) }

    # -EncodedCommand évite tout problème de guillemets / accents dans les arguments.
    $command = '& {0} {1}; exit $LASTEXITCODE' -f (ConvertTo-PsLiteral $script:ScriptPath), ($arguments -join ' ')
    $encoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($command))
    $exe = (Get-Process -Id $PID).Path
    $process = Start-Process -FilePath $exe -WindowStyle Hidden -PassThru -ArgumentList @(
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-EncodedCommand', $encoded)
    Write-Log ('Relance en arriere-plan dans un processus masque (PID {0}). Resultat dans : {1}' -f $process.Id, $script:LogFile) 'OK'
}

# =============================================================================
# Win32 (énumération des fenêtres, sans dépendre du focus)
# =============================================================================

function Initialize-Win32 {
    if ('EdgeSaveAsCheck.Win32' -as [type]) { return }
    Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace EdgeSaveAsCheck
{
    public class WindowInfo
    {
        public IntPtr Handle;
        public IntPtr Owner;
        public uint ProcessId;
        public string ClassName;
        public string Title;
        public bool Visible;
    }

    public static class Win32
    {
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hWnd, StringBuilder text, int maxCount);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowTextLength(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hWnd, uint cmd);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int cmdShow);
        [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        private const uint GW_OWNER = 4;

        public static List<WindowInfo> GetTopLevelWindows()
        {
            List<WindowInfo> result = new List<WindowInfo>();
            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                StringBuilder className = new StringBuilder(256);
                GetClassName(hWnd, className, className.Capacity);
                StringBuilder title = new StringBuilder(GetWindowTextLength(hWnd) + 1);
                GetWindowText(hWnd, title, title.Capacity);
                uint processId;
                GetWindowThreadProcessId(hWnd, out processId);

                WindowInfo info = new WindowInfo();
                info.Handle = hWnd;
                info.Owner = GetWindow(hWnd, GW_OWNER);
                info.ProcessId = processId;
                info.ClassName = className.ToString();
                info.Title = title.ToString();
                info.Visible = IsWindowVisible(hWnd);
                result.Add(info);
                return true;
            }, IntPtr.Zero);
            return result;
        }
    }
}
'@
}

function Get-EdgeProcessIds {
    return @(Get-Process -Name msedge -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
}

# Fenêtres de navigateur Edge (y compris réduites).
function Get-EdgeBrowserWindows {
    $edgePids = Get-EdgeProcessIds
    foreach ($window in [EdgeSaveAsCheck.Win32]::GetTopLevelWindows()) {
        if ($window.ClassName -ne 'Chrome_WidgetWin_1' -or -not $window.Visible) { continue }
        if ($window.Owner -ne [IntPtr]::Zero -or [string]::IsNullOrEmpty($window.Title)) { continue }
        if ($edgePids -notcontains [int]$window.ProcessId) { continue }
        try { [System.Windows.Automation.AutomationElement]::FromHandle($window.Handle) } catch { }
    }
}

# Boîtes de dialogue Windows (#32770) liées à Edge : processus msedge (le dialogue peut
# venir d'un processus utilitaire créé au moment du clic) ou fenêtre Edge propriétaire.
function Get-EdgeDialogWindows {
    param([IntPtr]$OwnerHandle)
    $edgePids = Get-EdgeProcessIds
    foreach ($window in [EdgeSaveAsCheck.Win32]::GetTopLevelWindows()) {
        if ($window.ClassName -ne '#32770' -or -not $window.Visible) { continue }
        if (($edgePids -contains [int]$window.ProcessId) -or ($window.Owner -eq $OwnerHandle)) { $window }
    }
}

# =============================================================================
# UI Automation
# =============================================================================

function New-ControlTypeCondition {
    param($ControlType)
    return [System.Windows.Automation.PropertyCondition]::new(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $ControlType)
}

function Initialize-Uia {
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    $controlType = [System.Windows.Automation.ControlType]
    $script:Descendants = [System.Windows.Automation.TreeScope]::Descendants
    $script:CondAll = [System.Windows.Automation.Condition]::TrueCondition
    $script:CondEdit = New-ControlTypeCondition $controlType::Edit
    $script:CondTabItem = New-ControlTypeCondition $controlType::TabItem
    $script:CondDocument = New-ControlTypeCondition $controlType::Document
    $script:CondClickable = [System.Windows.Automation.OrCondition]::new([System.Windows.Automation.Condition[]]@(
            (New-ControlTypeCondition $controlType::Button),
            (New-ControlTypeCondition $controlType::SplitButton),
            (New-ControlTypeCondition $controlType::MenuItem)))
}

function Get-ElementName {
    param($Element)
    try { return [string]$Element.Current.Name } catch { return '' }
}

function Get-ElementValue {
    param($Element)
    try {
        $pattern = $null
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) {
            return [string]$pattern.Current.Value
        }
    }
    catch { }
    return $null
}

# Vrai si l'élément appartient au contenu d'une page web (sous un « Document »),
# faux s'il appartient à l'interface du navigateur (barre d'outils, onglets).
function Test-InWebContent {
    param($Element)
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $node = $Element
    for ($depth = 0; $depth -lt 40 -and $node; $depth++) {
        try {
            if ($node.Current.ControlType -eq [System.Windows.Automation.ControlType]::Document) { return $true }
            $node = $walker.GetParent($node)
        }
        catch { return $false }
    }
    return $false
}

# -PreferBrowserUi : éléments de l'interface Edge d'abord, puis ceux du contenu web.
function Find-NamedElements {
    param(
        [object[]]$Roots,
        $Condition,
        [string[]]$NamePatterns,
        [switch]$PreferBrowserUi
    )
    $found = New-Object System.Collections.Generic.List[object]
    $inWebContent = New-Object System.Collections.Generic.List[object]
    foreach ($root in $Roots) {
        if (-not $root) { continue }
        try { $candidates = $root.FindAll($script:Descendants, $Condition) } catch { continue }
        foreach ($element in $candidates) {
            if (-not (Test-LikeAny (Get-ElementName $element) $NamePatterns)) { continue }
            if ($PreferBrowserUi -and (Test-InWebContent $element)) { $inWebContent.Add($element) }
            else { $found.Add($element) }
        }
    }
    $found.AddRange($inWebContent)
    return $found.ToArray()
}

function Invoke-UiaElement {
    param($Element)
    $pattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) {
        $pattern.Invoke(); return 'Invoke'
    }
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$pattern)) {
        $pattern.Toggle(); return 'Toggle'
    }
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$pattern)) {
        $pattern.Expand(); return 'Expand'
    }
    throw ("Aucune action UI Automation disponible sur '{0}'." -f (Get-ElementName $Element))
}

# =============================================================================
# Étape 1 : page Edge cible (URL + titre)
# =============================================================================

function Test-UrlMatch {
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return $false }
    if ($Url -like $UrlPattern) { return $true }
    # Edge peut masquer « https:// » et « www. » : on compare aussi sans ce préfixe.
    $prefix = '^[a-zA-Z][a-zA-Z0-9+.-]*://(www\.)?'
    return (($Url -replace $prefix, '') -like ($UrlPattern -replace $prefix, ''))
}

function Get-AddressBarValue {
    param($Window)
    try { $edits = $Window.FindAll($script:Descendants, $script:CondEdit) } catch { return $null }
    $fallback = $null
    foreach ($edit in $edits) {
        if (Test-LikeAny (Get-ElementName $edit) $AddressBarNames) { return Get-ElementValue $edit }
        # Repli si le nom diffère : 1er champ de saisie de l'interface Edge (hors page web).
        if (-not $fallback -and -not (Test-InWebContent $edit)) { $fallback = $edit }
    }
    if ($fallback) { return Get-ElementValue $fallback }
    return $null
}

function Test-TabSelected {
    param($Tab)
    try {
        $pattern = $null
        if ($Tab.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$pattern)) {
            return [bool]$pattern.Current.IsSelected
        }
    }
    catch { }
    return $false
}

function Select-Tab {
    param($Tab)
    try {
        $pattern = $null
        if ($Tab.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$pattern)) {
            $pattern.Select(); return $true
        }
        if ($Tab.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) {
            $pattern.Invoke(); return $true
        }
    }
    catch { Write-Log ("Selection de l'onglet impossible : {0}" -f $_.Exception.Message) 'WARN' }
    return $false
}

function New-PageResult {
    param($Window, [string]$Title, [string]$Url)
    return [pscustomobject]@{
        Window = $Window
        Handle = [IntPtr]$Window.Current.NativeWindowHandle
        Title  = $Title
        Url    = $Url
    }
}

function Find-TargetEdgePage {
    param([switch]$LogDetails)
    foreach ($window in @(Get-EdgeBrowserWindows)) {
        $windowName = Get-ElementName $window

        # 1) Onglet actif : le titre de la fenêtre commence par le titre de la page.
        $url = Get-AddressBarValue $window
        if ($LogDetails) { Write-Log ("Fenetre Edge : '{0}' | URL de l'onglet actif : '{1}'" -f $windowName, $url) }
        if ($windowName -like "$TitlePattern*" -and (Test-UrlMatch $url)) {
            return New-PageResult -Window $window -Title $windowName -Url $url
        }

        # 2) Onglets inactifs dont le titre correspond : on les active un par un pour lire l'URL.
        $allTabs = @()
        try { $allTabs = @($window.FindAll($script:Descendants, $script:CondTabItem)) } catch { }
        $candidates = @($allTabs | Where-Object {
                (Get-ElementName $_) -like "$TitlePattern*" -and -not (Test-TabSelected $_) -and -not (Test-InWebContent $_)
            })
        if ($candidates.Count -eq 0) { continue }

        $initialTab = $allTabs | Where-Object { (Test-TabSelected $_) -and -not (Test-InWebContent $_) } | Select-Object -First 1
        foreach ($tab in $candidates) {
            $tabName = Get-ElementName $tab
            if (-not (Select-Tab $tab)) { continue }
            Start-Sleep -Milliseconds 500
            $url = Get-AddressBarValue $window
            Write-Log ("  Onglet '{0}' active | URL : '{1}'" -f $tabName, $url)
            if (Test-UrlMatch $url) {
                return New-PageResult -Window $window -Title $tabName -Url $url
            }
        }
        if ($initialTab) {
            [void](Select-Tab $initialTab)
            Write-Log ("  Aucun onglet ne correspond : retour a l'onglet '{0}'" -f (Get-ElementName $initialTab))
        }
    }
    return $null
}

# =============================================================================
# Étapes 2-3 : panneau Téléchargements et bouton « Enregistrer sous »
# =============================================================================

# Le panneau peut être exposé dans l'arbre de la fenêtre Edge ou comme fenêtre
# contextuelle séparée appartenant à celle-ci : on cherche dans les deux.
function Get-PanelSearchRoots {
    param($Page)
    $roots = New-Object System.Collections.Generic.List[object]
    foreach ($window in [EdgeSaveAsCheck.Win32]::GetTopLevelWindows()) {
        if ($window.Owner -ne $Page.Handle -or -not $window.Visible -or $window.ClassName -notlike 'Chrome_WidgetWin*') { continue }
        try { $roots.Add([System.Windows.Automation.AutomationElement]::FromHandle($window.Handle)) } catch { }
    }
    $roots.Add($Page.Window)
    return $roots.ToArray()
}

# Remonte depuis le bouton jusqu'à l'élément de téléchargement et renvoie le texte
# correspondant à -FileNamePattern (ex. « Que voulez-vous faire avec Instruments (7).xlsx ? »).
function Get-PromptTextForButton {
    param($Button)
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $node = $Button
    for ($level = 1; $level -le 3; $level++) {
        try {
            $node = $walker.GetParent($node)
            if (-not $node) { return $null }
            $names = @(Get-ElementName $node) + @($node.FindAll($script:Descendants, $script:CondAll) | ForEach-Object { Get-ElementName $_ })
        }
        catch { return $null }
        # Conteneur commun à plusieurs invites : on est sorti de l'élément de ce bouton.
        if (@($names | Where-Object { Test-LikeAny $_ $SaveAsButtonNames }).Count -gt 1) { return $null }
        $match = $names | Where-Object { $_ -like "*$FileNamePattern*" } | Select-Object -First 1
        if ($match) { return $match }
    }
    return $null
}

function Find-SaveAsButton {
    param($Page)
    $buttons = @(Find-NamedElements -Roots (Get-PanelSearchRoots $Page) -Condition $script:CondClickable -NamePatterns $SaveAsButtonNames)
    foreach ($button in $buttons) {
        if ($FileNamePattern -eq '*') { return [pscustomobject]@{ Element = $button; PromptText = $null } }
        $text = Get-PromptTextForButton $button
        if ($text) { return [pscustomobject]@{ Element = $button; PromptText = $text } }
    }
    return $null
}

function Wait-SaveAsButton {
    param($Page, [int]$Seconds)
    $deadline = (Get-Date).AddSeconds($Seconds)
    do {
        $result = Find-SaveAsButton -Page $Page
        if ($result) { return $result }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)
    return $null
}

function Open-DownloadsPanel {
    param($Page)
    $buttons = @(Find-NamedElements -Roots @($Page.Window) -Condition $script:CondClickable -NamePatterns $DownloadsButtonNames -PreferBrowserUi)
    if ($buttons.Count -eq 0) { return $false }
    $action = Invoke-UiaElement $buttons[0]
    Write-Log ("Bouton '{0}' actionne ({1}) : ouverture du panneau Telechargements" -f (Get-ElementName $buttons[0]), $action)
    return $true
}

# Liste ce qu'UI Automation voit, pour ajuster les libellés si Edge les nomme autrement.
function Write-UiaDiagnostic {
    param($Page)
    $roots = Get-PanelSearchRoots $Page
    foreach ($entry in @(@{ Label = 'Boutons'; Condition = $script:CondClickable }, @{ Label = 'Documents'; Condition = $script:CondDocument })) {
        $names = foreach ($root in $roots) {
            try { $root.FindAll($script:Descendants, $entry.Condition) | ForEach-Object { Get-ElementName $_ } } catch { }
        }
        $names = @($names | Where-Object { $_ } | Select-Object -Unique | Select-Object -First 100)
        Write-Log ('Diagnostic - {0} vus par UI Automation : {1}' -f $entry.Label, ($names -join ' | ')) 'WARN'
    }
}

# =============================================================================
# Étape 4 : fenêtre Windows « Enregistrer sous »
# =============================================================================

function Wait-SaveDialog {
    param([IntPtr]$OwnerHandle, [IntPtr[]]$Exclude, [int]$Seconds)
    $deadline = (Get-Date).AddSeconds($Seconds)
    $reported = @{}
    do {
        foreach ($dialog in @(Get-EdgeDialogWindows -OwnerHandle $OwnerHandle)) {
            if ($Exclude -contains $dialog.Handle) { continue }
            if (Test-LikeAny $dialog.Title $SaveDialogTitles) { return $dialog }
            if (-not $reported.ContainsKey($dialog.Handle)) {
                $reported[$dialog.Handle] = $true
                Write-Log ("Boite de dialogue Edge ignoree (titre inattendu) : '{0}'" -f $dialog.Title) 'WARN'
            }
        }
        Start-Sleep -Milliseconds 250
    } while ((Get-Date) -lt $deadline)
    return $null
}

function Get-DialogFileName {
    param([IntPtr]$Handle)
    try {
        $dialog = [System.Windows.Automation.AutomationElement]::FromHandle($Handle)
        $condition = [System.Windows.Automation.AndCondition]::new([System.Windows.Automation.Condition[]]@(
                $script:CondEdit,
                [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::AutomationIdProperty, '1001')))
        $edit = $dialog.FindFirst($script:Descendants, $condition)
        if ($edit) { return Get-ElementValue $edit }
    }
    catch { }
    return $null
}

function Close-SaveDialog {
    param([IntPtr]$Handle)
    # WM_CLOSE = « Annuler » : aucun fichier n'est enregistré.
    [void][EdgeSaveAsCheck.Win32]::PostMessage($Handle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
    $deadline = (Get-Date).AddSeconds(5)
    while ([EdgeSaveAsCheck.Win32]::IsWindow($Handle) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 200 }
    return -not [EdgeSaveAsCheck.Win32]::IsWindow($Handle)
}

# =============================================================================
# Traitement principal
# =============================================================================

function Invoke-EdgeSaveAsCheck {
    Initialize-Uia
    Initialize-Win32

    if ((Get-EdgeProcessIds).Count -eq 0) {
        Stop-Check 2 "Microsoft Edge n'est pas lance (aucun processus msedge)."
    }

    # --- 1. Page cible ---------------------------------------------------------
    # Plusieurs passes : au premier accès, Edge construit son arbre d'accessibilité.
    $page = $null
    for ($attempt = 1; $attempt -le 3 -and -not $page; $attempt++) {
        if ($attempt -gt 1) { Start-Sleep -Milliseconds 1500 }
        $page = Find-TargetEdgePage -LogDetails:($attempt -eq 1)
    }
    if (-not $page) {
        Stop-Check 3 ("Aucune page Edge ne correspond a URL='{0}' et Titre='{1}' (voir les fenetres listees ci-dessus)." -f $UrlPattern, $TitlePattern)
    }
    Write-Log ("Page cible : '{0}' | {1} | HWND 0x{2:X}" -f $page.Title, $page.Url, $page.Handle.ToInt64()) 'OK'

    # Une fenêtre réduite ne peut pas afficher le panneau : restauration SANS activation.
    if ([EdgeSaveAsCheck.Win32]::IsIconic($page.Handle)) {
        [void][EdgeSaveAsCheck.Win32]::ShowWindow($page.Handle, 4)   # SW_SHOWNOACTIVATE
        $script:MinimizedHandle = $page.Handle
        Start-Sleep -Milliseconds 700
        Write-Log 'Fenetre Edge reduite : restauree sans prendre le focus'
    }

    # --- 2. Panneau Téléchargements ---------------------------------------------
    # Edge ouvre souvent le panneau tout seul au début du téléchargement : on vérifie
    # d'abord, car actionner le bouton refermerait un panneau déjà ouvert.
    $saveAs = Wait-SaveAsButton -Page $page -Seconds 2
    if ($saveAs) {
        Write-Log 'Panneau Telechargements deja ouvert sur une invite de telechargement'
    }
    else {
        if (-not (Open-DownloadsPanel -Page $page)) {
            Write-UiaDiagnostic -Page $page
            Stop-Check 4 'Bouton Telechargements introuvable dans la barre d outils Edge (parametre Edge : Apparence > Bouton Telechargements).'
        }
        $saveAs = Wait-SaveAsButton -Page $page -Seconds $TimeoutSeconds
    }
    if (-not $saveAs) {
        Write-UiaDiagnostic -Page $page
        $detail = if ($FileNamePattern -ne '*') { " pour un fichier '$FileNamePattern'" } else { '' }
        Stop-Check 5 ("Bouton '{0}' introuvable dans le panneau Telechargements{1}." -f ($SaveAsButtonNames -join "' / '"), $detail)
    }
    if ($saveAs.PromptText) { Write-Log ("Invite de telechargement : '{0}'" -f $saveAs.PromptText) }

    # --- 3. Clic « Enregistrer sous » --------------------------------------------
    $existing = @(Get-EdgeDialogWindows -OwnerHandle $page.Handle | ForEach-Object { $_.Handle })
    if ($existing.Count -gt 0) {
        Write-Log ('{0} boite(s) de dialogue Edge deja ouverte(s) avant le clic (ignoree(s))' -f $existing.Count) 'WARN'
    }
    $action = Invoke-UiaElement $saveAs.Element
    Write-Log ("Clic sur '{0}' ({1})" -f (Get-ElementName $saveAs.Element), $action)

    # --- 4. Vérification de la fenêtre Windows ----------------------------------
    $dialog = Wait-SaveDialog -OwnerHandle $page.Handle -Exclude $existing -Seconds $DialogTimeoutSeconds
    if (-not $dialog) {
        Stop-Check 6 ("La fenetre Windows d'enregistrement ne s'est pas ouverte dans les {0} s." -f $DialogTimeoutSeconds)
    }
    Write-Log ("Fenetre Windows detectee : '{0}' (classe {1}, HWND 0x{2:X}, PID {3})" -f $dialog.Title, $dialog.ClassName, $dialog.Handle.ToInt64(), $dialog.ProcessId) 'OK'
    $fileName = Get-DialogFileName $dialog.Handle
    if ($fileName) { Write-Log ("Nom de fichier propose : '{0}'" -f $fileName) }

    if ($CloseDialog) {
        if (Close-SaveDialog $dialog.Handle) { Write-Log 'Fenetre fermee (Annuler : aucun fichier enregistre)' }
        else { Write-Log 'La fenetre ne s est pas fermee dans les 5 s' 'WARN'; $script:DialogLeftOpen = $true }
    }
    else {
        $script:DialogLeftOpen = $true
    }
    $script:ResultText = "fenetre '$($dialog.Title)' ouverte"
}

# =============================================================================
# Point d'entrée : un seul try/catch/finally, journal garanti
# =============================================================================

try {
    $script:LogFile = Initialize-Log -Path $LogPath
    Write-Log ('Demarrage - PID {0}, PowerShell {1}, utilisateur {2}, journal : {3}' -f $PID, $PSVersionTable.PSVersion, $env:USERNAME, $script:LogFile)
    Write-Log ("Parametres - URL '{0}' | Titre '{1}' | Fichier '{2}' | CloseDialog {3}" -f $UrlPattern, $TitlePattern, $FileNamePattern, [bool]$CloseDialog)

    if ($RunHidden) {
        Start-HiddenInstance
        $script:ResultText = 'relance en arriere-plan'
    }
    else {
        Invoke-EdgeSaveAsCheck
    }
    $script:ExitCode = 0
}
catch {
    $script:ExitCode = Get-ExitCodeFromError $_
    $script:ResultText = $_.Exception.Message
    Write-Log $_.Exception.Message 'ERREUR'
    if ($script:ExitCode -eq 1) { Write-Log ('Detail : {0}' -f $_.ScriptStackTrace) 'ERREUR' }
    if ($script:QuietConsole) { [Console]::Error.WriteLine($_.Exception.Message) }
}
finally {
    # Fenêtre Edge réduite au départ : on la réduit à nouveau, sauf si le dialogue reste ouvert.
    if ($script:MinimizedHandle -ne [IntPtr]::Zero -and -not $script:DialogLeftOpen) {
        try { [void][EdgeSaveAsCheck.Win32]::ShowWindow($script:MinimizedHandle, 7) } catch { }   # SW_SHOWMINNOACTIVE
    }
    $elapsed = ((Get-Date) - $script:StartTime).TotalSeconds
    $level = if ($script:ExitCode -eq 0) { 'OK' } else { 'ERREUR' }
    Write-Log ('RESULTAT : code {0} - {1} - duree {2:N1} s' -f $script:ExitCode, $script:ResultText, $elapsed) $level
}
if ($script:QuietConsole) { Write-Output $script:ExitCode }
exit $script:ExitCode
