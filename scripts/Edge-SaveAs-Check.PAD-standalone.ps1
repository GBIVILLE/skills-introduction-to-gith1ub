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

$parameters = @{
    UrlPattern      = @'
%BrowserPageURL%
'@
    TitlePattern    = @'
%BrowserPageTitle%
'@
    FileNamePattern = '*'                  # ex. 'Instruments*.xlsx'
    LogPath         = $(if ($env:LOCALAPPDATA) { Join-Path $env:LOCALAPPDATA 'EdgeSaveAsCheck\Logs' } else { '' })
    CloseDialog     = $false               # $true : ferme la fenetre apres verification
    # AllowRealMouse = $true              # dernier recours : vrai clic souris (curseur deplace un instant)
}
& {
<#
.SYNOPSIS
    Repère une page Edge (URL + titre), ouvre le panneau Téléchargements, clique sur
    « Enregistrer sous » et vérifie que la fenêtre Windows d'enregistrement s'ouvre.

.DESCRIPTION
    Fonctionne en arrière-plan : UI Automation, accessibilité MSAA et messages Windows postés
    directement aux fenêtres d'Edge. Le script ne déplace pas le curseur, n'envoie aucune frappe
    au clavier et ne met pas Edge au premier plan (seul -AllowRealMouse, désactivé par défaut,
    déplace brièvement le curseur). Ce que l'utilisateur voit quand même, car c'est Edge qui le
    fait : le panneau Téléchargements qui s'ouvre, l'onglet cible qui devient actif s'il ne
    l'était pas, et la fenêtre « Enregistrer sous » (laissée ouverte, ou fermée avec -CloseDialog).
    Les messages ne sont jamais postés qu'à des fenêtres d'Edge, même si une autre application
    le recouvre.

    Indépendant de la langue d'Edge : les éléments sont reconnus par leur classe UI Automation
    (non traduite) et la boîte « Enregistrer sous » par sa classe de boîte de dialogue Windows ;
    les noms français et anglais ne servent que de repli. Le titre de l'onglet (-TitlePattern)
    est celui de la page, donné par l'appelant.

    Étapes :
      1. Recherche, dans les fenêtres Edge, de l'onglet dont le titre correspond à
         -TitlePattern et l'URL à -UrlPattern. Si c'est un onglet inactif, il est
         sélectionné (sans activer la fenêtre).
      2. Ouverture du panneau Téléchargements de cette fenêtre via le bouton de la barre
         d'outils (sauf si l'invite « Que voulez-vous faire avec ... ? » est déjà visible).
      3. Clic sur le bouton « Enregistrer sous » de l'invite de téléchargement. Le panneau
         d'Edge est une page web qui peut ignorer le clic simulé d'UI Automation : les
         méthodes sont essayées l'une après l'autre (message souris natif, Invoke,
         accDoDefaultAction, message souris à la fenêtre enfant, touche Entrée) jusqu'à
         l'ouverture de la fenêtre ; le journal indique laquelle a fonctionné.
      4. Attente d'une NOUVELLE fenêtre Windows « Enregistrer sous » (absente avant le clic).

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
      8  PowerShell en mode linguistique restreint (stratégie de sécurité) : script inutilisable
      9  paramètre obligatoire vide (UrlPattern ou TitlePattern, ex. variable PAD vide)

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
    Délai final d'attente de la fenêtre Windows « Enregistrer sous » quand aucune méthode de
    clic n'a produit de fenêtre.

.PARAMETER ClickWaitSeconds
    Délai d'attente de la fenêtre après chaque méthode de clic, avant d'essayer la suivante.

.PARAMETER AllowRealMouse
    Ajoute en dernier recours un vrai clic souris (le curseur se déplace un instant) : à
    n'utiliser que si les méthodes d'arrière-plan échouent. Le panneau doit être visible.

.PARAMETER LogPath
    Fichier journal, ou dossier dans lequel un fichier est créé à chaque exécution.
    Par défaut : dossier Logs à côté du script, ou dossier TEMP\EdgeSaveAsCheck si le script
    n'est pas lancé depuis un fichier (lanceur PAD) ou si le dossier est en lecture seule.

.PARAMETER CloseDialog
    Ferme la fenêtre « Enregistrer sous » après vérification (équivaut à « Annuler » :
    aucun fichier n'est enregistré). Sans ce switch, la fenêtre reste ouverte.

.PARAMETER RunHidden
    Relance le script dans un processus PowerShell masqué (aucune console) et rend la
    main immédiatement. Le résultat est à lire dans le journal.

.PARAMETER DownloadsButtonClasses
    Classe UI Automation du bouton Téléchargements de la barre d'outils. Les classes ne sont
    pas traduites : c'est le critère principal, qui marche quelle que soit la langue d'Edge.

.PARAMETER SaveAsButtonClasses
    Classe UI Automation du bouton « Enregistrer sous » de l'invite (critère principal).

.PARAMETER AddressBarClasses
    Classe UI Automation de la barre d'adresse (critère principal).

.PARAMETER LogRetentionDays
    Les journaux du dossier de journaux plus anciens que ce nombre de jours sont supprimés au
    démarrage (30 par défaut, 0 pour tout conserver).

.PARAMETER DownloadsButtonNames
    Noms possibles du bouton Téléchargements (jokers acceptés). Repli, utilisé seulement si la
    classe ne correspond à rien : dépend de la langue d'Edge.

.PARAMETER SaveAsButtonNames
    Noms possibles du bouton « Enregistrer sous » (jokers acceptés). Repli dépendant de la langue.

.PARAMETER AddressBarNames
    Noms possibles de la barre d'adresse Edge (jokers acceptés). Repli dépendant de la langue.

.PARAMETER SaveDialogTitles
    Titres possibles de la fenêtre Windows d'enregistrement (jokers acceptés). Facultatif : une
    nouvelle boîte de dialogue Windows (classe #32770) issue d'Edge est reconnue sans titre.

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
    - Fichier en UTF-8 sans BOM : un BOM copié depuis GitHub se retrouve devant #Requires et
      empêche l'analyse du script. Les accents ne figurent que dans les commentaires (les
      libellés accentués sont construits avec [char]) : le script marche quel que soit l'encodage.
    - Compatible Windows PowerShell 5.1 et PowerShell 7.
    - Power Automate Desktop, deux façons (voir README.md) : coller dans l'action « Exécuter un
      script PowerShell » la version autonome Edge-SaveAs-Check.PAD-standalone.ps1 (le script
      voyage avec le flux, rien à installer sur les postes), ou le lanceur
      Edge-SaveAs-Check.PAD.ps1 qui lit ce fichier à un emplacement partagé. Le script ne
      contient aucun signe pourcent (PAD l'interpréterait comme une variable). -RunHidden est
      inutile dans PAD, qui exécute déjà PowerShell sans fenêtre.
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
    [ValidateRange(1, 60)][int]$ClickWaitSeconds = 3,
    [string]$LogPath = '',
    [switch]$CloseDialog,
    [switch]$AllowRealMouse,
    [switch]$RunHidden,
    [string[]]$DownloadsButtonClasses = @('ToolbarDownloadsButton'),
    [string[]]$SaveAsButtonClasses = @('saveAs'),
    [string[]]$AddressBarClasses = @('OmniboxViewViews'),
    [ValidateRange(0, 3650)][int]$LogRetentionDays = 30,
    [string[]]$DownloadsButtonNames = @(('T{0}l{0}chargements*' -f [char]0xE9), 'Downloads*'),
    [string[]]$SaveAsButtonNames = @('Enregistrer sous', 'Save as'),
    [string[]]$AddressBarNames = @("Barre d'adresse*", 'Address and search bar*'),
    [string[]]$SaveDialogTitles = @('Enregistrer sous*', 'Save As*')
)

# Mode linguistique restreint (AppLocker / WDAC) : Add-Type et UI Automation y sont interdits.
# Vérifié avant tout appel .NET, qui échouerait avec un message moins clair.
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-Output 8
    Write-Error -ErrorAction Continue ('PowerShell est en mode {0} (strategie de securite) : Add-Type et UI Automation sont bloques, le script ne peut pas fonctionner.' -f $ExecutionContext.SessionState.LanguageMode)
    exit 8
}

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

# Supprime les journaux du dossier plus anciens que -LogRetentionDays jours (0 = tout garder) :
# après des mois d'exécutions, ou sur un dossier partagé, il ne grossit pas indéfiniment.
function Remove-OldLogs {
    param([string]$Folder)
    if ($LogRetentionDays -le 0) { return }
    try {
        if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { return }
        $limit = (Get-Date).AddDays(-$LogRetentionDays)
        Get-ChildItem -LiteralPath $Folder -Filter 'EdgeSaveAsCheck_*.log' -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $limit } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
    catch { }
}

function Initialize-Log {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        # Par défaut : à côté du script s'il est lancé depuis un fichier, sinon dossier propre à
        # l'utilisateur (le script peut servir à plusieurs utilisateurs d'un même poste).
        $Path = if ($script:ScriptPath) { Join-Path (Split-Path -Parent $script:ScriptPath) 'Logs' }
        elseif ($env:LOCALAPPDATA) { Join-Path $env:LOCALAPPDATA 'EdgeSaveAsCheck\Logs' }
        else { Join-Path ([System.IO.Path]::GetTempPath()) 'EdgeSaveAsCheck' }
    }
    # Un dossier (existant, ou chemin sans extension) reçoit un fichier par exécution, nommé avec
    # l'utilisateur et l'heure (un dossier partagé entre utilisateurs ne mélange pas les journaux).
    if ((Test-Path -LiteralPath $Path -PathType Container) -or -not [System.IO.Path]::HasExtension($Path)) {
        Remove-OldLogs -Folder $Path
        $user = ([string]$env:USERNAME) -replace '[^\w.-]', '_'
        if (-not $user) { $user = 'user' }
        $Path = Join-Path $Path ('EdgeSaveAsCheck_{0}_{1:yyyyMMdd_HHmmss}.log' -f $user, $script:StartTime)
    }
    # Si le dossier demandé n'est pas accessible en écriture, repli sur le dossier TEMP.
    foreach ($candidate in @($Path, (Join-Path ([System.IO.Path]::GetTempPath()) (Split-Path -Leaf $Path)))) {
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

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
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

        // Default action (click) through MSAA accessibility, without moving the mouse.
        [DllImport("oleacc.dll")] public static extern int AccessibleObjectFromPoint(POINT pt, [MarshalAs(UnmanagedType.IUnknown)] out object acc, [MarshalAs(UnmanagedType.Struct)] out object child);

        [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(POINT pt);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr hWnd, uint flags);
        [DllImport("user32.dll")] private static extern bool ScreenToClient(IntPtr hWnd, ref POINT pt);
        [DllImport("user32.dll")] private static extern bool SetProcessDpiAwarenessContext(IntPtr value);
        [DllImport("user32.dll")] private static extern bool SetProcessDPIAware();
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern uint MapVirtualKey(uint code, uint mapType);
        [DllImport("user32.dll")] private static extern bool SetCursorPos(int x, int y);
        [DllImport("user32.dll")] private static extern bool GetCursorPos(out POINT pt);
        [DllImport("user32.dll")] private static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);

        private const uint GW_OWNER = 4;
        private const uint GA_ROOT = 2;

        // Screen coordinates must be physical pixels to match the target window, so the
        // process is made DPI aware (a no-op if the host already chose a mode).
        public static string EnableDpiAwareness()
        {
            try { if (SetProcessDpiAwarenessContext(new IntPtr(-4))) { return "per-monitor v2"; } }
            catch (EntryPointNotFoundException) { }
            try { if (SetProcessDPIAware()) { return "system"; } }
            catch (EntryPointNotFoundException) { }
            return "unchanged";
        }

        public static uint DpiOf(IntPtr hWnd)
        {
            try { return GetDpiForWindow(hWnd); }
            catch (EntryPointNotFoundException) { return 0; }
        }

        public static string ClassOf(IntPtr hWnd)
        {
            StringBuilder name = new StringBuilder(256);
            GetClassName(hWnd, name, name.Capacity);
            return name.ToString();
        }

        public static IntPtr WindowAt(int x, int y)
        {
            POINT p;
            p.X = x;
            p.Y = y;
            return WindowFromPoint(p);
        }

        public static IntPtr RootOf(IntPtr hWnd)
        {
            return GetAncestor(hWnd, GA_ROOT);
        }

        public static uint ProcessOf(IntPtr hWnd)
        {
            uint processId;
            GetWindowThreadProcessId(hWnd, out processId);
            return processId;
        }

        public static POINT ToClient(IntPtr hWnd, int x, int y)
        {
            POINT p;
            p.X = x;
            p.Y = y;
            ScreenToClient(hWnd, ref p);
            return p;
        }

        public static IntPtr MakeLParam(int x, int y)
        {
            return new IntPtr((y << 16) | (x & 0xFFFF));
        }

        // Native mouse messages posted to the window: they reach Chromium as real input
        // (trusted events), unlike the simulated click of an accessibility action, and
        // they need neither focus nor the foreground window.
        public static void PostClick(IntPtr hWnd, int x, int y)
        {
            IntPtr lp = MakeLParam(x, y);
            PostMessage(hWnd, 0x0200, IntPtr.Zero, lp);          // WM_MOUSEMOVE
            System.Threading.Thread.Sleep(80);
            PostMessage(hWnd, 0x0201, new IntPtr(1), lp);        // WM_LBUTTONDOWN, MK_LBUTTON
            System.Threading.Thread.Sleep(60);
            PostMessage(hWnd, 0x0202, IntPtr.Zero, lp);          // WM_LBUTTONUP
        }

        // Native key press posted to the window (Enter activates the focused button).
        public static void PostKey(IntPtr hWnd, int vk, int ch)
        {
            uint scan = MapVirtualKey((uint)vk, 0);
            uint down = 1u | (scan << 16);
            uint up = down | (1u << 30) | (1u << 31);
            PostMessage(hWnd, 0x0100, new IntPtr(vk), new IntPtr(unchecked((int)down)));   // WM_KEYDOWN
            PostMessage(hWnd, 0x0102, new IntPtr(ch), new IntPtr(unchecked((int)down)));   // WM_CHAR
            System.Threading.Thread.Sleep(50);
            PostMessage(hWnd, 0x0101, new IntPtr(vk), new IntPtr(unchecked((int)up)));     // WM_KEYUP
        }

        // Last resort, opt-in: a real mouse click. It moves the cursor, so it is not background.
        public static void RealClick(int x, int y)
        {
            POINT old;
            GetCursorPos(out old);
            SetCursorPos(x, y);
            System.Threading.Thread.Sleep(80);
            mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);          // MOUSEEVENTF_LEFTDOWN
            System.Threading.Thread.Sleep(60);
            mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);          // MOUSEEVENTF_LEFTUP
            System.Threading.Thread.Sleep(80);
            SetCursorPos(old.X, old.Y);
        }

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

function Get-ElementClass {
    param($Element)
    try { return [string]$Element.Current.ClassName } catch { return '' }
}

# Vrai si la classe UI Automation de l'élément figure dans la liste (égalité exacte, sans
# tenir compte de la casse). Les classes ne sont pas traduites : c'est le critère qui ne
# dépend pas de la langue du navigateur.
function Test-ClassMatch {
    param($Element, [string[]]$Classes)
    if (-not $Classes) { return $false }
    $class = Get-ElementClass $Element
    if (-not $class) { return $false }
    foreach ($candidate in $Classes) {
        if ($class -ieq $candidate) { return $true }
    }
    return $false
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

# Éléments correspondant à une classe UI Automation (critère principal, indépendant de la
# langue) ou, à défaut, à un nom (repli, dépendant de la langue d'Edge). Ordre du résultat :
# classe avant nom ; à critère égal, avec -PreferBrowserUi, interface d'Edge avant contenu web.
function Find-NamedElements {
    param(
        [object[]]$Roots,
        $Condition,
        [string[]]$NamePatterns,
        [string[]]$Classes,
        [switch]$PreferBrowserUi
    )
    $buckets = @(
        (New-Object System.Collections.Generic.List[object]),
        (New-Object System.Collections.Generic.List[object]),
        (New-Object System.Collections.Generic.List[object]),
        (New-Object System.Collections.Generic.List[object])
    )
    foreach ($root in $Roots) {
        if (-not $root) { continue }
        try { $candidates = $root.FindAll($script:Descendants, $Condition) } catch { continue }
        foreach ($element in $candidates) {
            $byClass = Test-ClassMatch $element $Classes
            if (-not $byClass -and -not (Test-LikeAny (Get-ElementName $element) $NamePatterns)) { continue }
            $index = 0
            if (-not $byClass) { $index += 2 }
            if ($PreferBrowserUi -and (Test-InWebContent $element)) { $index += 1 }
            $buckets[$index].Add($element)
        }
    }
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($bucket in $buckets) { $result.AddRange($bucket) }
    return $result.ToArray()
}

# Actionne l'élément. Selon la version d'Edge, le même bouton n'expose pas les mêmes
# actions, et Expand peut renvoyer E_FAIL (notamment sur un bouton désactivé) ; on essaie
# donc chaque action disponible dans son propre try/catch, jusqu'à la première qui aboutit.
# Tout reste en UI Automation managé (pas de type COM comme LegacyIAccessiblePattern, absent
# de l'assembly .NET utilisé par Windows PowerShell 5.1).
function Invoke-UiaElement {
    param($Element)
    $attempts = @(
        @{ Name = 'Invoke'; Pattern = [System.Windows.Automation.InvokePattern]::Pattern; Action = { param($p) $p.Invoke() } },
        @{ Name = 'Toggle'; Pattern = [System.Windows.Automation.TogglePattern]::Pattern; Action = { param($p) $p.Toggle() } },
        @{ Name = 'Expand'; Pattern = [System.Windows.Automation.ExpandCollapsePattern]::Pattern; Action = { param($p) $p.Expand() } }
    )
    $lastError = $null
    foreach ($attempt in $attempts) {
        $pattern = $null
        if (-not $Element.TryGetCurrentPattern($attempt.Pattern, [ref]$pattern)) { continue }
        try { & $attempt.Action $pattern; return $attempt.Name }
        catch {
            $lastError = $_
            # Expand renvoie parfois E_FAIL tout en ayant ouvert le panneau : on vérifie l'état.
            if ($attempt.Name -eq 'Expand') {
                try {
                    if ($pattern.Current.ExpandCollapseState -eq [System.Windows.Automation.ExpandCollapseState]::Expanded) { return 'Expand' }
                }
                catch { }
            }
            Write-Log ("Action '{0}' en echec, essai suivant : {1}" -f $attempt.Name, $_.Exception.Message) 'WARN'
        }
    }
    # Dernier recours : action par défaut via l'accessibilité (MSAA) au centre de l'élément.
    # Certaines versions d'Edge n'exposent qu'Expand sur le bouton Telechargements, et Expand
    # renvoie E_FAIL ; accDoDefaultAction clique le bouton sans bouger la souris.
    try { return (Invoke-DefaultActionAtPoint $Element) }
    catch {
        $detail = if ($lastError) { $lastError.Exception.Message } else { 'aucun pattern utilisable' }
        throw ("Aucune action n'a abouti sur '{0}' : {1} / clic accessibilite : {2}" -f (Get-ElementName $Element), $detail, $_.Exception.Message)
    }
}

# Clique l'élément via son action par défaut MSAA, sans déplacer le curseur. On vise le point
# cliquable calculé par UI Automation (ou, à défaut, le centre du rectangle de l'élément).
function Invoke-DefaultActionAtPoint {
    param($Element)
    $point = Get-ElementScreenPoint $Element
    $pt = New-Object EdgeSaveAsCheck.POINT
    $pt.X = $point.X; $pt.Y = $point.Y
    $acc = $null; $child = $null
    $hr = [EdgeSaveAsCheck.Win32]::AccessibleObjectFromPoint($pt, [ref]$acc, [ref]$child)
    if ($hr -ne 0 -or $null -eq $acc) { throw ("AccessibleObjectFromPoint a echoue (HRESULT 0x{0:X8}) au point {1},{2}" -f $hr, $point.X, $point.Y) }
    $acc.accDoDefaultAction($child)
    return 'accDoDefaultAction'
}

# Point d'écran (pixels physiques, le processus étant DPI aware) où cliquer l'élément :
# point cliquable d'UI Automation, ou à défaut centre de son rectangle.
function Get-ElementScreenPoint {
    param($Element)
    try {
        $point = $Element.GetClickablePoint()
        return [pscustomobject]@{ X = [int]$point.X; Y = [int]$point.Y }
    }
    catch {
        $rect = $Element.Current.BoundingRectangle
        if ([double]::IsInfinity($rect.Width) -or $rect.Width -le 0 -or $rect.Height -le 0) {
            throw 'element sans point cliquable (hors ecran ou masque)'
        }
        return [pscustomobject]@{ X = [int]($rect.X + $rect.Width / 2); Y = [int]($rect.Y + $rect.Height / 2) }
    }
}

# Fenêtre native qui porte l'élément : l'ancêtre UI Automation le plus proche (lui compris)
# doté d'un handle de fenêtre. Ne dépend pas de l'ordre d'empilement des fenêtres.
function Get-ElementWindow {
    param($Element)
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $node = $Element
    for ($depth = 0; $depth -lt 40 -and $node; $depth++) {
        try {
            $handle = [IntPtr][int]$node.Current.NativeWindowHandle
            if ($handle -ne [IntPtr]::Zero) { return $handle }
            $node = $walker.GetParent($node)
        }
        catch { return [IntPtr]::Zero }
    }
    return [IntPtr]::Zero
}

# Détermine à quelle fenêtre poster un message d'entrée pour l'élément. On vise la fenêtre
# visible sous le point (celle qui a fonctionné lors des essais) à condition qu'elle appartienne
# à Edge. Si une autre application recouvre Edge à cet endroit, cette fenêtre-là n'est pas la
# bonne : on prend celle de l'arbre UI Automation de l'élément, indépendante de l'ordre des
# fenêtres. Aucun message n'est jamais posté à une fenêtre qui n'est pas celle d'Edge.
function Resolve-ClickWindows {
    param($Element)
    $point = Get-ElementScreenPoint $Element
    $edgePids = @(Get-EdgeProcessIds)
    $near = [EdgeSaveAsCheck.Win32]::WindowAt($point.X, $point.Y)
    $source = 'fenetre sous le point'
    if ($near -ne [IntPtr]::Zero) {
        $owner = [int][EdgeSaveAsCheck.Win32]::ProcessOf($near)
        if ($edgePids -notcontains $owner) {
            $other = Get-Process -Id $owner -ErrorAction SilentlyContinue
            Write-Log ("Edge est recouvert par '{0}' a cet endroit : cible prise dans l'arbre UI Automation (aucun message envoye a cette application)" -f $other.ProcessName) 'WARN'
            $near = [IntPtr]::Zero
        }
    }
    if ($near -eq [IntPtr]::Zero) {
        $near = Get-ElementWindow $Element
        $source = 'arbre UI Automation'
        if ($near -ne [IntPtr]::Zero -and ($edgePids -notcontains [int][EdgeSaveAsCheck.Win32]::ProcessOf($near))) { $near = [IntPtr]::Zero }
    }
    if ($near -eq [IntPtr]::Zero) { throw ('aucune fenetre d Edge atteignable pour le point {0},{1}' -f $point.X, $point.Y) }
    $top = [EdgeSaveAsCheck.Win32]::RootOf($near)
    if ($top -eq [IntPtr]::Zero) { $top = $near }
    return [pscustomobject]@{ Point = $point; Near = $near; Top = $top; Source = $source }
}

# Clic « natif » posté à la fenêtre qui porte l'élément. -Target Root : fenêtre de premier
# niveau (le panneau Téléchargements est une fenêtre contextuelle propre) ; Child : fenêtre
# enfant qui affiche la page web. Renvoie $null si Child est identique à Root (déjà essayé).
function Send-NativeClick {
    param($Element, [ValidateSet('Root', 'Child')][string]$Target)
    $w = Resolve-ClickWindows $Element
    if ($Target -eq 'Child' -and $w.Near -eq $w.Top) { return $null }
    $hwnd = if ($Target -eq 'Root') { $w.Top } else { $w.Near }
    $client = [EdgeSaveAsCheck.Win32]::ToClient($hwnd, $w.Point.X, $w.Point.Y)
    Write-Log ("Clic natif ({0}) : ecran {1},{2} -> fenetre 0x{3:X} classe '{4}' DPI {5}, client {6},{7} (cible : {8})" -f $Target, $w.Point.X, $w.Point.Y, $hwnd.ToInt64(), [EdgeSaveAsCheck.Win32]::ClassOf($hwnd), [EdgeSaveAsCheck.Win32]::DpiOf($hwnd), $client.X, $client.Y, $w.Source)
    [EdgeSaveAsCheck.Win32]::PostClick($hwnd, $client.X, $client.Y)
    return ('message souris poste a la fenetre ' + $Target.ToLower())
}

# Touche Entrée postée à la fenêtre du panneau. Elle active le bouton qui a le focus : on ne
# l'envoie donc que si c'est bien « Enregistrer sous » (sinon elle pourrait ouvrir le fichier).
function Send-NativeEnter {
    param($Element)
    $focused = $false
    try { $focused = [bool]$Element.Current.HasKeyboardFocus } catch { }
    if (-not $focused) {
        Write-Log "Touche Entree non envoyee : le bouton n'a pas le focus clavier (elle activerait un autre bouton)"
        return $null
    }
    $w = Resolve-ClickWindows $Element
    [EdgeSaveAsCheck.Win32]::PostKey($w.Top, 0x0D, 13)
    return 'touche Entree postee a la fenetre'
}

# Un essai de clic, par ordre de discrétion croissante ; renvoie une description de ce qui
# a été fait, ou $null si la méthode ne s'applique pas.
function Invoke-SaveAsClick {
    param([string]$Method, $Element)
    switch ($Method) {
        'Invoke' { return (Invoke-UiaElement $Element) }
        'accDoDefaultAction' { return (Invoke-DefaultActionAtPoint $Element) }
        'MouseRoot' { return (Send-NativeClick -Element $Element -Target 'Root') }
        'MouseChild' { return (Send-NativeClick -Element $Element -Target 'Child') }
        'Keyboard' { return (Send-NativeEnter $Element) }
        'RealMouse' {
            $point = Get-ElementScreenPoint $Element
            [EdgeSaveAsCheck.Win32]::RealClick($point.X, $point.Y)
            return 'vrai clic souris (curseur deplace)'
        }
        default { throw ("methode inconnue : {0}" -f $Method) }
    }
}

# Propriétés de l'élément et de ses parents : montre s'il s'agit d'un bouton de page web
# (panneau WebUI) et quelles actions UI Automation il propose.
function Write-ElementDiagnostic {
    param($Element, [string]$Label)
    try {
        $info = $Element.Current
        $rect = $info.BoundingRectangle
        $patterns = @($Element.GetSupportedPatterns() | ForEach-Object { $_.ProgrammaticName -replace 'PatternIdentifiers\.Pattern$', '' })
        $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
        $chain = New-Object System.Collections.Generic.List[string]
        $node = $Element
        for ($level = 0; $level -lt 6; $level++) {
            $node = $walker.GetParent($node)
            if (-not $node) { break }
            $type = $node.Current.ControlType.ProgrammaticName -replace '^ControlType\.', ''
            $chain.Add(('{0}:{1}' -f $type, (Get-ElementName $node)))
        }
        Write-Log ("{0} : nom='{1}' classe='{2}' framework='{3}' id='{4}' actif={5} horsecran={6} focus={7} rect={8:N0},{9:N0} {10:N0}x{11:N0} page-web={12} patterns=[{13}] parents=[{14}]" -f $Label, $info.Name, $info.ClassName, $info.FrameworkId, $info.AutomationId, $info.IsEnabled, $info.IsOffscreen, $info.HasKeyboardFocus, $rect.X, $rect.Y, $rect.Width, $rect.Height, (Test-InWebContent $Element), ($patterns -join ','), ($chain -join ' < '))
    }
    catch { Write-Log ("{0} : diagnostic impossible ({1})" -f $Label, $_.Exception.Message) 'WARN' }
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
    # Par ordre de fiabilité : classe de la barre d'adresse (indépendante de la langue), puis
    # nom (français ou anglais), puis 1er champ de saisie de l'interface d'Edge (hors page web).
    $byName = $null
    $fallback = $null
    foreach ($edit in $edits) {
        if (Test-ClassMatch $edit $AddressBarClasses) { return Get-ElementValue $edit }
        if (-not $byName -and (Test-LikeAny (Get-ElementName $edit) $AddressBarNames)) { $byName = $edit }
        if (-not $fallback -and -not (Test-InWebContent $edit)) { $fallback = $edit }
    }
    if ($byName) { return Get-ElementValue $byName }
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
        $saveAsButtons = @($node.FindAll($script:Descendants, $script:CondClickable) | Where-Object {
                (Test-ClassMatch $_ $SaveAsButtonClasses) -or (Test-LikeAny (Get-ElementName $_) $SaveAsButtonNames)
            })
        if ($saveAsButtons.Count -gt 1) { return $null }
        $match = $names | Where-Object { $_ -like "*$FileNamePattern*" } | Select-Object -First 1
        if ($match) { return $match }
    }
    return $null
}

function Find-SaveAsButton {
    param($Page)
    $buttons = @(Find-NamedElements -Roots (Get-PanelSearchRoots $Page) -Condition $script:CondClickable -NamePatterns $SaveAsButtonNames -Classes $SaveAsButtonClasses)
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
    $buttons = @(Find-NamedElements -Roots @($Page.Window) -Condition $script:CondClickable -NamePatterns $DownloadsButtonNames -Classes $DownloadsButtonClasses -PreferBrowserUi)
    if ($buttons.Count -eq 0) { return $false }
    $identifiedBy = if (Test-ClassMatch $buttons[0] $DownloadsButtonClasses) { 'sa classe (independant de la langue)' } else { 'son nom (depend de la langue)' }
    Write-Log ('Bouton Telechargements identifie par {0}' -f $identifiedBy)
    Write-ElementDiagnostic -Element $buttons[0] -Label 'Bouton Telechargements'
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
            try { $root.FindAll($script:Descendants, $entry.Condition) | ForEach-Object { '{0}[{1}]' -f (Get-ElementName $_), (Get-ElementClass $_) } } catch { }
        }
        $names = @($names | Where-Object { $_ -and $_ -ne '[]' } | Select-Object -Unique | Select-Object -First 100)
        Write-Log ('Diagnostic - {0} vus par UI Automation (nom[classe]) : {1}' -f $entry.Label, ($names -join ' | ')) 'WARN'
    }
}

# =============================================================================
# Étape 4 : fenêtre Windows « Enregistrer sous »
# =============================================================================

# Instantané des fenêtres de premier niveau (clé = handle), pris avant le clic.
function Get-WindowSnapshot {
    $snapshot = @{}
    foreach ($window in [EdgeSaveAsCheck.Win32]::GetTopLevelWindows()) { $snapshot[[int64]$window.Handle] = $true }
    return $snapshot
}

# Attend une NOUVELLE fenêtre (absente de l'instantané) qui est la boîte « Enregistrer sous ».
# Reconnue sans dépendre de la langue : boîte de dialogue Windows (classe #32770) issue d'Edge.
# Le titre (-SaveDialogTitles) sert d'appoint : avec lui, une fenêtre d'une autre classe issue
# d'Edge, ou une boîte #32770 hors du processus Edge (WARN), est aussi acceptée. Toute autre
# nouvelle fenêtre visible est journalisée (classe, titre, processus).
function Wait-SaveDialog {
    param([IntPtr]$OwnerHandle, [hashtable]$Known, [hashtable]$Reported, [int]$Seconds)
    $deadline = (Get-Date).AddSeconds($Seconds)
    do {
        $edgePids = @(Get-EdgeProcessIds)
        foreach ($window in [EdgeSaveAsCheck.Win32]::GetTopLevelWindows()) {
            if (-not $window.Visible -or $Known.ContainsKey([int64]$window.Handle)) { continue }
            $fromEdge = ($edgePids -contains [int]$window.ProcessId) -or ($window.Owner -eq $OwnerHandle)
            $isDialog = ($window.ClassName -eq '#32770')
            $titleMatches = Test-LikeAny $window.Title $SaveDialogTitles
            if (($fromEdge -and ($isDialog -or $titleMatches)) -or ($isDialog -and $titleMatches)) {
                if (-not $fromEdge) { Write-Log ("Fenetre trouvee hors processus Edge (PID {0}) : verifier qu'elle vient bien du clic" -f $window.ProcessId) 'WARN' }
                return $window
            }
            $key = [int64]$window.Handle
            if (-not $Reported.ContainsKey($key) -and $Reported.Count -lt 15) {
                $Reported[$key] = $true
                $process = Get-Process -Id ([int]$window.ProcessId) -ErrorAction SilentlyContinue
                Write-Log ("Nouvelle fenetre ignoree : classe '{0}' titre '{1}' processus '{2}' (Edge : {3})" -f $window.ClassName, $window.Title, $process.ProcessName, $fromEdge) 'WARN'
            }
        }
        Start-Sleep -Milliseconds 250
    } while ((Get-Date) -lt $deadline)
    return $null
}

# Le panneau Téléchargements d'Edge est une page web (WebUI) : l'action UI Automation
# « Invoke » peut donner le focus au bouton sans déclencher son action. On essaie donc
# plusieurs méthodes, en s'arrêtant dès qu'une fenêtre « Enregistrer sous » apparaît. Aucune
# ne déplace la souris, sauf RealMouse (désactivée par défaut, voir -AllowRealMouse).
# Renvoie la fenêtre trouvée (ou $null) et les méthodes essayées.
function Invoke-SaveAsClickLadder {
    param($Page, $SaveAs)
    $known = Get-WindowSnapshot
    $reported = @{}
    Write-ElementDiagnostic -Element $SaveAs.Element -Label 'Bouton Enregistrer sous'
    # Ordre : le clic natif sur la fenêtre du panneau a fonctionné sur Edge 26100 (les actions
    # UI Automation ne donnent que le focus), il passe donc en premier ; les autres suivent.
    $methods = @('MouseRoot', 'Invoke', 'accDoDefaultAction', 'MouseChild', 'Keyboard')
    if ($AllowRealMouse) { $methods += 'RealMouse' }
    $dialog = $null
    $tried = New-Object System.Collections.Generic.List[string]
    foreach ($method in $methods) {
        try { $done = Invoke-SaveAsClick -Method $method -Element $SaveAs.Element }
        catch {
            Write-Log ("Methode '{0}' en echec : {1}" -f $method, $_.Exception.Message) 'WARN'
            continue
        }
        if (-not $done) {
            Write-Log ("Methode '{0}' non applicable ici" -f $method)
            continue
        }
        $tried.Add($method)
        Write-Log ("Clic sur '{0}' (methode {1} : {2})" -f (Get-ElementName $SaveAs.Element), $method, $done)
        $dialog = Wait-SaveDialog -OwnerHandle $Page.Handle -Known $known -Reported $reported -Seconds $ClickWaitSeconds
        if ($dialog) {
            Write-Log ("Fenetre obtenue avec la methode '{0}'" -f $method) 'OK'
            break
        }
        # Bouton disparu : le clic a eu un effet (fenêtre lente ?). Ne pas cliquer une 2e fois.
        $again = Find-SaveAsButton -Page $Page
        if (-not $again) {
            Write-Log ("Le bouton Enregistrer sous a disparu apres la methode '{0}' : attente de la fenetre sans autre clic" -f $method) 'WARN'
            break
        }
        $SaveAs = $again
        Write-Log ("Aucune fenetre apres la methode '{0}' (bouton toujours present)" -f $method) 'WARN'
    }
    if (-not $dialog) {
        $dialog = Wait-SaveDialog -OwnerHandle $Page.Handle -Known $known -Reported $reported -Seconds $DialogTimeoutSeconds
    }
    return [pscustomobject]@{ Dialog = $dialog; Tried = @($tried) }
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
    Initialize-Win32
    # Avant tout appel UI Automation : coordonnées en pixels physiques, cohérentes avec celles
    # des fenêtres d'Edge (sinon un écran avec mise à l'échelle décale les clics postés).
    # Étape facultative : un échec ici ne doit jamais interrompre la vérification.
    $dpiMode = 'indisponible'
    try { $dpiMode = [EdgeSaveAsCheck.Win32]::EnableDpiAwareness() }
    catch { $dpiMode = 'indisponible : ' + $_.Exception.Message }
    Write-Log ('Mode DPI du processus : {0}' -f $dpiMode)
    Initialize-Uia

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
        Stop-Check 5 ("Bouton Enregistrer sous (classe '{0}' ou nom '{1}') introuvable dans le panneau Telechargements{2}." -f ($SaveAsButtonClasses -join "' / '"), ($SaveAsButtonNames -join "' / '"), $detail)
    }
    if ($saveAs.PromptText) { Write-Log ("Invite de telechargement : '{0}'" -f $saveAs.PromptText) }

    # --- 3. Clic « Enregistrer sous » + 4. Vérification de la fenêtre Windows ------
    $result = Invoke-SaveAsClickLadder -Page $page -SaveAs $saveAs
    $dialog = $result.Dialog
    if (-not $dialog) {
        Stop-Check 6 ("La fenetre Windows d'enregistrement ne s'est pas ouverte (methodes essayees : {0}). Voir les lignes du journal sur le bouton et les nouvelles fenetres." -f ($result.Tried -join ', '))
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

    # Valeurs venant de PAD : espaces ou retours à la ligne parasites, variable vide.
    $UrlPattern = $UrlPattern.Trim()
    $TitlePattern = $TitlePattern.Trim()
    $FileNamePattern = $FileNamePattern.Trim()
    if (-not $FileNamePattern) { $FileNamePattern = '*' }
    Write-Log ("Parametres - URL '{0}' | Titre '{1}' | Fichier '{2}' | CloseDialog {3}" -f $UrlPattern, $TitlePattern, $FileNamePattern, [bool]$CloseDialog)
    if (-not $UrlPattern -or -not $TitlePattern) {
        Stop-Check 9 'UrlPattern et TitlePattern ne doivent pas etre vides (variable PAD vide ?).'
    }

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
} @parameters
