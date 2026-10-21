# =============================================================================
# Edge-SaveAs-Check : script autonome pour Power Automate Desktop. C'est le SEUL fichier
# du projet : on le modifie directement ici (il ne contient plus de commentaires de detail,
# voir le principe de fonctionnement ci-dessous et l'historique git pour la version commentee).
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
#  - Regles pour modifier ce fichier : UTF-8 sans BOM, caracteres ASCII seulement (PAD le
#    copie dans un fichier temporaire), jamais de signe pourcent hors des deux variables PAD.
#
# Principe de fonctionnement
#  1. Trouve l'onglet Edge (URL + titre) et ouvre le panneau Telechargements.
#  2. Reconnait les elements par leur classe UI Automation (ToolbarDownloadsButton, saveAs),
#     non traduite : independant de la langue d'Edge. Les noms FR/EN ne servent que de repli.
#  3. Clique " Enregistrer sous " par une echelle de methodes : message souris natif poste a
#     la fenetre du panneau (celui qui fonctionne sur Edge), puis Invoke, accDoDefaultAction,
#     message natif a nouveau, fenetre enfant, touche Entree postee. Jamais de message a une
#     fenetre qui n'est pas Edge ; pas de curseur ni de clavier reels sauf AllowRealMouse.
#  4. Attend une NOUVELLE boite Windows (classe #32770) issue d'Edge, sans lire son titre.
#  5. Ecrit toujours un journal se terminant par RESULTAT, et le code sur la sortie standard.
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
[CmdletBinding()]
param(
    [string]$UrlPattern = 'https://mon-application.example/*',
    [string]$TitlePattern = '*Titre de ma page*',
    [string]$FileNamePattern = '*',
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
$script:PageHandle = [IntPtr]::Zero
$script:DialogLeftOpen = $false
$script:AddressBarHow = $null
$script:QuietConsole = ($Host.Name -eq 'ConsoleHost') -and [Console]::IsOutputRedirected


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
        $Path = if ($script:ScriptPath) { Join-Path (Split-Path -Parent $script:ScriptPath) 'Logs' }
        elseif ($env:LOCALAPPDATA) { Join-Path $env:LOCALAPPDATA 'EdgeSaveAsCheck\Logs' }
        else { Join-Path ([System.IO.Path]::GetTempPath()) 'EdgeSaveAsCheck' }
    }
    if ((Test-Path -LiteralPath $Path -PathType Container) -or -not [System.IO.Path]::HasExtension($Path)) {
        Remove-OldLogs -Folder $Path
        $user = ([string]$env:USERNAME) -replace '[^\w.-]', '_'
        if (-not $user) { $user = 'user' }
        $stamp = $script:StartTime.ToString('yyyyMMdd_HHmmss', [System.Globalization.CultureInfo]::InvariantCulture)
        $Path = Join-Path $Path ('EdgeSaveAsCheck_{0}_{1}_{2}.log' -f $user, $stamp, $PID)
    }
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

function Remove-InvisibleMarks {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [regex]::Replace($Text, '[\u200B-\u200F\u202A-\u202E\u2066-\u2069\uFEFF]', '')
}

function Test-PatternMatch {
    param([string]$Value, [string]$Pattern)
    if ($Value -like $Pattern) { return $true }
    $literal = $Pattern -replace '\[', '`[' -replace '\]', '`]'
    if ($literal -ne $Pattern -and $Value -like $literal) { return $true }
    return $false
}


function ConvertTo-PsLiteral {
    param([string]$Value)
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
    if ($script:LogFile) { $arguments.Add((ConvertTo-PsArgument -Name 'LogPath' -Value $script:LogFile)) }

    $command = '& {0} {1}; exit $LASTEXITCODE' -f (ConvertTo-PsLiteral $script:ScriptPath), ($arguments -join ' ')
    $encoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($command))
    $exe = (Get-Process -Id $PID).Path
    $process = Start-Process -FilePath $exe -WindowStyle Hidden -PassThru -ArgumentList @(
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-EncodedCommand', $encoded)
    Write-Log ('Relance en arriere-plan dans un processus masque (PID {0}). Resultat dans : {1}' -f $process.Id, $script:LogFile) 'OK'
}


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

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
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

        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
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
        // Returns false when Windows refused a message (for example Edge running elevated while
        // this process is not: the input of a lower-integrity process is dropped).
        public static bool PostClick(IntPtr hWnd, int x, int y)
        {
            IntPtr lp = MakeLParam(x, y);
            bool ok = PostMessage(hWnd, 0x0200, IntPtr.Zero, lp);          // WM_MOUSEMOVE
            System.Threading.Thread.Sleep(80);
            ok = PostMessage(hWnd, 0x0201, new IntPtr(1), lp) && ok;       // WM_LBUTTONDOWN, MK_LBUTTON
            System.Threading.Thread.Sleep(60);
            ok = PostMessage(hWnd, 0x0202, IntPtr.Zero, lp) && ok;         // WM_LBUTTONUP
            return ok;
        }

        // Native key press posted to the window (Enter activates the focused button).
        public static bool PostKey(IntPtr hWnd, int vk, int ch)
        {
            uint scan = MapVirtualKey((uint)vk, 0);
            uint down = 1u | (scan << 16);
            uint up = down | (1u << 30) | (1u << 31);
            bool ok = PostMessage(hWnd, 0x0100, new IntPtr(vk), new IntPtr(unchecked((int)down)));   // WM_KEYDOWN
            ok = PostMessage(hWnd, 0x0102, new IntPtr(ch), new IntPtr(unchecked((int)down))) && ok;  // WM_CHAR
            System.Threading.Thread.Sleep(50);
            ok = PostMessage(hWnd, 0x0101, new IntPtr(vk), new IntPtr(unchecked((int)up))) && ok;    // WM_KEYUP
            return ok;
        }

        public static IntPtr ForegroundWindow()
        {
            return GetForegroundWindow();
        }

        public static IntPtr OwnerOf(IntPtr hWnd)
        {
            return GetWindow(hWnd, GW_OWNER);
        }

        // True when the screen point lies inside the window rectangle.
        public static bool WindowContains(IntPtr hWnd, int x, int y)
        {
            RECT r;
            if (!GetWindowRect(hWnd, out r)) { return false; }
            return x >= r.Left && x < r.Right && y >= r.Top && y < r.Bottom;
        }

        public static string CursorText()
        {
            POINT p;
            if (!GetCursorPos(out p)) { return "?"; }
            return p.X.ToString() + "," + p.Y.ToString();
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
    $sessionId = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    return @(Get-Process -Name msedge -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $sessionId } | ForEach-Object { $_.Id })
}

function Get-EdgeBrowserWindows {
    $edgePids = Get-EdgeProcessIds
    foreach ($window in [EdgeSaveAsCheck.Win32]::GetTopLevelWindows()) {
        if ($window.ClassName -notlike 'Chrome_WidgetWin*' -or -not $window.Visible) { continue }
        if ($window.Owner -ne [IntPtr]::Zero -or [string]::IsNullOrEmpty($window.Title)) { continue }
        if ($edgePids -notcontains [int]$window.ProcessId) { continue }
        try { [System.Windows.Automation.AutomationElement]::FromHandle($window.Handle) } catch { }
    }
}


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

function Test-ClassMatch {
    param($Element, [string[]]$Classes)
    if (-not $Classes) { return $false }
    $class = Get-ElementClass $Element
    if (-not $class) { return $false }
    $tokens = @($class -split '\s+' | Where-Object { $_ })
    foreach ($candidate in $Classes) {
        foreach ($token in $tokens) {
            if ($token -ieq $candidate) { return $true }
        }
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

function Find-NamedElements {
    param(
        [object[]]$Roots,
        $Condition,
        [string[]]$NamePatterns,
        [string[]]$Classes,
        [switch]$PreferBrowserUi,
        [switch]$ExcludeWebContent
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
            if ($ExcludeWebContent -and (Test-InWebContent $element)) { continue }
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
            if ($attempt.Name -eq 'Expand') {
                try {
                    if ($pattern.Current.ExpandCollapseState -eq [System.Windows.Automation.ExpandCollapseState]::Expanded) { return 'Expand' }
                }
                catch { }
            }
            Write-Log ("Action '{0}' en echec, essai suivant : {1}" -f $attempt.Name, $_.Exception.Message) 'WARN'
        }
    }
    try { return (Invoke-DefaultActionAtPoint $Element) }
    catch {
        $detail = if ($lastError) { $lastError.Exception.Message } else { 'aucun pattern utilisable' }
        throw ("Aucune action n'a abouti sur '{0}' : {1} / clic accessibilite : {2}" -f (Get-ElementName $Element), $detail, $_.Exception.Message)
    }
}

function Assert-PointOnEdge {
    param($Point)
    $near = [EdgeSaveAsCheck.Win32]::WindowAt($Point.X, $Point.Y)
    if ($near -eq [IntPtr]::Zero) { return }
    $owner = [int][EdgeSaveAsCheck.Win32]::ProcessOf($near)
    if (@(Get-EdgeProcessIds) -notcontains $owner) {
        $other = Get-Process -Id $owner -ErrorAction SilentlyContinue
        throw ("le point {0},{1} est recouvert par '{2}' : aucune action envoyee" -f $Point.X, $Point.Y, $other.ProcessName)
    }
}

function Invoke-DefaultActionAtPoint {
    param($Element)
    $point = Get-ElementScreenPoint $Element
    Assert-PointOnEdge $point
    $pt = New-Object EdgeSaveAsCheck.POINT
    $pt.X = $point.X; $pt.Y = $point.Y
    $acc = $null; $child = $null
    $hr = [EdgeSaveAsCheck.Win32]::AccessibleObjectFromPoint($pt, [ref]$acc, [ref]$child)
    if ($hr -ne 0 -or $null -eq $acc) { throw ("AccessibleObjectFromPoint a echoue (HRESULT 0x{0:X8}) au point {1},{2}" -f $hr, $point.X, $point.Y) }
    $acc.accDoDefaultAction($child)
    return 'accDoDefaultAction'
}

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

function Find-PanelWindowAt {
    param($Point)
    if ($script:PageHandle -eq [IntPtr]::Zero) { return [IntPtr]::Zero }
    foreach ($window in [EdgeSaveAsCheck.Win32]::GetTopLevelWindows()) {
        if (-not $window.Visible -or $window.Owner -ne $script:PageHandle -or $window.ClassName -notlike 'Chrome_WidgetWin*') { continue }
        if ([EdgeSaveAsCheck.Win32]::WindowContains($window.Handle, $Point.X, $Point.Y)) { return $window.Handle }
    }
    return [IntPtr]::Zero
}

function Get-WindowVerification {
    param([IntPtr]$Top, $Element)
    $checks = New-Object System.Collections.Generic.List[string]
    if ($script:PageHandle -ne [IntPtr]::Zero) {
        if ($Top -eq $script:PageHandle) { $checks.Add('fenetre de la page') }
        elseif ([EdgeSaveAsCheck.Win32]::OwnerOf($Top) -eq $script:PageHandle) { $checks.Add('popup possede par la page') }
    }
    $uia = Get-ElementWindow $Element
    if ($uia -ne [IntPtr]::Zero) {
        $uiaTop = [EdgeSaveAsCheck.Win32]::RootOf($uia)
        if ($uiaTop -eq [IntPtr]::Zero) { $uiaTop = $uia }
        if ($uiaTop -eq $Top) { $checks.Add('meme fenetre que l arbre UI Automation') }
    }
    if ($checks.Count -eq 0) { return 'NON VERIFIEE (ni fenetre de la page, ni popup possede par elle, ni arbre UI Automation)' }
    return ($checks -join ' + ')
}

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
            Write-Log ("Edge est recouvert par '{0}' a cet endroit : cible cherchee sans dependre de l'ordre des fenetres (aucun message envoye a cette application)" -f $other.ProcessName) 'WARN'
            $near = [IntPtr]::Zero
        }
    }
    if ($near -eq [IntPtr]::Zero) {
        $near = Find-PanelWindowAt $point
        $source = 'popup d Edge possede par la page'
        if ($near -eq [IntPtr]::Zero) {
            $near = Get-ElementWindow $Element
            $source = 'arbre UI Automation'
            if ($near -ne [IntPtr]::Zero) {
                if ($edgePids -notcontains [int][EdgeSaveAsCheck.Win32]::ProcessOf($near)) { $near = [IntPtr]::Zero }
                elseif ($script:PageHandle -ne [IntPtr]::Zero -and [EdgeSaveAsCheck.Win32]::RootOf($near) -eq $script:PageHandle) {
                    throw "la fenetre trouvee par l'arbre UI Automation est celle de la page elle-meme, pas celle du panneau : aucun message envoye pour ne pas cliquer la page"
                }
            }
        }
    }
    if ($near -eq [IntPtr]::Zero) { throw ('aucune fenetre d Edge atteignable pour le point {0},{1}' -f $point.X, $point.Y) }
    $top = [EdgeSaveAsCheck.Win32]::RootOf($near)
    if ($top -eq [IntPtr]::Zero) { $top = $near }
    return [pscustomobject]@{ Point = $point; Near = $near; Top = $top; Source = $source; Verification = (Get-WindowVerification -Top $top -Element $Element) }
}

function Send-NativeClick {
    param($Element, [ValidateSet('Root', 'Child')][string]$Target)
    $w = Resolve-ClickWindows $Element
    if ($Target -eq 'Child' -and $w.Near -eq $w.Top) { return $null }
    $hwnd = if ($Target -eq 'Root') { $w.Top } else { $w.Near }
    $client = [EdgeSaveAsCheck.Win32]::ToClient($hwnd, $w.Point.X, $w.Point.Y)
    Write-Log ("Clic natif ({0}) : ecran {1},{2} -> fenetre 0x{3:X} classe '{4}' proprietaire 0x{5:X} DPI {6}, client {7},{8} (cible : {9} ; verification : {10})" -f $Target, $w.Point.X, $w.Point.Y, $hwnd.ToInt64(), [EdgeSaveAsCheck.Win32]::ClassOf($hwnd), [EdgeSaveAsCheck.Win32]::OwnerOf($hwnd).ToInt64(), [EdgeSaveAsCheck.Win32]::DpiOf($hwnd), $client.X, $client.Y, $w.Source, $w.Verification)
    if (-not [EdgeSaveAsCheck.Win32]::PostClick($hwnd, $client.X, $client.Y)) {
        throw "Windows a refuse le message souris (Edge lance avec des droits eleves alors que ce script ne l'est pas ?)"
    }
    return ('message souris poste a la fenetre ' + $Target.ToLower())
}

function Send-NativeEnter {
    param($Element)
    $focused = $false
    try { $focused = [bool]$Element.Current.HasKeyboardFocus } catch { }
    if (-not $focused) {
        Write-Log "Touche Entree non envoyee : le bouton n'a pas le focus clavier (elle activerait un autre bouton)"
        return $null
    }
    $w = Resolve-ClickWindows $Element
    if (-not [EdgeSaveAsCheck.Win32]::PostKey($w.Top, 0x0D, 13)) {
        throw "Windows a refuse la touche Entree (Edge lance avec des droits eleves alors que ce script ne l'est pas ?)"
    }
    return 'touche Entree postee a la fenetre'
}

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
            Assert-PointOnEdge $point
            [EdgeSaveAsCheck.Win32]::RealClick($point.X, $point.Y)
            return 'vrai clic souris (curseur deplace)'
        }
        default { throw ("methode inconnue : {0}" -f $Method) }
    }
}

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


function Test-UrlMatch {
    param([string]$Url)
    $Url = Remove-InvisibleMarks $Url
    if ([string]::IsNullOrWhiteSpace($Url)) { return $false }
    if (Test-PatternMatch $Url $UrlPattern) { return $true }
    $prefix = '^[a-zA-Z][a-zA-Z0-9+.-]*://(www\.)?'
    return (Test-PatternMatch ($Url -replace $prefix, '') ($UrlPattern -replace $prefix, ''))
}

function Get-AddressBarValue {
    param($Window)
    try { $edits = $Window.FindAll($script:Descendants, $script:CondEdit) } catch { return $null }
    $byClass = $null
    $byName = $null
    $fallback = $null
    foreach ($edit in $edits) {
        if (Test-ClassMatch $edit $AddressBarClasses) { $byClass = $edit; break }
        if (-not $byName -and (Test-LikeAny (Get-ElementName $edit) $AddressBarNames)) { $byName = $edit }
        if (-not $fallback -and -not (Test-InWebContent $edit)) { $fallback = $edit }
    }
    $chosen = $null
    $how = ''
    if ($byClass) { $chosen = $byClass; $how = 'classe' }
    elseif ($byName) { $chosen = $byName; $how = 'nom (repli)' }
    elseif ($fallback) { $chosen = $fallback; $how = 'premier champ de saisie hors page web (repli)' }
    if ($chosen) {
        if (-not $script:AddressBarHow) {
            $script:AddressBarHow = $how
            Write-Log ('Barre d adresse identifiee par : {0}' -f $how)
        }
        return Get-ElementValue $chosen
    }
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
        $windowName = Remove-InvisibleMarks (Get-ElementName $window)

        $url = Remove-InvisibleMarks (Get-AddressBarValue $window)
        if ($LogDetails) { Write-Log ("Fenetre Edge : '{0}' | URL de l'onglet actif : '{1}'" -f $windowName, $url) }
        if ((Test-PatternMatch $windowName ($TitlePattern + '*')) -and (Test-UrlMatch $url)) {
            return New-PageResult -Window $window -Title $windowName -Url $url
        }

        $allTabs = @()
        try { $allTabs = @($window.FindAll($script:Descendants, $script:CondTabItem)) } catch { }
        $candidates = @($allTabs | Where-Object {
                (Test-PatternMatch (Remove-InvisibleMarks (Get-ElementName $_)) ($TitlePattern + '*')) -and -not (Test-TabSelected $_) -and -not (Test-InWebContent $_)
            })
        if ($candidates.Count -eq 0) { continue }

        $initialTab = $allTabs | Where-Object { (Test-TabSelected $_) -and -not (Test-InWebContent $_) } | Select-Object -First 1
        foreach ($tab in $candidates) {
            $tabName = Remove-InvisibleMarks (Get-ElementName $tab)
            if (-not (Select-Tab $tab)) { continue }
            Start-Sleep -Milliseconds 500
            $url = Remove-InvisibleMarks (Get-AddressBarValue $window)
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

function Get-PromptTextForButton {
    param($Button)
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $node = $Button
    for ($level = 1; $level -le 3; $level++) {
        try {
            $node = $walker.GetParent($node)
            if (-not $node) { return $null }
            $names = @(Get-ElementName $node) + @($node.FindAll($script:Descendants, $script:CondAll) | ForEach-Object { Get-ElementName $_ })
            $saveAsButtons = @($node.FindAll($script:Descendants, $script:CondClickable) | Where-Object {
                    (Test-ClassMatch $_ $SaveAsButtonClasses) -or (Test-LikeAny (Get-ElementName $_) $SaveAsButtonNames)
                })
        }
        catch { return $null }
        if ($saveAsButtons.Count -gt 1) { return $null }
        $match = $names | Where-Object { $_ -like "*$FileNamePattern*" } | Select-Object -First 1
        if ($match) { return $match }
    }
    return $null
}

function Get-DocumentAncestorName {
    param($Element)
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $node = $Element
    for ($depth = 0; $depth -lt 40 -and $node; $depth++) {
        try {
            if ($node.Current.ControlType -eq [System.Windows.Automation.ControlType]::Document) { return (Get-ElementName $node) }
            $node = $walker.GetParent($node)
        }
        catch { return $null }
    }
    return $null
}

function Test-InTargetPage {
    param($Element)
    if (-not (($TitlePattern -replace '[\*\?]', '').Trim())) { return $false }
    $documentName = Get-DocumentAncestorName $Element
    if ($null -eq $documentName) { return $false }
    return (Test-PatternMatch (Remove-InvisibleMarks $documentName) ($TitlePattern + '*'))
}

function Find-SaveAsButton {
    param($Page)
    $buttons = @(Find-NamedElements -Roots (Get-PanelSearchRoots $Page) -Condition $script:CondClickable -NamePatterns $SaveAsButtonNames -Classes $SaveAsButtonClasses)
    foreach ($button in $buttons) {
        if (Test-InTargetPage $button) { continue }
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
    $buttons = @(Find-NamedElements -Roots @($Page.Window) -Condition $script:CondClickable -NamePatterns $DownloadsButtonNames -Classes $DownloadsButtonClasses -ExcludeWebContent)
    if ($buttons.Count -eq 0) { return $false }
    $identifiedBy = if (Test-ClassMatch $buttons[0] $DownloadsButtonClasses) { 'sa classe (independant de la langue)' } else { 'son nom (depend de la langue)' }
    Write-Log ('Bouton Telechargements identifie par {0}' -f $identifiedBy)
    Write-ElementDiagnostic -Element $buttons[0] -Label 'Bouton Telechargements'
    $action = Invoke-UiaElement $buttons[0]
    Write-Log ("Bouton '{0}' actionne ({1}) : ouverture du panneau Telechargements" -f (Get-ElementName $buttons[0]), $action)
    return $true
}

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


function Get-WindowSnapshot {
    $snapshot = @{}
    foreach ($window in [EdgeSaveAsCheck.Win32]::GetTopLevelWindows()) { $snapshot[[int64]$window.Handle] = $true }
    return $snapshot
}

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

function Invoke-SaveAsClickLadder {
    param($Page, $SaveAs)
    $known = Get-WindowSnapshot
    $reported = @{}
    $edgePids = @(Get-EdgeProcessIds)
    foreach ($window in [EdgeSaveAsCheck.Win32]::GetTopLevelWindows()) {
        if ($window.Visible -and $window.ClassName -eq '#32770' -and (($edgePids -contains [int]$window.ProcessId) -or $window.Owner -eq $Page.Handle)) {
            Write-Log ("Une boite de dialogue Windows issue d'Edge est deja ouverte ('{0}') : resultat d'un essai precedent ? Elle bloque Edge et fausse cette verification" -f $window.Title) 'WARN'
        }
    }
    Write-ElementDiagnostic -Element $SaveAs.Element -Label 'Bouton Enregistrer sous'
    if (Wait-ElementStable -Element $SaveAs.Element) { Write-Log 'Bouton stable : position lue deux fois identique' }
    else { Write-Log "Bouton non stabilise apres 1,5 s : clic tente quand meme" 'WARN' }
    Write-Log ('Etat avant le clic : {0}' -f (Get-StateText))
    $methods = @('MouseRoot', 'Invoke', 'accDoDefaultAction', 'MouseRoot', 'MouseChild', 'Keyboard')
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
    Write-Log ('Etat apres le clic : {0}' -f (Get-StateText))
    return [pscustomobject]@{ Dialog = $dialog; Tried = @($tried) }
}

function Wait-ElementStable {
    param($Element, [int]$PollMilliseconds = 150, [int]$MaxMilliseconds = 1500)
    $deadline = (Get-Date).AddMilliseconds($MaxMilliseconds)
    $previous = $null
    do {
        try {
            $rect = $Element.Current.BoundingRectangle
            if ($null -eq $rect) { return $false }
            $current = '{0:N0},{1:N0},{2:N0},{3:N0}' -f $rect.X, $rect.Y, $rect.Width, $rect.Height
        }
        catch { return $false }
        if ($current -eq $previous) { return $true }
        $previous = $current
        Start-Sleep -Milliseconds $PollMilliseconds
    } while ((Get-Date) -lt $deadline)
    return $false
}

function Get-StateText {
    try {
        $foreground = [EdgeSaveAsCheck.Win32]::ForegroundWindow()
        $text = 'aucune'
        if ($foreground -ne [IntPtr]::Zero) {
            $process = Get-Process -Id ([int][EdgeSaveAsCheck.Win32]::ProcessOf($foreground)) -ErrorAction SilentlyContinue
            $text = "classe '{0}' processus '{1}'" -f [EdgeSaveAsCheck.Win32]::ClassOf($foreground), $process.ProcessName
        }
        return ('fenetre au premier plan : {0} ; curseur : {1}' -f $text, [EdgeSaveAsCheck.Win32]::CursorText())
    }
    catch { return 'etat indisponible' }
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
    [void][EdgeSaveAsCheck.Win32]::PostMessage($Handle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
    $deadline = (Get-Date).AddSeconds(5)
    while ([EdgeSaveAsCheck.Win32]::IsWindow($Handle) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 200 }
    return -not [EdgeSaveAsCheck.Win32]::IsWindow($Handle)
}


function Invoke-EdgeSaveAsCheck {
    Initialize-Win32
    $dpiMode = 'indisponible'
    try { $dpiMode = [EdgeSaveAsCheck.Win32]::EnableDpiAwareness() }
    catch { $dpiMode = 'indisponible : ' + $_.Exception.Message }
    Write-Log ('Mode DPI du processus : {0}' -f $dpiMode)
    Initialize-Uia

    if ((Get-EdgeProcessIds).Count -eq 0) {
        Stop-Check 2 "Microsoft Edge n'est pas lance (aucun processus msedge)."
    }

    $page = $null
    for ($attempt = 1; $attempt -le 3 -and -not $page; $attempt++) {
        if ($attempt -gt 1) { Start-Sleep -Milliseconds 1500 }
        $page = Find-TargetEdgePage -LogDetails:($attempt -eq 1)
    }
    if (-not $page) {
        Stop-Check 3 ("Aucune page Edge ne correspond a URL='{0}' et Titre='{1}' (voir les fenetres listees ci-dessus)." -f $UrlPattern, $TitlePattern)
    }
    Write-Log ("Page cible : '{0}' | {1} | HWND 0x{2:X}" -f $page.Title, $page.Url, $page.Handle.ToInt64()) 'OK'
    $script:PageHandle = $page.Handle

    if ([EdgeSaveAsCheck.Win32]::IsIconic($page.Handle)) {
        [void][EdgeSaveAsCheck.Win32]::ShowWindow($page.Handle, 4)
        $script:MinimizedHandle = $page.Handle
        Start-Sleep -Milliseconds 700
        Write-Log 'Fenetre Edge reduite : restauree sans prendre le focus'
    }

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

    $result = Invoke-SaveAsClickLadder -Page $page -SaveAs $saveAs
    $dialog = $result.Dialog
    if (-not $dialog) {
        Stop-Check 6 ("La fenetre Windows d'enregistrement ne s'est pas ouverte (methodes essayees : {0}). Voir les lignes du journal sur le bouton et les nouvelles fenetres." -f ($result.Tried -join ', '))
    }
    $dialogFromEdge = (@(Get-EdgeProcessIds) -contains [int]$dialog.ProcessId) -or ($dialog.Owner -eq $page.Handle)
    Write-Log ("Fenetre Windows detectee : '{0}' (classe {1}, HWND 0x{2:X}, PID {3}, issue d'Edge : {4}, proprietaire 0x{5:X})" -f $dialog.Title, $dialog.ClassName, $dialog.Handle.ToInt64(), $dialog.ProcessId, $dialogFromEdge, $dialog.Owner.ToInt64()) 'OK'
    $fileName = Get-DialogFileName $dialog.Handle
    if ($fileName) { Write-Log ("Nom de fichier propose : '{0}'" -f $fileName) }

    if ($CloseDialog -and -not $dialogFromEdge) {
        Write-Log "Fenetre non fermee : elle n'appartient pas a Edge (boite d'une autre application ?)" 'WARN'
        $script:DialogLeftOpen = $true
    }
    elseif ($CloseDialog) {
        if (Close-SaveDialog $dialog.Handle) { Write-Log 'Fenetre fermee (Annuler : aucun fichier enregistre)' }
        else { Write-Log 'La fenetre ne s est pas fermee dans les 5 s' 'WARN'; $script:DialogLeftOpen = $true }
    }
    else {
        $script:DialogLeftOpen = $true
    }
    $script:ResultText = "fenetre '$($dialog.Title)' ouverte"
}


try {
    $script:LogFile = Initialize-Log -Path $LogPath
    Write-Log ('Demarrage - PID {0}, PowerShell {1}, utilisateur {2}, journal : {3}' -f $PID, $PSVersionTable.PSVersion, $env:USERNAME, $script:LogFile)

    $UrlPattern = (Remove-InvisibleMarks $UrlPattern).Trim()
    $TitlePattern = (Remove-InvisibleMarks $TitlePattern).Trim()
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
    if ($script:MinimizedHandle -ne [IntPtr]::Zero -and -not $script:DialogLeftOpen) {
        try { [void][EdgeSaveAsCheck.Win32]::ShowWindow($script:MinimizedHandle, 7) } catch { }
    }
    $elapsed = ((Get-Date) - $script:StartTime).TotalSeconds
    $level = if ($script:ExitCode -eq 0) { 'OK' } else { 'ERREUR' }
    Write-Log ('RESULTAT : code {0} - {1} - duree {2:N1} s' -f $script:ExitCode, $script:ResultText, $elapsed) $level
}
if ($script:QuietConsole) { Write-Output $script:ExitCode }
exit $script:ExitCode
} @parameters
