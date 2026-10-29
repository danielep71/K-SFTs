<#
.SYNOPSIS
Synchronize repository VBA source into an existing .xlsm workbook built from
src\workbook\SFT_Template.xlsx.

.DESCRIPTION
Updates VBA code only. A full synchronization imports the components that
tools\vba-projects.json lists for one of two workbooks, and asks which one
unless -Mode gives the answer:

- Production: the components under src\, for a workbook that is used to
  calculate.
- Development: the production components and every test module
  (tests\modules), for a workbook that runs the test suites. Run the tests
  on synthetic data only.

With -SourceDir only the standard and class modules found in that folder are
imported (see the third example). Standard and class modules are replaced from
source. ThisWorkbook and worksheet document modules keep their workbook
objects and only their code text is replaced. The script
does not write cells, formulas, names, formatting, tables, worksheets, or
workbook structure.

A full Development synchronization also prepares the evidence folder of an
Excel run (INSTALLATION.md, Recording an Excel run): test-results\excel-run
in the repository, or -EvidenceDir. After saving the workbook it removes the
files an earlier run left there (examples.log, injected.log, skipped.log,
host.txt and every file in exported\) and exports every standard, class and
form module of the workbook into exported\. Production and -SourceDir
synchronizations leave the folder alone.

-RecordRun records the Excel run in one step: a Development synchronization
as above, then the run itself. The script reopens the saved workbook in its
own Excel instance with macros enabled, runs
TEST_Examples.SaveExampleEvidence on the evidence folder, closes the
workbook without saving, and runs tools\check_excel_run.py on the folder
against the checked-out commit (HEAD), which needs Python 3.12 or later and
Git on PATH. tools\Record-KSFT-ExcelRun.cmd does the same for a workbook
dropped onto it.

Before the workbook is opened the script lists the worksheet and ThisWorkbook
modules in the source and asks whether to replace their code as well. Answer
N for a workbook built from a newer template, whose sheet code must stay as
it is. -IncludeSheets or -SkipSheets gives the answer in advance.

.NOTES
Requirements:
- Windows desktop Excel.
- Excel > Options > Trust Center > Macro Settings >
  "Trust access to the VBA project object model" enabled.
- VBA project not password-locked.

The script is optional. The manual import of INSTALLATION.md needs no such
setting; either route can record an Excel run.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-KSFT-VBA.ps1 -WorkbookPath "C:\path\K-SFT.xlsm"

Asks whether the workbook is a production or a development workbook, then,
when the source has worksheet or ThisWorkbook modules, whether to replace
their code as well.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-KSFT-VBA.ps1 -WorkbookPath "C:\path\K-SFT.xlsm" -Mode Development -SkipSheets

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-KSFT-VBA.ps1 -WorkbookPath "C:\path\K-SFT.xlsm" -RecordRun

Synchronizes in Development mode, saves the three logs of the Excel run and
checks them against HEAD. Exit code 3: synchronized, but the run was not
recorded.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-KSFT-VBA.ps1 -WorkbookPath "C:\path\K-SFT.xlsm" -SourceDir "$env:USERPROFILE\Downloads"

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-KSFT-VBA.ps1 -WorkbookPath "C:\path\K-SFT.xlsm" -SkipSheets

Asks for production or development, then imports every component except the
worksheet and ThisWorkbook modules, whose code it leaves as it is, without
asking about it.

With -SourceDir, only the modules found in that folder are imported, for
example modules downloaded for a test, without copying them into the
repository. A file is identified only by the Attribute VB_Name it declares,
never by its file name: a file without exactly one valid declaration is
ignored. It is taken when that VB_Name is a component of the
repository source, or when it is a standard module (.bas) whose name has a
project prefix (CORE_, SFT_, M_, TEST_), such as a test suite added by
the pull request under test; it is listed apart as a new module. A form
(.frm) is imported from the repository only, beside its .frx. Other
files in the folder are ignored. When several files carry the same VB_Name
(for example "CORE_Validation (1).bas"), the newest is taken. The script
lists the files and asks for confirmation unless -Force is given, converts
their line endings to CRLF and imports them. Every other
module in the workbook is left as it is.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $false)]
    [string]$RepoRoot,

    [Parameter(Mandatory = $false)]
    [switch]$NoBackup,

    [Parameter(Mandatory = $false)]
    [switch]$Visible,

    [Parameter(Mandatory = $false)]
    [string]$SourceDir,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSheets,

    [Parameter(Mandatory = $false)]
    [switch]$SkipSheets,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Production", "Development")]
    [string]$Mode,

    [Parameter(Mandatory = $false)]
    [string]$EvidenceDir,

    [Parameter(Mandatory = $false)]
    [switch]$RecordRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Resolve the repository root after parameter binding. $PSScriptRoot is not
# reliable as a parameter default in every Windows PowerShell invocation mode.
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $scriptPath = $MyInvocation.MyCommand.Path
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        throw "Cannot determine the script path. Supply -RepoRoot explicitly."
    }

    $scriptDirectory = Split-Path -Parent $scriptPath
    $RepoRoot = Split-Path -Parent $scriptDirectory
}

$vbext_ct_StdModule = 1
$vbext_ct_ClassModule = 2
$vbext_ct_MSForm = 3
$vbext_ct_Document = 100
$xlCalculationManual = -4135
$xlCalculationAutomatic = -4105
$msoAutomationSecurityLow = 1
$msoAutomationSecurityForceDisable = 3

# With -SourceDir, a standard module the repository checkout does not have
# is offered for import when its name has a project prefix; any other
# unknown file is ignored.
$NewModulePattern = '^(CORE|SFT|M|TEST)_[A-Za-z0-9_]+$'

function Resolve-FullPath {
    param([Parameter(Mandatory = $true)][string]$PathValue)
    return [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $PathValue).Path)
}

function Get-VbaComponentName {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    foreach ($line in [System.IO.File]::ReadLines($SourcePath)) {
        if ($line -match '^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$') {
            return $Matches[1]
        }
    }

    return [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
}

function Get-DeclaredVbaName {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    # The identity a file declares, for files from -SourceDir: exactly
    # one Attribute VB_Name line holding a valid VBA name. The file name is
    # never used instead, so a "CORE_Validation.bas" without the attribute
    # cannot replace CORE_Validation. Returns the name and, when there is none, why.
    # A list, not a set: a declaration repeated identically, or with another
    # case, is a second declaration too.
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($line in [System.IO.File]::ReadLines($SourcePath)) {
        if ($line -match '^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"\s*$') {
            $names.Add($Matches[1])
        }
        elseif ($line -match '^\s*Attribute\s+VB_Name\b') {
            return [pscustomobject]@{ Name = ""; Problem = "malformed Attribute VB_Name line" }
        }
    }

    if ($names.Count -eq 0) {
        return [pscustomobject]@{ Name = ""; Problem = "no Attribute VB_Name line; the file name is not used as its identity" }
    }
    if ($names.Count -gt 1) {
        return [pscustomobject]@{ Name = ""; Problem = "several Attribute VB_Name lines" }
    }

    $name = $names[0]
    if ($name -notmatch '^[A-Za-z][A-Za-z0-9_]{0,30}$') {
        return [pscustomobject]@{ Name = ""; Problem = "VB_Name '$name' is not a valid VBA name" }
    }

    return [pscustomobject]@{ Name = $name; Problem = "" }
}

function Get-VbaEncoding {
    # Repository VBA is ASCII only (docs/VBA_HOUSE_STYLE.md). Code text read
    # for a document module is decoded as cp1252, the code page the VBE
    # exports in, which reads ASCII unchanged and keeps any other byte as a
    # visible character rather than failing as UTF-8 would. Windows
    # PowerShell has the code page built in; PowerShell 7 may need the
    # code-page provider registered.
    try {
        return [System.Text.Encoding]::GetEncoding(1252)
    }
    catch {
        [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
        return [System.Text.Encoding]::GetEncoding(1252)
    }
}

function Test-DocumentSource {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    # A document module (ThisWorkbook, a worksheet or a chart sheet) is
    # exported with VB_Base naming Excel's Workbook (00020819), Worksheet
    # (00020820) or Chart (00020821) class, whatever its code name; an
    # ordinary class module has no VB_Base. The names Sync-VbaSource
    # recognizes (ThisWorkbook, sh*) also count, for a file without it.
    if ([System.IO.Path]::GetExtension($SourcePath).ToLowerInvariant() -ne ".cls") { return $false }
    foreach ($line in [System.IO.File]::ReadLines($SourcePath)) {
        if ($line -match '^\s*Attribute\s+VB_Base\s*=\s*"0\{(00020819|00020820|00020821)-') {
            return $true
        }
    }
    $name = Get-VbaComponentName -SourcePath $SourcePath
    return ($name -eq "ThisWorkbook" -or $name -like "sh*")
}

function Get-VbaCodeBody {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    $lines = [System.IO.File]::ReadAllLines($SourcePath, (Get-VbaEncoding))
    if ($lines.Count -eq 0) { return "" }

    # The module header is Attribute VB_<name> = ...; a member attribute
    # has a dot after the member's name, which may itself start with VB_,
    # so it never counts as the header and cuts no code above it.
    $lastAttribute = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*Attribute\s+VB_[^\s.=]*\s*=') {
            $lastAttribute = $i
        }
    }
    if (($lastAttribute + 1) -ge $lines.Count) { return "" }
    $body = $lines[($lastAttribute + 1)..($lines.Count - 1)]

    # An exported member attribute, such as
    # Attribute Worksheet_Change.VB_Description = "...", follows its
    # procedure's first line and is not valid as code, so it is dropped
    # wherever it is, with or without a module header. The member
    # name may hold code-page letters, so it is any run of characters
    # other than space, dot and equals sign.
    $code = @($body | Where-Object { $_ -notmatch '^\s*Attribute\s+[^\s.=]+\.VB_' })
    return [string]::Join([Environment]::NewLine, $code)
}

function Copy-WithCrlf {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$TargetPath
    )

    # ISO-8859-1 maps every byte to one character and back, so the bytes of
    # an exported module are kept exactly; only line endings change.
    $encoding = [System.Text.Encoding]::GetEncoding(28591)
    $text = $encoding.GetString([System.IO.File]::ReadAllBytes($SourcePath))
    $text = $text.Replace("`r`n", "`n").Replace("`r", "`n").Replace("`n", "`r`n")
    [System.IO.File]::WriteAllBytes($TargetPath, $encoding.GetBytes($text))
}

function Select-FolderSources {
    param(
        [Parameter(Mandatory = $true)][string]$FolderPath,
        [Parameter(Mandatory = $true)][hashtable]$RepositoryNames,
        [Parameter(Mandatory = $true)][string]$StagePath
    )

    $picked = @{}
    $newNames = @{}
    $ignored = New-Object System.Collections.Generic.List[string]
    $candidates = @(
        Get-ChildItem -LiteralPath $FolderPath -File |
            Where-Object { $_.Extension.ToLowerInvariant() -in @(".bas", ".cls", ".frm") }
    )

    foreach ($file in $candidates) {
        if ($file.Extension.ToLowerInvariant() -eq ".frm") {
            # A form imports with its .frx, which a downloads folder may not
            # hold beside it: forms come from the repository only.
            $ignored.Add("$($file.Name) (forms are imported from the repository only)")
            continue
        }
        $declared = Get-DeclaredVbaName -SourcePath $file.FullName
        if (-not [string]::IsNullOrEmpty($declared.Problem)) {
            $ignored.Add("$($file.Name) ($($declared.Problem))")
            continue
        }
        $name = $declared.Name
        $key = $name.ToLowerInvariant()
        if ($RepositoryNames.ContainsKey($key)) {
            $repoExtension = [System.IO.Path]::GetExtension($RepositoryNames[$key]).ToLowerInvariant()
            if ($file.Extension.ToLowerInvariant() -ne $repoExtension) {
                $ignored.Add("$($file.Name) ('$name' is a $repoExtension component in the repository)")
                continue
            }
        }
        elseif ($file.Extension.ToLowerInvariant() -eq ".bas" -and $name -match $NewModulePattern) {
            # A standard module the checkout does not have yet, for example a
            # test suite added by the pull request under test: listed apart.
            $newNames[$key] = $true
        }
        else {
            $ignored.Add("$($file.Name) (VB_Name '$name' is not a repository component)")
            continue
        }

        if ($picked.ContainsKey($key)) {
            if ($file.LastWriteTime -gt $picked[$key].LastWriteTime) {
                $ignored.Add("$($picked[$key].Name) (older copy of '$name')")
                $picked[$key] = $file
            }
            else {
                $ignored.Add("$($file.Name) (older copy of '$name')")
            }
        }
        else {
            $picked[$key] = $file
        }
    }

    if ($picked.Count -eq 0) {
        # Say why each file was passed over before stopping.
        if ($ignored.Count -gt 0) {
            Write-Host "Ignored:"
            foreach ($line in $ignored) { Write-Host "  $line" }
            Write-Host ""
        }
        throw "No repository VBA module found in $FolderPath."
    }

    Write-Host "Modules to import from $FolderPath"
    foreach ($key in ($picked.Keys | Sort-Object)) {
        if ($newNames.ContainsKey($key)) { continue }
        $file = $picked[$key]
        Write-Host ("  {0,-26} {1,-34} {2:yyyy-MM-dd HH:mm}" -f (Get-VbaComponentName -SourcePath $file.FullName), $file.Name, $file.LastWriteTime)
    }
    if ($newNames.Count -gt 0) {
        Write-Host "New modules, not in the repository checkout (imported as new standard modules):"
        foreach ($key in ($picked.Keys | Sort-Object)) {
            if (-not $newNames.ContainsKey($key)) { continue }
            $file = $picked[$key]
            Write-Host ("  {0,-26} {1,-34} {2:yyyy-MM-dd HH:mm}" -f (Get-VbaComponentName -SourcePath $file.FullName), $file.Name, $file.LastWriteTime)
        }
    }
    if ($ignored.Count -gt 0) {
        Write-Host "Ignored:"
        foreach ($line in $ignored) { Write-Host "  $line" }
    }
    Write-Host ""

    $null = New-Item -ItemType Directory -Path $StagePath -Force
    $staged = @(
        foreach ($key in ($picked.Keys | Sort-Object)) {
            $file = $picked[$key]
            $name = Get-VbaComponentName -SourcePath $file.FullName
            $target = Join-Path $StagePath ($name + $file.Extension.ToLowerInvariant())
            Copy-WithCrlf -SourcePath $file.FullName -TargetPath $target
            Get-Item -LiteralPath $target
        }
    )

    return $staged
}

function ConvertTo-SyncMode {
    # The answer to the production/development question: Production,
    # Development, or $null for anything else.
    param([AllowNull()][AllowEmptyString()][string]$Answer)

    switch -Regex (([string]$Answer).Trim()) {
        '^(?i)(p|prod|production)$' { return "Production" }
        '^(?i)(d|dev|development)$' { return "Development" }
    }
    return $null
}

function Get-ManifestSources {
    # The components of one workbook, from tools\vba-projects.json: the
    # modules of its configuration and the document modules, sorted by
    # path. Every listed file must exist.
    param(
        [Parameter(Mandatory = $true)][string]$RepoPath,
        [Parameter(Mandatory = $true)][ValidateSet("Production", "Development")][string]$Mode
    )

    $manifestPath = Join-Path (Join-Path $RepoPath "tools") "vba-projects.json"
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw "Component manifest not found: $manifestPath"
    }

    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $configuration = $manifest.configurations.($Mode.ToLowerInvariant())
    if ($null -eq $configuration) {
        throw "Configuration '$($Mode.ToLowerInvariant())' not found in $manifestPath"
    }

    $relative = @($configuration.modules) + @($manifest.document_modules)
    $files = @(
        foreach ($entry in $relative) {
            $path = Join-Path $RepoPath ([string]$entry)
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                throw "Component listed in $manifestPath not found: $entry"
            }
            Get-Item -LiteralPath $path
        }
    )
    return @($files | Sort-Object FullName)
}

function Find-VbaComponent {
    param(
        [Parameter(Mandatory = $true)]$VBProject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    foreach ($component in $VBProject.VBComponents) {
        if ([string]::Equals(
            [string]$component.Name,
            $Name,
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
            return $component
        }
    }

    return $null
}

function Get-SheetSnapshot {
    param([Parameter(Mandatory = $true)]$Workbook)

    $items = New-Object System.Collections.Generic.List[string]
    foreach ($sheet in $Workbook.Sheets) {
        $codeName = ""
        try { $codeName = [string]$sheet.CodeName }
        catch { $codeName = "<no-codename>" }

        $items.Add("$($sheet.Name)|$codeName")
    }

    return @($items)
}

function Assert-SheetSnapshotUnchanged {
    param(
        [Parameter(Mandatory = $true)][string[]]$Before,
        [Parameter(Mandatory = $true)][string[]]$After
    )

    if ($Before.Count -ne $After.Count) {
        throw "Workbook sheet count changed. Save aborted."
    }

    for ($i = 0; $i -lt $Before.Count; $i++) {
        if ($Before[$i] -cne $After[$i]) {
            throw "Workbook sheet structure changed. Save aborted."
        }
    }
}

function Replace-DocumentCode {
    param(
        [Parameter(Mandatory = $true)]$Component,
        [Parameter(Mandatory = $true)][string]$SourcePath
    )

    if ([int]$Component.Type -ne $vbext_ct_Document) {
        throw "Component '$($Component.Name)' is not a document module."
    }

    $module = $Component.CodeModule
    $count = [int]$module.CountOfLines
    if ($count -gt 0) {
        $module.DeleteLines(1, $count)
    }

    $code = Get-VbaCodeBody -SourcePath $SourcePath
    if (-not [string]::IsNullOrWhiteSpace($code)) {
        $module.AddFromString($code)
    }
}

function Sync-VbaSource {
    param(
        [Parameter(Mandatory = $true)]$VBProject,
        [Parameter(Mandatory = $true)][string]$SourcePath
    )

    $extension = [System.IO.Path]::GetExtension($SourcePath).ToLowerInvariant()
    $name = Get-VbaComponentName -SourcePath $SourcePath
    $existing = Find-VbaComponent -VBProject $VBProject -Name $name

    if ($null -ne $existing -and [int]$existing.Type -eq $vbext_ct_Document) {
        if ($extension -ne ".cls") {
            throw "Document module '$name' must come from a .cls source."
        }

        Replace-DocumentCode -Component $existing -SourcePath $SourcePath
        Write-Host ("{0,-30} document code replaced" -f $name)
        return
    }

    if ($extension -eq ".cls" -and $null -eq $existing) {
        if ($name -eq "ThisWorkbook" -or $name -like "sh*") {
            throw "Workbook document module '$name' was not found. Save aborted."
        }
    }

    switch ($extension) {
        ".bas" {
            if ($null -ne $existing) {
                if ([int]$existing.Type -ne $vbext_ct_StdModule) {
                    throw "Existing '$name' is not a standard module."
                }
                $VBProject.VBComponents.Remove($existing)
            }

            $imported = $VBProject.VBComponents.Import($SourcePath)
            Write-Host ("{0,-30} standard module imported" -f $imported.Name)
        }

        ".cls" {
            if ($null -ne $existing) {
                if ([int]$existing.Type -ne $vbext_ct_ClassModule) {
                    throw "Existing '$name' is not a class module."
                }
                $VBProject.VBComponents.Remove($existing)
            }

            $imported = $VBProject.VBComponents.Import($SourcePath)
            Write-Host ("{0,-30} class module imported" -f $imported.Name)
        }

        ".frm" {
            if ($null -ne $existing) {
                if ([int]$existing.Type -ne $vbext_ct_MSForm) {
                    throw "Existing '$name' is not a UserForm."
                }
                $VBProject.VBComponents.Remove($existing)
            }

            $imported = $VBProject.VBComponents.Import($SourcePath)
            Write-Host ("{0,-30} UserForm imported" -f $imported.Name)
        }

        default {
            throw "Unsupported VBA source extension '$extension'."
        }
    }
}

function Export-RunEvidence {
    # Start a new evidence folder: remove what an earlier run left, then
    # export the workbook's modules for check_excel_run.py.
    param(
        [Parameter(Mandatory = $true)]$VBProject,
        [Parameter(Mandatory = $true)][string]$FolderPath
    )

    $exportedPath = Join-Path $FolderPath "exported"
    $null = New-Item -ItemType Directory -Force -Path $exportedPath
    foreach ($name in @("examples.log", "injected.log", "skipped.log", "host.txt")) {
        $path = Join-Path $FolderPath $name
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            Remove-Item -LiteralPath $path -Force
        }
    }
    # exported\ holds only this export: any file left there, a module or
    # not, would fail check_excel_run.py.
    Get-ChildItem -LiteralPath $exportedPath -File | Remove-Item -Force

    $extensions = @{
        $vbext_ct_StdModule   = ".bas"
        $vbext_ct_ClassModule = ".cls"
        $vbext_ct_MSForm      = ".frm"
    }
    $count = 0
    foreach ($component in $VBProject.VBComponents) {
        $type = [int]$component.Type
        if ($extensions.ContainsKey($type)) {
            $component.Export((Join-Path $exportedPath ($component.Name + $extensions[$type])))
            $count++
        }
    }
    return $count
}

$excel = $null
$workbook = $null
$saved = $false
$backupPath = $null
$stagePath = $null
$syncMode = $null
$evidencePath = $null

try {
    $workbookFullPath = Resolve-FullPath -PathValue $WorkbookPath
    $repoFullPath = Resolve-FullPath -PathValue $RepoRoot

    if ([System.IO.Path]::GetExtension($workbookFullPath).ToLowerInvariant() -ne ".xlsm") {
        throw "WorkbookPath must point to an .xlsm file."
    }

    # The homes of docs/REPOSITORY_STRUCTURE.md. Each is created with its
    # first component, so a home that does not exist yet holds none.
    $sourceRoots = @(
        foreach ($relative in @("src\core", "src\modules", "src\classes", "src\workbook", "src\forms", "tests\modules")) {
            $root = Join-Path $repoFullPath $relative
            if (Test-Path -LiteralPath $root -PathType Container) { $root }
        }
    )

    $sourceFiles = @(
        foreach ($root in $sourceRoots) {
            Get-ChildItem -LiteralPath $root -File -Recurse |
                Where-Object {
                    $_.Extension.ToLowerInvariant() -in @(".bas", ".cls", ".frm")
                }
        }
    ) | Sort-Object FullName

    if ($sourceFiles.Count -eq 0) {
        throw "No VBA source files found."
    }

    $seen = @{}
    foreach ($file in $sourceFiles) {
        $name = Get-VbaComponentName -SourcePath $file.FullName
        $key = $name.ToLowerInvariant()

        if ($seen.ContainsKey($key)) {
            throw "Duplicate VBA component name '$name' in repository source."
        }

        $seen[$key] = $file.FullName
    }

    # Every component in the source folders must be listed in the manifest,
    # so that a new module is never left out of both workbooks unnoticed.
    $developmentFiles = @(Get-ManifestSources -RepoPath $repoFullPath -Mode Development)
    $listed = @{}
    foreach ($file in $developmentFiles) {
        $listed[$file.FullName.ToLowerInvariant()] = $true
    }
    foreach ($file in $sourceFiles) {
        if (-not $listed.ContainsKey($file.FullName.ToLowerInvariant())) {
            throw "Component not listed in tools\vba-projects.json: $($file.FullName)"
        }
    }

    # -RecordRun runs the test suite, which only a Development workbook has.
    if ($RecordRun) {
        if (-not [string]::IsNullOrWhiteSpace($SourceDir)) {
            throw "Use -RecordRun only with a Development synchronization, not with -SourceDir."
        }
        if ($Mode -eq "Production") {
            throw "Use -RecordRun only with a Development synchronization."
        }
        $Mode = "Development"
    }

    # Full synchronization: the production or the development component set.
    # -SourceDir instead takes what the folder holds.
    if (-not [string]::IsNullOrWhiteSpace($SourceDir) -and -not [string]::IsNullOrWhiteSpace($Mode)) {
        throw "Use -Mode only for a full synchronization, not with -SourceDir."
    }
    if (-not [string]::IsNullOrWhiteSpace($SourceDir) -and -not [string]::IsNullOrWhiteSpace($EvidenceDir)) {
        throw "Use -EvidenceDir only with a Development synchronization, not with -SourceDir."
    }

    if ([string]::IsNullOrWhiteSpace($SourceDir)) {
        $syncMode = $Mode
        if ([string]::IsNullOrWhiteSpace($syncMode)) {
            Write-Host "Which workbook is this?"
            Write-Host "  P = production: the components under src\"
            Write-Host "  D = development: the production components and every test module"
            Write-Host "      (run the tests on synthetic data only)"
            $syncMode = ConvertTo-SyncMode -Answer (Read-Host "Production or development? [P/D]")
            if ($null -eq $syncMode) {
                Write-Host "Cancelled: the workbook was not changed."
                exit 2
            }
        }

        if ($syncMode -eq "Development") {
            $evidencePath = $EvidenceDir
            if ([string]::IsNullOrWhiteSpace($evidencePath)) {
                $evidencePath = Join-Path $repoFullPath "test-results\excel-run"
            }
            # The folder need not exist yet: Export-RunEvidence creates it.
            $evidencePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($evidencePath)
        }
        elseif (-not [string]::IsNullOrWhiteSpace($EvidenceDir)) {
            throw "Use -EvidenceDir only with a Development synchronization."
        }

        $sourceFiles = @(Get-ManifestSources -RepoPath $repoFullPath -Mode $syncMode)
        Write-Host "Mode: $syncMode ($($sourceFiles.Count) components)"
        Write-Host ""
        if ($sourceFiles.Count -eq 0) {
            Write-Host "Nothing to import: tools\vba-projects.json lists no $($syncMode.ToLowerInvariant()) component. The workbook was not changed."
            exit 2
        }
    }

    # -SourceDir: import the repository modules found in that folder, and new
    # modules with a project prefix, after the user has seen the list. The repository source is used above
    # only to know which component names are valid.
    $sourceLabel = $repoFullPath
    if (-not [string]::IsNullOrWhiteSpace($SourceDir)) {
        $sourceDirFull = Resolve-FullPath -PathValue $SourceDir
        if (-not (Test-Path -LiteralPath $sourceDirFull -PathType Container)) {
            throw "SourceDir is not a folder: $sourceDirFull"
        }

        $stagePath = Join-Path ([System.IO.Path]::GetTempPath()) ("ksft-vba-sync-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
        $sourceFiles = @(Select-FolderSources -FolderPath $sourceDirFull -RepositoryNames $seen -StagePath $stagePath)
        $sourceLabel = $sourceDirFull

        if (-not $Force) {
            $answer = Read-Host "Import these $($sourceFiles.Count) module(s) into the workbook? [Y/N]"
            if ($answer -notmatch '^\s*(y|yes)\s*$') {
                Write-Host "Cancelled: the workbook was not changed."
                exit 2
            }
        }
    }

    # Worksheet and ThisWorkbook modules: their code is replaced only when the
    # user agrees, so that a workbook built from a newer template can keep
    # its own sheet code.
    if ($IncludeSheets -and $SkipSheets) {
        throw "Use -IncludeSheets or -SkipSheets, not both."
    }

    $documentFiles = @($sourceFiles | Where-Object { Test-DocumentSource -SourcePath $_.FullName })
    if ($documentFiles.Count -gt 0) {
        $importSheets = [bool]$IncludeSheets
        if (-not $IncludeSheets -and -not $SkipSheets) {
            Write-Host "Worksheet and ThisWorkbook modules in the source:"
            foreach ($file in $documentFiles) {
                Write-Host ("  {0}" -f (Get-VbaComponentName -SourcePath $file.FullName))
            }
            $answer = Read-Host "Replace the code of these $($documentFiles.Count) sheet module(s) as well? [Y/N]"
            $importSheets = ($answer -match '^\s*(y|yes)\s*$')
        }

        if ($importSheets) {
            Write-Host "Sheet modules: code replaced from the source."
        }
        else {
            $sourceFiles = @($sourceFiles | Where-Object { -not (Test-DocumentSource -SourcePath $_.FullName) })
            Write-Host "Sheet modules: skipped; the workbook keeps its own worksheet and ThisWorkbook code."
            if ($sourceFiles.Count -eq 0) {
                Write-Host "Nothing else to import: the workbook was not changed."
                exit 2
            }
        }
        Write-Host ""
    }

    if (-not $NoBackup) {
        $directory = [System.IO.Path]::GetDirectoryName($workbookFullPath)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($workbookFullPath)
        $extension = [System.IO.Path]::GetExtension($workbookFullPath)
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $backupPath = Join-Path $directory "$baseName.pre-vba-sync.$timestamp$extension"

        Copy-Item -LiteralPath $workbookFullPath -Destination $backupPath -Force
        Write-Host "Backup: $backupPath"
    }

    Write-Host "Workbook:   $workbookFullPath"
    Write-Host "Repository: $repoFullPath"
    Write-Host "Source:     $sourceLabel"
    Write-Host "Sources:    $($sourceFiles.Count)"
    Write-Host ""

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = $msoAutomationSecurityForceDisable

    # Some Excel builds reject Application.Calculation changes until at least
    # one workbook is open (HRESULT 0x800A03EC). Open first, then disable
    # calculation-before-save for this isolated Excel instance.
    $workbook = $excel.Workbooks.Open($workbookFullPath, 0, $false)

    if ($workbook.ReadOnly) {
        throw "Workbook opened read-only. Close any other Excel instance using it."
    }

    $excel.Calculation = $xlCalculationManual
    $excel.CalculateBeforeSave = $false

    $before = Get-SheetSnapshot -Workbook $workbook

    try {
        $vbProject = $workbook.VBProject
        $null = $vbProject.VBComponents.Count
    }
    catch {
        throw (
            "Cannot access the VBA project. Enable Excel Trust Center > Macro Settings > " +
            "'Trust access to the VBA project object model' and ensure the VBA project " +
            "is not password-locked."
        )
    }

    # A production synchronization never removes modules: test modules
    # already in the workbook stay, and the user is told.
    if ($syncMode -eq "Production") {
        $testModules = @(
            foreach ($component in $vbProject.VBComponents) {
                if ($component.Name -like "TEST_*") { $component.Name }
            }
        )
        if ($testModules.Count -gt 0) {
            Write-Warning ("This workbook already holds test modules, which a production " +
                "synchronization does not remove: " + ($testModules -join ", ") +
                ". For a production workbook, build it from the template.")
        }
    }

    foreach ($file in $sourceFiles) {
        Sync-VbaSource -VBProject $vbProject -SourcePath $file.FullName
    }

    $after = Get-SheetSnapshot -Workbook $workbook
    Assert-SheetSnapshotUnchanged -Before $before -After $after

    # The synchronization uses manual calculation only to avoid recalculation
    # while modules are being replaced. The workbook must leave this isolated
    # Excel instance in Automatic mode before it is saved, in either mode.
    $excel.Calculation = $xlCalculationAutomatic
    $excel.CalculateBeforeSave = $true

    $workbook.Save()
    $saved = $true

    # The backup is only a transactional safety copy. Remove it after a
    # successful save so repeated synchronizations do not accumulate files.
    if (-not [string]::IsNullOrWhiteSpace($backupPath) -and
        (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        Remove-Item -LiteralPath $backupPath -Force
        Write-Host "Backup removed: $backupPath"
    }

    if ($null -ne $evidencePath) {
        $exportedCount = Export-RunEvidence -VBProject $vbProject -FolderPath $evidencePath
        Write-Host "Evidence:   $exportedCount modules exported to $(Join-Path $evidencePath 'exported'); earlier run files removed"
    }

    # -RecordRun: run the suite in the saved workbook. Macros are enabled for
    # this Excel instance only; the workbook is opened read-only and closed
    # without saving, so the run leaves it as synchronized.
    if ($RecordRun) {
        $workbook.Close($false)
        $workbook = $null
        $excel.AutomationSecurity = $msoAutomationSecurityLow
        $workbook = $excel.Workbooks.Open($workbookFullPath, 0, $true)
        $macro = "'" + $workbook.Name.Replace("'", "''") + "'!TEST_Examples.SaveExampleEvidence"
        $null = $excel.Run($macro, $evidencePath)
        foreach ($name in @("examples.log", "injected.log", "skipped.log")) {
            if (-not (Test-Path -LiteralPath (Join-Path $evidencePath $name) -PathType Leaf)) {
                throw "SaveExampleEvidence did not save $name in $evidencePath."
            }
        }
        Write-Host "Run:        examples.log, injected.log and skipped.log saved in $evidencePath"
    }

    Write-Host ""
    Write-Host "SUCCESS: VBA synchronized. Workbook sheets were not removed or recreated."
}
catch {
    # Under $ErrorActionPreference = "Stop" a plain Write-Error would itself
    # stop the script here, before the warnings below and exit 1.
    Write-Error -ErrorRecord $_ -ErrorAction Continue
    if ($null -ne $workbook -and -not $saved) {
        Write-Warning "Synchronization failed. Workbook will close without saving."
    }

    if (-not [string]::IsNullOrWhiteSpace($backupPath) -and
        (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        Write-Warning "Safety backup retained: $backupPath"
    }

    exit 1
}
finally {
    if (-not [string]::IsNullOrWhiteSpace($stagePath) -and
        (Test-Path -LiteralPath $stagePath -PathType Container)) {
        try { Remove-Item -LiteralPath $stagePath -Recurse -Force } catch {}
    }

    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        } catch {}
    }

    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        } catch {}
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# -RecordRun: check the evidence against the checked-out commit, after Excel
# has closed. A failed check leaves the workbook synchronized: exit code 3.
if ($RecordRun) {
    Write-Host ""
    $commit = (& git -C $repoFullPath rev-parse HEAD 2>$null)
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($commit)) {
        Write-Warning "RUN NOT RECORDED: cannot read the commit with git rev-parse HEAD in $repoFullPath."
        exit 3
    }
    $commit = $commit.Trim()

    # The first candidate that runs Python 3.12 or later: "python" can be
    # the Microsoft Store alias, which runs nothing, beside a "py" launcher.
    $python = $null
    foreach ($candidate in @(Get-Command python, py -CommandType Application -ErrorAction SilentlyContinue)) {
        $usable = $false
        $previousPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = "Continue"
            $null = & $candidate.Path -c "import sys; sys.exit(0 if sys.version_info >= (3, 12) else 1)" 2>$null
            $usable = ($LASTEXITCODE -eq 0)
        }
        catch {
            $usable = $false
        }
        finally {
            $ErrorActionPreference = $previousPreference
        }
        if ($usable) {
            $python = $candidate.Path
            break
        }
    }
    if ($null -eq $python) {
        Write-Warning ("RUN NOT RECORDED: no Python 3.12 or later found. Run: python tools\check_excel_run.py " +
            "`"$evidencePath`" --commit $commit")
        exit 3
    }

    Write-Host "Check:      tools\check_excel_run.py against $commit, with $python"
    & $python (Join-Path (Join-Path $repoFullPath "tools") "check_excel_run.py") $evidencePath --commit $commit
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "RUN NOT RECORDED: check_excel_run.py did not pass (exit code $LASTEXITCODE)."
        exit 3
    }
    Write-Host ""
    Write-Host "RECORDED: the Excel run at $commit passed check_excel_run.py."
}
