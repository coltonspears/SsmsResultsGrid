param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",
    [switch]$CloseSsms,
    [switch]$RelaunchSsms,
    [switch]$SkipBuild,
    [switch]$SkipUninstall
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Write-Step {
    param([string]$Message)
    Write-Host "[build-reinstall] $Message" -ForegroundColor Cyan
}

function Invoke-External {
    param(
        [string]$FilePath,
        [string[]]$Arguments,
        [switch]$Quiet
    )

    if ($Quiet) {
        & $FilePath @Arguments | Out-Null
    } else {
        & $FilePath @Arguments
    }

    if ($LASTEXITCODE -ne 0) {
        $argString = if ($Arguments) { $Arguments -join " " } else { "" }
        throw "Command failed with exit code ${LASTEXITCODE}: $FilePath $argString"
    }
}

function Get-VsWherePath {
    $candidate = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $candidate) { return $candidate }
    return $null
}

function Resolve-MSBuildPath {
    $command = Get-Command msbuild -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }

    $vswhere = Get-VsWherePath
    if (-not $vswhere) { return $null }

    $installationPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
    if (-not [string]::IsNullOrWhiteSpace($installationPath)) {
        $msbuild = Join-Path $installationPath "MSBuild\Current\Bin\MSBuild.exe"
        if (Test-Path $msbuild) { return $msbuild }
    }

    return $null
}

function Resolve-VsixInstallerPath {
    $candidates = New-Object System.Collections.Generic.List[string]

    $vswhere = Get-VsWherePath
    if ($vswhere) {
        $installPath = & $vswhere -latest -products * -property installationPath
        if (-not [string]::IsNullOrWhiteSpace($installPath)) {
            $candidates.Add((Join-Path $installPath "Common7\IDE\VSIXInstaller.exe"))
        }
    }

    $candidates.Add((Join-Path ${env:ProgramFiles} "Microsoft SQL Server Management Studio 22\Common7\IDE\VSIXInstaller.exe"))
    $candidates.Add((Join-Path ${env:ProgramFiles(x86)} "Microsoft SQL Server Management Studio 22\Common7\IDE\VSIXInstaller.exe"))

    foreach ($path in $candidates) {
        if ($path -and (Test-Path $path)) { return $path }
    }

    return $null
}

function Resolve-SsmsExePath {
    param([string]$VsixInstallerPath)

    $command = Get-Command Ssms.exe -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }

    $candidates = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($VsixInstallerPath)) {
        $candidates.Add((Join-Path (Split-Path -Parent $VsixInstallerPath) "Ssms.exe"))
    }

    $candidates.Add((Join-Path ${env:ProgramFiles} "Microsoft SQL Server Management Studio 22\Release\Common7\IDE\Ssms.exe"))
    $candidates.Add((Join-Path ${env:ProgramFiles(x86)} "Microsoft SQL Server Management Studio 22\Release\Common7\IDE\Ssms.exe"))
    $candidates.Add((Join-Path ${env:ProgramFiles} "Microsoft SQL Server Management Studio 22\Common7\IDE\Ssms.exe"))
    $candidates.Add((Join-Path ${env:ProgramFiles(x86)} "Microsoft SQL Server Management Studio 22\Common7\IDE\Ssms.exe"))

    foreach ($path in $candidates) {
        if ($path -and (Test-Path $path)) { return $path }
    }

    return $null
}

function Stop-ToolingProcesses {
    Write-Step "Stopping VSIX installer processes (best effort)."
    Get-Process -Name "VSIXInstaller" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Remove-ExistingVsixForBuild {
    param([string]$VsixPath)

    if (-not (Test-Path $VsixPath)) { return }
    try {
        Remove-Item -LiteralPath $VsixPath -Force
        return
    } catch {
        Write-Step "VSIX file is locked. Waiting briefly and retrying."
        Start-Sleep -Seconds 1
    }

    Stop-ToolingProcesses
    Start-Sleep -Seconds 1
    Remove-Item -LiteralPath $VsixPath -Force
}

function Invoke-Build {
    param(
        [string]$SolutionPath,
        [string]$Configuration
    )

    $msbuild = Resolve-MSBuildPath
    if ($msbuild) {
        Write-Step "Using MSBuild: $msbuild"
        Invoke-External -FilePath $msbuild -Arguments @($SolutionPath, "/t:Restore", "/p:Configuration=$Configuration", "/m")
        Invoke-External -FilePath $msbuild -Arguments @($SolutionPath, "/t:Build", "/p:Configuration=$Configuration", "/m")
        return
    }

    $dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
    if (-not $dotnet) {
        throw "Neither MSBuild nor dotnet CLI is available."
    }

    Write-Step "MSBuild not found, falling back to dotnet build."
    Invoke-External -FilePath $dotnet.Source -Arguments @("build", $SolutionPath, "-c", $Configuration)
}

function Get-ManifestIdentity {
    param([string]$ManifestPath)

    [xml]$xml = Get-Content -LiteralPath $ManifestPath
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $nsMgr.AddNamespace("vsix", "http://schemas.microsoft.com/developer/vsx-schema/2011")

    $identityNode = $xml.SelectSingleNode("/vsix:PackageManifest/vsix:Metadata/vsix:Identity", $nsMgr)
    if (-not $identityNode) {
        throw "Could not read Identity from $ManifestPath"
    }

    return [pscustomobject]@{
        Id = $identityNode.Id
        Version = $identityNode.Version
    }
}

function Find-InstalledExtension {
    param(
        [string]$ExtensionId
    )

    $roots = @(
        (Join-Path $env:LocalAppData "Microsoft\SQL Server Management Studio\22.0\Extensions"),
        (Join-Path $env:LocalAppData "Microsoft\VisualStudio")
    )

    $hits = New-Object System.Collections.Generic.List[object]
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        $manifests = Get-ChildItem -LiteralPath $root -Filter "*.vsixmanifest" -Recurse -ErrorAction SilentlyContinue
        foreach ($manifest in $manifests) {
            try {
                [xml]$xml = Get-Content -LiteralPath $manifest.FullName
                $identity = $xml.PackageManifest.Metadata.Identity
                if (-not $identity) { continue }
                if ($identity.Id -eq $ExtensionId) {
                    $hits.Add([pscustomobject]@{
                        Path = $manifest.FullName
                        Version = [string]$identity.Version
                        LastWriteTime = $manifest.LastWriteTimeUtc
                    })
                }
            } catch {
                continue
            }
        }
    }

    if ($hits.Count -eq 0) { return $null }
    return $hits | Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

$repoRoot = Split-Path -Parent $PSCommandPath
$solutionPath = Join-Path $repoRoot "SsmsResultsGrid.sln"
$projectDir = Join-Path $repoRoot "src\SsmsResultsGrid"
$manifestPath = Join-Path $projectDir "source.extension.vsixmanifest"
$vsixPath = Join-Path $projectDir "bin\$Configuration\SsmsResultsGrid.vsix"

if (-not (Test-Path $solutionPath)) { throw "Solution not found at $solutionPath" }
if (-not (Test-Path $manifestPath)) { throw "Manifest not found at $manifestPath" }

$identity = Get-ManifestIdentity -ManifestPath $manifestPath
Write-Step "Extension ID: $($identity.Id)"
Write-Step "Expected version: $($identity.Version)"

if ($CloseSsms) {
    Write-Step "Stopping running SSMS processes."
    Get-Process -Name "Ssms" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

if (-not $SkipBuild) {
    Stop-ToolingProcesses
    Remove-ExistingVsixForBuild -VsixPath $vsixPath
    Write-Step "Building solution ($Configuration)."
    Invoke-Build -SolutionPath $solutionPath -Configuration $Configuration
}

if (-not (Test-Path $vsixPath)) {
    throw "VSIX not found at $vsixPath"
}

$vsixInstaller = Resolve-VsixInstallerPath
if (-not $vsixInstaller) {
    throw "VSIXInstaller.exe not found. Install Visual Studio Build/IDE tools or SSMS 22."
}

Write-Step "Using VSIXInstaller: $vsixInstaller"
if (-not $SkipUninstall) {
    Write-Step "Uninstalling previous extension version (best effort)."
    try {
        Invoke-External -FilePath $vsixInstaller -Arguments @("/quiet", "/uninstall:$($identity.Id)") -Quiet
    } catch {
        Write-Step "Previous uninstall may have failed or extension was not installed. Continuing."
    }
}

Write-Step "Installing VSIX: $vsixPath"
Invoke-External -FilePath $vsixInstaller -Arguments @("/quiet", "/shutdownprocesses", $vsixPath) -Quiet

Write-Step "Verifying installed extension version."
$installed = Find-InstalledExtension -ExtensionId $identity.Id
if (-not $installed) {
    Write-Warning "Install completed but extension manifest could not be located under LocalAppData."
    Write-Warning "Expected ID: $($identity.Id), expected version: $($identity.Version)"
    exit 0
}

Write-Step "Installed manifest: $($installed.Path)"
Write-Step "Installed version: $($installed.Version)"

if ($installed.Version -ne $identity.Version) {
    throw "Version mismatch. Expected $($identity.Version), found $($installed.Version)."
}

Write-Step "Success. Extension is installed with expected version $($identity.Version)."

$shouldRelaunch = $CloseSsms -or $RelaunchSsms
if ($shouldRelaunch) {
    $ssmsExe = Resolve-SsmsExePath -VsixInstallerPath $vsixInstaller
    if ($ssmsExe) {
        Write-Step "Relaunching SSMS: $ssmsExe"
        Start-Process -FilePath $ssmsExe | Out-Null
    } else {
        Write-Warning "Install succeeded, but SSMS executable could not be located for relaunch."
    }
}
