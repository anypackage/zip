using module AnyPackage
using namespace AnyPackage.Provider
using namespace System.IO
using namespace System.IO.Compression
using namespace System.Management.Automation

[PackageProvider('ZIP', FileExtensions = '.zip', PackageByName = $false)]
class ZipProvider : PackageProvider, IFindPackage, IGetPackage, IInstallPackage, IUninstallPackage, IUpdatePackage {
    [object] GetDynamicParameters([string] $commandName) {
        return $(switch ($commandName) {
                'Get-Package' { return [GetPackageDynamicParameters]::new() }
                'Install-Package' { return [InstallPackageDynamicParameters]::new() }
                'Uninstall-Package' { return [UninstallPackageDynamicParameters]::new() }
                default { return $null }
            })
    }

    [PackageProviderInfo] Initialize([PackageProviderInfo] $providerInfo) {
        return [ZipProviderInfo]::new($providerInfo)
    }

    [void] FindPackage([PackageRequest] $request) {
        $package = Get-PackageInfo -Path $request.Path
        $request.WritePackage($package)
    }

    [void] GetPackage([PackageRequest] $request) {
        $installPath = if ($request.DynamicParameters.Path) {
            $request.DynamicParameters.Path
        } else {
            $request.ProviderInfo.InstallPath
        }

        if (-not (Test-Path $installPath)) {
            return
        }

        $getChildItemParams = @{
            Path = (Join-Path $installPath '*/.package.json')
        }

        $files = Get-ChildItem @getChildItemParams

        foreach ($file in $files) {
            $package = Get-PackageInfo -Path $file

            if ($request.IsMatch($package.Name, $package.Version)) {
                $request.WritePackage($package)
            }
        }
    }

    [void] InstallPackage([PackageRequest] $request) {
        if ($request.ParameterSetName -eq 'InputObject') {
            $path = $request.Source
            $package = $request.Package
        } else {
            $path = $request.Path
            $package = Get-PackageInfo -Path $path
        }

        $getPackageParams = @{
            Name        = $package.Name
            Provider    = $request.ProviderInfo.FullName
            ErrorAction = 'SilentlyContinue'
        }

        $installedPackage = Get-Package @getPackageParams

        if ($installedPackage.Version -eq $package.Version) {
            $request.WriteVerbose('Package already installed')
            $request.WritePackage($package)
            return
        }

        $tempPath = Join-Path -Path ([Path]::GetTempPath()) -ChildPath ([Path]::GetRandomFileName())
        $request.WriteVerbose("Temp path: $tempPath")

        $request.WriteVerbose('Extracting package to temp path')
        Expand-Archive -Path $path -DestinationPath $tempPath -ErrorAction Stop

        $installScript = Join-Path -Path $tempPath -ChildPath 'tools/install.ps1'
        $request.WriteVerbose("Install script: $installScript")

        if (Test-Path -Path $installScript) {
            $request.WriteVerbose('Calling install script')
            if ($request.DynamicParameters.PackageParameters) {
                $installScriptParams = $request.DynamicParameters.PackageParameters
            } else {
                $installScriptParams = @{ }
            }

            $installScriptParams['Verbose'] = $true
            $installScriptParams['Debug'] = $true

            & $installScript @installScriptParams 2>&1 3>&1 4>&1 5>&1 6>&1 | Write-PackageTrace -Request $request
        } else {
            $request.WriteVerbose('Install script not found.')
        }

        $installPath = Join-Path -Path $request.ProviderInfo.InstallPath -ChildPath $package.Name
        $request.WriteVerbose("Package cache path: $installPath")

        if (Test-Path -Path $installPath) {
            $request.WriteVerbose("Removing existing package cache for: $($package.Name)")
            Remove-Item -Path "$installPath/*" -Recurse -ErrorAction Stop
        } else {
            New-Item -Path $installPath -ItemType Directory -ErrorAction Stop
        }

        $request.WriteVerbose('Moving files to package cache')
        $packagePath = Join-Path -Path $tempPath -ChildPath '.package.json'
        Move-Item -Path $packagePath -Destination $installPath -ErrorAction Stop

        $toolsPath = Join-Path -Path $tempPath -ChildPath 'tools'

        if (Test-Path -Path $toolsPath) {
            Move-Item -Path $toolsPath -Destination $installPath -ErrorAction Stop
        }

        $request.WriteVerbose('Removing temp directory')
        Remove-Item -Path $tempPath -Recurse

        $request.WritePackage($package)
    }

    [void] UninstallPackage([PackageRequest] $request) {
        if ($request.ParameterSetName -eq 'Name') {
            $getPackageParams = @{
                Name     = $request.Name
                Provider = $request.ProviderInfo.FullName
            }

            if ($request.Version) {
                $getPackageParams['Version'] = $request.Version
            }

            $packages = Get-Package @getPackageParams
        } else {
            $packages = $request.Package
        }

        foreach ($package in $packages) {
            $installPath = Split-Path -Path $package.Source.Location -Parent

            if (Test-Path -Path $installPath) {
                $uninstallScript = Join-Path -Path $installPath -ChildPath 'tools/uninstall.ps1'
                $request.WriteVerbose("Uninstall script: $uninstallScript")

                if (Test-Path -Path $uninstallScript) {
                    $request.WriteVerbose('Calling uninstall script')
                    if ($request.DynamicParameters.PackageParameters) {
                        $uninstallScriptParams = $request.DynamicParameters.PackageParameters
                    } else {
                        $uninstallScriptParams = @{ }
                    }

                    $uninstallScriptParams['Verbose'] = $true
                    $uninstallScriptParams['Debug'] = $true

                    & $uninstallScript @uninstallScriptParams 2>&1 3>&1 4>&1 5>&1 6>&1 | Write-PackageTrace -Request $request
                } else {
                    $request.WriteVerbose('Uninstall script not found.')
                }

                Remove-Item -Path $installPath -Recurse -ErrorAction Stop
                $request.WritePackage($package)
            }
        }
    }

    [void] UpdatePackage([PackageRequest] $request) {
        if ($request.ParameterSetName -eq 'Path') {
            $findPackageParams = @{
                Path     = $request.Path
                Provider = $request.ProviderInfo.FullName
            }

            $findPackage = Find-Package @findPackageParams
        } else {
            $findPackage = $request.Package
        }

        if ($null -eq $findPackage) {
            return
        }

        $getPackageParams = @{
            Name     = $findPackage.Name
            Provider = $request.ProviderInfo.FullName
        }

        $latest = Get-Package @getPackageParams |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if ($null -eq $latest) {
            return
        }

        if ($findPackage.Version -lt $latest.Version) {
            throw "Package '$($findPackage.Name)' version '$($findPackage.Version)' is less than installed version '$($latest.Version)'."
        }

        $package = $findPackage | Install-Package -PassThru -ErrorAction Stop
        $request.WritePackage($package)
    }
}

class GetPackageDynamicParameters {
    [Parameter()]
    [string]
    $Path
}

class InstallPackageDynamicParameters {
    [Parameter()]
    [hashtable]
    $PackageParameters
}

class UninstallPackageDynamicParameters {
    [Parameter()]
    [hashtable]
    $PackageParameters
}

class ZipProviderInfo : PackageProviderInfo {
    [string] $InstallPath

    ZipProviderInfo([PackageProviderInfo] $providerInfo) : base($providerInfo) {
        if ($global:IsLinux -or $global:IsMacOS) {
            $this.InstallPath = Join-Path -Path $global:Home -ChildPath '.local/share/anypackage/zip'
        } else {
            $this.InstallPath = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'anypackage/zip'
        }
    }
}

function Get-PackageInfo {
    param(
        [string]
        $Path
    )

    if ([Path]::GetExtension($Path) -eq '.zip') {
        try {
            $fs = [FileStream]::new($Path, [FileMode]::Open)
            $zip = [ZipArchive]::new($fs)
            $file = $zip.Entries | Where-Object Name -EQ '.package.json'
            if (-not $file) { throw '.package.json not found in zip file.' }
            $sr = [StreamReader]::new($file.Open())
            $info = $sr.ReadToEnd() | ConvertFrom-Json
        } finally {
            if ($fs) { $fs.Dispose() }
            if ($zip) { $zip.Dispose() }
            if ($sr) { $sr.Dispose() }
        }
    } elseif ([Path]::GetExtension($Path) -eq '.json') {
        $info = Get-Content -Path $Path -ErrorAction Stop | ConvertFrom-Json
    }

    $sourceParams = @{
        Name     = $Path
        Location = $Path
        Provider = $request.ProviderInfo
    }

    $source = New-SourceInfo @sourceParams

    $packageParams = @{
        Name        = $info.Name
        Version     = $info.Version
        Source      = $source
        Description = $info.Description
        Metadata    = ($info.Metadata | ConvertTo-Hashtable)
        Provider    = $request.ProviderInfo
    }

    New-PackageInfo @packageParams
}

function ConvertTo-Hashtable {
    param (
        [Parameter(ValueFromPipeline)]
        [PSObject]
        $InputObject
    )

    process {
        $props = $InputObject | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
        $ht = @{ }

        foreach ($prop in $props) {
            $ht[$prop] = $InputObject.$prop
        }

        $ht
    }
}

function Write-PackageTrace {
    param (
        [Parameter(ValueFromPipeline)]
        [Object]
        $InputObject,

        [Parameter()]
        [Request]
        $Request
    )

    process {
        switch ($InputObject.GetType()) {
            { $_ -eq [VerboseRecord] } { $Request.WriteVerbose($InputObject.Message) }
            { $_ -eq [DebugRecord] } { $Request.WriteDebug($InputObject.Message) }
            { $_ -eq [WarningRecord] } { $Request.WriteWarning($InputObject.Message) }
            { $_ -eq [ErrorRecord] } { $Request.WriteError($InputObject) }
            { $_ -eq [InformationalRecord] } { $Request.WriteInformation($InputObject) }
        }
    }
}

[guid] $id = 'f502ce32-5147-4e46-a774-d2dbd6acad67'
[PackageProviderManager]::RegisterProvider($id, [ZipProvider], $MyInvocation.MyCommand.ScriptBlock.Module)

$MyInvocation.MyCommand.ScriptBlock.Module.OnRemove = {
    [PackageProviderManager]::UnregisterProvider($id)
}
