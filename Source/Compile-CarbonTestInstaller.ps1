<#
#>

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
# 
#     http://www.apache.org/licenses/LICENSE-2.0
# 
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [String] $Configuration
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 'Latest'
#Requires -Version 5.1

$idePath = ''
if( -not (Get-Command -Name 'devenv' -ErrorAction Ignore) )
{
    Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath '..\PSModules\VSSetup' -Resolve)
    $instances = Get-VSSetupInstance 
    $instances | Format-List | Out-String | Write-Verbose
    $instance =
        $instances | 
        Where-Object { $_.DisplayName -notlike '* Test Agent *' } |
        Sort-Object -Descending -Property InstallationVersion | 
        Select-Object -First 1

    if( -not $instance )
    {
        $vswherePath = Join-Path ${env:ProgramFiles(x86)} -ChildPath 'Microsoft Visual Studio\Installer\'
        $env:Path = "$($vswherePath)$([IO.Path]::PathSeparator)$($env:Path)"
        if( -not (Get-Command -Name 'vswhere' -ErrorAction Ignore) )
        {
            $vswherePath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath 'vswhere.exe'
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri 'https://github.com/microsoft/vswhere/releases/download/2.8.4/vswhere.exe' `
                              -OutFile $vsWherePath
        }
        $instance = & $vswherePath -latest -format json -nologo | ConvertFrom-Json | Select-Object -First 1
    }

    if( -not $instance )
    {
        $instance = 
            @('14', '12', '11', '10', '9', '8' ) |
            ForEach-Object { "env:VS$($_)0COMNTOOLS" } |
            Where-Object { Test-Path -Path $_ } |
            ForEach-Object { Get-Item -Path $_ } |
            Select-Object -ExpandProperty 'Value' |
            Split-Path |
            Join-Path -ChildPath 'IDE' |
            Join-Path -ChildPath 'devenv.exe' |
            Where-Object { Test-Path -Path $_ } |
            ForEach-Object {
                $devenvInfo = Get-Item -Path $_
                [pscustomobject]@{
                    ProductPath = $_;
                    DisplayName = $devenvInfo.VersionInfo.FileDescription;
                    InstallationVersion = $devenvInfo.VersionInfo.ProductVersion;
                }
            } |
            Select-Object -First 1
    }

    if( $instance )
    {
        $idePath = $instance.ProductPath | Split-Path -Parent
        $msg = "Using $($instance.DisplayName) $($instance.InstallationVersion) found at ""$($idePath)""."
        Write-Verbose -Message $msg
    }
    else
    {
        Get-ChildItem -Path 'env:'
        Write-Error -Message 'Unable to find "devenv".'
        exit 1
    }

    $env:PATH = "$($env:PATH)$([IO.Path]::PathSeparator)$($idePath)"
}

$installerSlnPath = Join-Path -Path $PSScriptRoot -ChildPath 'Carbon.Installer.sln' -Resolve
Write-Verbose ('devenv "{0}" /build "{1}"' -f $installerSlnPath,$Configuration)

if( (Test-Path -Path 'env:APPVEYOR_BUILD_WORKER_IMAGE') -and `
    ($env:APPVEYOR_BUILD_WORKER_IMAGE -eq 'Visual Studio 2013') )
{
    Push-Location -Path $idePath
    try
    {
        Write-Verbose "IdePath  $($idePath)" -Verbose
        Get-ChildItem -Recurse
        # & '.\CommonExtensions\Microsoft\VSI\DisableOutOfProcBuild\DisableOutOfProcBuild.exe'
    }
    finally
    {
        Pop-Location
    }
}

devenv $installerSlnPath /build $Configuration
if( $LASTEXITCODE )
{
    $msg = 'Failed to build Carbon test installers. Check the output above for details. If the build failed because ' +
           'of this error: "ERROR: An error occurred while validating. HRESULT = ''8000000A''", open a command ' +
           "prompt, move into the ""$($idePath)"" directory, and run " +
           '".\CommonExtensions\Microsoft\VSI\DisableOutOfProcBuild\DisableOutOfProcBuild.exe".'
    Write-Error -Message $msg
}