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

& (Join-Path -Path $PSScriptRoot -ChildPath 'Initialize-Test.ps1' -Resolve)

$msiRootPath = Join-Path -Path $PSScriptRoot -ChildPath 'MSI' -Resolve
$carbonTestInstallerPath = Join-Path -Path $msiRootPath -ChildPath 'CarbonTestInstaller.msi' -Resolve
$carbonTestInstallerActionsPath =
    Join-Path -Path $msiRootPath -ChildPath 'CarbonTestInstallerWithCustomActions.msi' -Resolve
$isPwsh6 = $PSVersionTable['PSVersion'].Major -eq 6
if( $isPwsh6 )
{
    # PowerShell 6 doesn't support COM APIs
    return
}

$testInstaller = Get-CMsi -Path $carbonTestInstallerPath
$testInstallerWithActions = Get-CMsi -Path $carbonTestInstallerActionsPath
$testRoot = $null
$testNum = 0

function Assert-CarbonTestInstallerInstalled
{
    $Global:Error.Count | Should -Be 0
    $maxTries = 200
    $tryNum = 0
    $writeNewline = $false
    do
    {
        $item = Get-CInstalledProgram -Name 'Carbon Test Installer*' -ErrorAction Ignore
        if( $item )
        {
            break
        }

        Write-Host '.' -NoNewline
        $writeNewline = $true
        Start-Sleep -Milliseconds 100
    }
    while( $tryNum++ -lt $maxTries )
    if( $writeNewline )
    {
        Write-Host ''
    }
    $item | Should -Not -BeNullOrEmpty
}

function Assert-CarbonTestInstallerNotInstalled
{
    $maxTries = 200
    $tryNum = 0
    $writeNewline = $false
    do
    {
        $item = Get-CInstalledProgram -Name 'Carbon Test Installer*' -ErrorAction Ignore
        if( -not $item )
        {
            break
        }

        Write-Host '.' -NoNewline
        $writeNewline = $true
        Start-Sleep -Milliseconds 100
    }
    while( $tryNum++ -lt $maxTries )

    if( $writeNewline )
    {
        Write-Host ''
    }
    $item | Should -BeNullOrEmpty
}

function GivenInstalled
{
    Install-CMsi -Path $carbonTestInstallerPath
    Assert-CarbonTestInstallerInstalled
}

function Init
{
    Uninstall-CarbonTestInstaller

    $script:testRoot = Join-Path -Path $TestDrive.FullName -ChildPath ($script:testNum++)
    New-Item -Path $script:testRoot -ItemType 'Directory' | Out-Null

    $Global:Error.Clear()
}

function Reset
{
    Uninstall-CarbonTestInstaller
}

function ThenMsi
{
    [CmdletBinding()]
    param(
        [switch] $Not,

        [switch] $Downloaded,

        [switch] $Installed,

        [switch] $InTempDirectory,

        [switch] $Reinstalled
    )

    if( $Downloaded )
    {
        $times = 1
        if( $Not )
        {
            $times = 0
        }
        Assert-MockCalled -CommandName 'Invoke-WebRequest' -ModuleName 'Carbon.Windows.Installer' -Times $times
    }

    if( $Installed )
    {
        if( $Not )
        {
            Assert-CarbonTestInstallerNotInstalled
        }
        else
        {
            Assert-CarbonTestInstallerInstalled
        }
    }

    if( $InTempDirectory )
    {
        $path = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath 'Carbon*.msi'
        $path | Should -Not:$Not -Exist
    }

    if( $Reinstalled )
    {
        $times = 1
        if( $Not )
        {
            $times = 0
        }
        Assert-MockCalled -CommandName 'Start-Process' -ModuleName 'Carbon.Windows.Installer' -Times $times
    }
}

function Uninstall-CarbonTestInstaller
{
    Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'MSI') -Filter '*.msi' |
        Get-CMsi |
        Where-Object { Get-CInstalledProgram -Name $_.ProductName -ErrorAction Ignore } |
        ForEach-Object {
            #msiexec /fa $_.Path /quiet /l*vx 'D:\restore.log'
            $msiProcess = Start-Process -FilePath "msiexec.exe" -ArgumentList "/quiet","/fa",('"{0}"' -f $_.Path) -NoNewWindow -Wait -PassThru
            if( $null -ne $msiProcess.ExitCode -and $msiProcess.ExitCode -ne 0 )
            {
                Write-Error ("{0} {1} repair failed. (Exit code: {2}; MSI: {3})" -f $_.ProductName,$_.ProductVersion,$msiProcess.ExitCode,$_.Path)
            }
            #msiexec /uninstall $_.Path /quiet /l*vx 'D:\uninstall.log'
            $msiProcess = Start-Process -FilePath "msiexec.exe" -ArgumentList "/quiet","/uninstall",('"{0}"' -f $_.Path) -NoNewWindow -Wait -PassThru
            if( $null -ne $msiProcess.ExitCode -and $msiProcess.ExitCode -ne 0 )
            {
                Write-Error ("{0} {1} uninstall failed. (Exit code: {2}; MSI: {3})" -f $_.ProductName,$_.ProductVersion,$msiProcess.ExitCode,$_.Path)
            }
        }
    Assert-CarbonTestInstallerNotInstalled

    Get-ChildItem -Path ([IO.Path]::GetTempPath()) -Filter '*.msi' | Remove-Item -Force
}

function WhenInstalling
{
    [CmdletBinding(DefaultParameterSetName='FromFile')]
    param(
        [Parameter(Mandatory, ParameterSetName='FromWeb')]
        [switch] $FromWeb,

        [Parameter(ParameterSetName='FromWeb')]
        [Uri] $AtUrl = 'https://httpstat.us/200' ,

        [Parameter(ParameterSetName='FromWeb')]
        $WithExpectedChecksum,

        [switch] $MockInstall,

        [switch] $Forced
    )

    if( $MockInstall )
    {
        New-Alias -Name 'Start-InstallMsiTestProcess' -Value 'Microsoft.PowerShell.Management\Start-Process'
        Mock -CommandName 'Start-Process' `
             -ModuleName 'Carbon.Windows.Installer' `
             -ParameterFilter { $FilePath -eq 'msiexec.exe' } `
             -MockWith { Start-Process -FilePath 'cmd' -ArgumentList '/c', 'exit' -NoNewWindow -PassThru }
    }

    if( $FromWeb )
    {
        $installerPath = $script:carbonTestInstallerPath
        Mock -CommandName 'Invoke-WebRequest' `
             -ModuleName 'Carbon.Windows.Installer' `
             -MockWith {
                 param(
                     [String] $OutFile
                 )
                 Copy-Item -Path $installerPath -Destination $OutFile }.GetNewClosure()

        if( -not $WithExpectedChecksum )
        {
            $WithExpectedChecksum = (Get-FileHash -Path $carbonTestInstallerPath).Hash
        }
        $output = Install-CMsi -Url $AtUrl `
                               -Checksum $WithExpectedChecksum `
                               -ProductName $testInstaller.ProductName `
                               -ProductCode $testInstaller.ProductCode `
                               -Force:$Forced
        $output | Should -BeNullOrEmpty
        return
    }

    $output = Install-CMsi -Path $script:carbonTestInstallerPath -Force:$Forced
    $output | Should -BeNullOrEmpty
}

Describe 'Install-Msi.when passed path to a non-MSI file' {
    AfterEach { Reset }
    It 'should validate file is an MSI' {
        Init
        Install-CMsi -Path $PSCommandPath -ErrorAction SilentlyContinue
        $Global:Error.Count | Should -BeGreaterThan 0
        $Global:Error[0] | Should -Match 'Exception opening MSI database'
    }
}

Describe 'Install-Msi.when using WhatIf' {
    AfterEach { Reset }
    It 'should support what if' {
        Init
        Assert-CarbonTestInstallerNotInstalled
        Install-CMsi -Path $carbonTestInstallerPath -WhatIf
        $Global:Error.Count | Should -Be 0
        Assert-CarbonTestInstallerNotInstalled
    }
}

Describe 'Install-Msi.when installing' {
    AfterEach { Reset }
    It 'should install msi' {
        Init
        Assert-CarbonTestInstallerNotInstalled
        Install-CMsi -Path $carbonTestInstallerPath
        Assert-CarbonTestInstallerInstalled
    }
}

Describe 'Install-Msi.when installer fails' {
    AfterEach { Reset }
    It 'should handle failed installer' {
        Init
        $logFilePath = $carbonTestInstallerActionsPath | Split-Path -Leaf
        $logFilePath = "$($logFilePath).*.*.log"
        $logFilePath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath $logFilePath
        Get-Item -Path $logFilePath -ErrorAction Ignore | Remove-Item
        $envVarName = 'CARBON_TEST_INSTALLER_THROW_INSTALL_EXCEPTION'
        [Environment]::SetEnvironmentVariable($envVarName, $true.ToString(), 'User')
        try
        {
            Install-CMsi -Path $carbonTestInstallerActionsPath -ErrorAction SilentlyContinue
            Assert-CarbonTestInstallerNotInstalled
        }
        finally
        {
            [Environment]::SetEnvironmentVariable($envVarName, '', 'User')
        }
        $logFilePath | Should -Exist
        $logFilePath | Should -FileContentMatch -ExpectedContent 'msiexec.exe'
        $logFilePath | Should -FileContentMatch -ExpectedContent 'MSI \(s\)'
        $logFilePath | Should -FileContentMatch -ExpectedContent '^DEBUG  :'
    }
}

Describe 'Install-Msi.when using wildcards in path to installer' {
    AfterEach { Reset }
    It 'should support wildcards' {
        Init
        Copy-Item $carbonTestInstallerPath -Destination (Join-Path -Path $script:testRoot -ChildPath 'One.msi')
        Copy-Item $carbonTestInstallerPath -Destination (Join-Path -Path $script:testRoot -ChildPath 'Two.msi')
        Install-CMsi -Path (Join-Path -Path $script:testRoot -ChildPath '*.msi')
        Assert-CarbonTestInstallerInstalled
    }
}

Describe 'Install-Msi.when already installed' {
    AfterEach { Reset }
    It 'should not reinstall if already installed' {
        Init
        GivenInstalled
        WhenInstalling -MockInstall
        ThenMsi -Not -Reinstalled
    }
}

Describe 'Install-Msi.when forcing install' {
    AfterEach { Reset }
    It 'should repair' {
        Init
        GivenInstalled
        WhenInstalling -Forced -MockInstall
        ThenMsi -Reinstalled
    }
}

Describe 'Install-Msi.when there are spaces in the path to the MSI' {
    AfterEach { Reset }
    It 'should install msi with spaces in path' {
        Init
        $newInstaller = Join-Path -Path $script:testRoot -ChildPath 'Installer With Spaces.msi'
        Copy-Item -Path $carbonTestInstallerPath -Destination $newInstaller
        Install-CMsi -Path $newInstaller
        Assert-CarbonTestInstallerInstalled
    }
}

Describe 'Install-Msi.when downloading installer' {
    AfterEach { Reset }
    It 'should download and install the program' {
        Init
        WhenInstalling -FromWeb
        ThenMsi -Downloaded -Installed
        ThenMsi -Not -InTempDirectory
    }
}

Describe 'Install-Msi.when installer to download already installed' {
    AfterEach { Reset }
    It 'should not download or install the program' {
        Init
        GivenInstalled
        WhenInstalling -FromWeb -MockInstall
        ThenMsi -Not -Downloaded
        ThenMsi -Not -Reinstalled
    }
}

Describe 'Install-Msi.when forcing install of downloaded installer already installed' {
    AfterEach { Reset }
    It 'should download and repair the program' {
        Init
        GivenInstalled
        WhenInstalling -FromWeb -MockInstall -Forced
        ThenMsi -Downloaded
        ThenMsi -Reinstalled
    }
}

Describe 'Install-Msi.when downloaded file doesn''t match checksum' {
    AfterEach { Reset }
    It 'should download but not install the program' {
        Init
        WhenInstalling -FromWeb -WithExpectedChecksum 'deadbee' -ErrorAction SilentlyContinue
        ThenMsi -Downloaded
        ThenMsi -Not -Installed
        $Global:Error | Should -Match 'does not match expected checksum'
    }
}

Describe 'Install-Msi.when downloaded file url doesn''t have a path' {
    AfterEach { Reset }
    It 'should use url as downloaded file name' {
        Init
        WhenInstalling -FromWeb -AtUrl 'https://httpstat.us/'
        ThenMsi -Downloaded
        ThenMsi -Installed
        Assert-MockCalled -CommandName 'Invoke-WebRequest' `
                          -ModuleName 'Carbon.Windows.Installer' `
                          -ParameterFilter { $OutFile -match '\\https___httpstat\.us_' }
    }
}