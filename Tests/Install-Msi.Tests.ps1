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
if( -not $isPwsh6 )
{
    $testInstallerWithActions = Get-CMsi -Path $carbonTestInstallerActionsPath
}
$testRoot = $null
$testNum = 0

function Assert-CarbonTestInstallerInstalled
{
    $Global:Error.Count | Should -Be 0
    $maxTries = 200
    $tryNum = 0
    do
    {
        $item = Get-CInstalledProgram -Name 'Carbon Test Installer*' -ErrorAction Ignore
        if( $item )
        {
            break
        }

        Start-Sleep -Milliseconds 100
    }
    while( $tryNum++ -lt $maxTries )
    $item | Should -Not -BeNullOrEmpty
}

function Assert-CarbonTestInstallerNotInstalled
{
    $maxTries = 200
    $tryNum = 0
    do
    {
        $item = Get-CInstalledProgram -Name 'Carbon Test Installer*' -ErrorAction Ignore
        if( -not $item )
        {
            break
        }

        Start-Sleep -Milliseconds 100
    }
    while( $tryNum++ -lt $maxTries )

    $item | Should -BeNullOrEmpty
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

    # @( 'CarbonTestInstaller.msi', 'CarbonTestInstallerWithCustomActions.msi', 'One.msi', 'Two.msi', 'Installer With Spaces.msi') | 
    #     ForEach-Object { & "C:\Sysinternals\handle.exe" ($_ | Split-Path -Leaf) -nobanner } |
    #     Where-Object { $_ } |
    #     ForEach-Object { "  $($_)" } |
    #     Write-Verbose -Verbose
}


Describe 'Install-Msi.when passed path to a non-MSI file' {
    AfterEach { Reset }
    It 'should validate file is an MSI' -Skip:$isPwsh6 {
        Init
        Install-CMsi -Path $PSCommandPath -ErrorAction SilentlyContinue
        $Global:Error.Count | Should -BeGreaterThan 0
        $Global:Error[0] | Should -Match 'Exception opening MSI database'
    }
}

Describe 'Install-Msi.when using WhatIf' {
    AfterEach { Reset }
    It 'should support what if' -Skip:$isPwsh6 {
        Init
        Assert-CarbonTestInstallerNotInstalled
        Install-CMsi -Path $carbonTestInstallerPath -WhatIf
        $Global:Error.Count | Should -Be 0
        Assert-CarbonTestInstallerNotInstalled
    }
}

Describe 'Install-Msi.when installing' {
    AfterEach { Reset }
    It 'should install msi' -Skip:$isPwsh6 {
        Init
        Assert-CarbonTestInstallerNotInstalled
        Install-CMsi -Path $carbonTestInstallerPath
        Assert-CarbonTestInstallerInstalled
    }
}

Describe 'Install-Msi.when installer fails' {
    AfterEach { Reset }
    It 'should handle failed installer' -Skip:$isPwsh6 {
        Init
        $logFilePath = $carbonTestInstallerActionsPath | Split-Path -Leaf
        $logFilePath = "$($logFilePath).*.*.log"
        $logFilePath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath $logFilePath
        Get-Item -Path $logFilePath -ErrorAction Ignore | Remove-Item -Verbose
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
    It 'should support wildcards' -Skip:$isPwsh6 {
        Init
        Copy-Item $carbonTestInstallerPath -Destination (Join-Path -Path $script:testRoot -ChildPath 'One.msi')
        Copy-Item $carbonTestInstallerPath -Destination (Join-Path -Path $script:testRoot -ChildPath 'Two.msi')
        Install-CMsi -Path (Join-Path -Path $script:testRoot -ChildPath '*.msi')
        Assert-CarbonTestInstallerInstalled
    }
}

Describe 'Install-Msi.when product already installed' {
    AfterEach { Reset }
    It 'should not reinstall if already installed' -Skip:$isPwsh6 {
        Init
        Install-CMsi -Path $carbonTestInstallerActionsPath
        Assert-CarbonTestInstallerInstalled
        $installDir = Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath $testInstallerWithActions.Manufacturer
        $installDir = Join-Path -Path $installDir -ChildPath $testInstallerWithActions.ProductName
        $installDir | Should -Exist
        $tempName = [IO.Path]::GetRandomFileName()
        Rename-Item -Path $installDir -NewName $tempName
        try
        {
            Install-CMsi -Path $carbonTestInstallerActionsPath
            $installDir | Should -Not -Exist
        }
        finally
        {
            $tempDir = Split-Path -Path $installDir -Parent
            $tempDir = Join-Path -Path $tempDir -ChildPath $tempName
            Rename-Item -Path $tempDir -NewName (Split-Path -Path $installDir -Leaf)
        }
    }
}

Describe 'Install-Msi.when forcing product re-installation' {
    AfterEach { Reset }
    It 'should reinstall if forced to' -Skip:$isPwsh6 {
        Init
        Install-CMsi -Path $carbonTestInstallerActionsPath
        Assert-CarbonTestInstallerInstalled
    
        $installDir = Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath $testInstallerWithActions.Manufacturer
        $installDir = Join-Path -Path $installDir -ChildPath $testInstallerWithActions.ProductName
        $maxTries = 100
        $tryNum = 0
        do
        {
            if( (Test-Path -Path $installDir -PathType Container) )
            {
                break
            }
            Start-Sleep -Milliseconds 100
        }
        while( $tryNum++ -lt $maxTries )
    
        $installDir | Should -Exist
    
        $tryNum = 0
        do
        {
            Remove-Item -Path $installDir -Recurse -ErrorAction Ignore
            if( -not (Test-Path -Path $installDir -PathType Container) )
            {
                break
            }
            Start-Sleep -Milliseconds 100
        }
        while( $tryNum++ -lt $maxTries )
    
        $installDir | Should -Not -Exist
    
        Install-CMsi -Path $carbonTestInstallerActionsPath -Force
        $installDir | Should -Exist
    }
}

Describe 'Install-Msi.when there are spaces in the path to the MSI' {
    AfterEach { Reset }
    It 'should install msi with spaces in path' -Skip:$isPwsh6 {
        Init
        $newInstaller = Join-Path -Path $script:testRoot -ChildPath 'Installer With Spaces.msi'
        Copy-Item -Path $carbonTestInstallerPath -Destination $newInstaller
        Install-CMsi -Path $newInstaller
        Assert-CarbonTestInstallerInstalled
    }
}
