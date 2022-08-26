
#Requires -Version 5.1
Set-StrictMode -Version 'Latest'

$isPwsh6 = $PSVersionTable['PSVersion'].Major -eq 6
if( $isPwsh6 )
{
    return
}

BeforeAll {
    & (Join-Path -Path $PSScriptRoot -ChildPath 'Initialize-Test.ps1' -Resolve)

    $script:msiRootPath = Join-Path -Path $PSScriptRoot -ChildPath 'MSI' -Resolve

    $script:carbonTestInstallerPath = Join-Path -Path $msiRootPath -ChildPath 'CarbonTestInstaller.msi' -Resolve
    $script:testInstaller = Get-TCMsi -Path $carbonTestInstallerPath

    $script:carbonTestInstallerActionsPath =
            Join-Path -Path $msiRootPath -ChildPath 'CarbonTestInstallerWithCustomActions.msi' -Resolve
    $script:testInstallerWithActions = Get-TCMsi -Path $script:carbonTestInstallerActionsPath

    $script:testRoot = $null
    $script:testNum = 0

    function Assert-CarbonTestInstallerInstalled
    {
        [CmdletBinding()]
        param(
            [Object] $Msi = $script:testInstaller
        )

        Test-Installed -Msi $Msi | Should -BeTrue

    }

    function Assert-CarbonTestInstallerNotInstalled
    {
        Test-Installed | Should -BeFalse
    }

    function GivenInstalled
    {
        Install-TCMsi -Path $carbonTestInstallerPath
        Assert-CarbonTestInstallerInstalled
    }

    function Test-Installed
    {
        [CmdletBinding()]
        param(
            [Object] $Msi = $script:testInstaller
        )

        $installerRegPath =
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{$($Msi.ProductCode)}"
        $installed = (Test-Path -Path $installerRegPath)
        $installedMsg = 'not '
        if( $installed )
        {
            $installedMsg = '    '
        }
        Write-Debug "$($installedMsg)installed  $($Msi.ProductName)"
        return $installed
    }

    function ThenInstallerError
    {
        [CmdletBinding()]
        param(
            [String] $WithErrorMatching
        )

        $Global:Error | Should -Not -BeNullOrEmpty
        if( $WithErrorMatching )
        {
            $Global:Error | Should -Match $WithErrorMatching
        }
    }

    function ThenInstallerSucceeded
    {
        $Global:Error | Should -BeNullOrEmpty
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
        @( $script:testInstaller, $script:testInstallerWithActions ) |
            Where-Object { Test-Installed -Msi $_ } |
            ForEach-Object {
                #msiexec /fa $_.Path /quiet /l*vx 'D:\restore.log'
                $msiProcess = Start-Process -FilePath "msiexec.exe" -ArgumentList "/quiet","/fa",('"{0}"' -f $_.Path) -NoNewWindow -Wait -PassThru
                if( $null -ne $msiProcess.ExitCode -and $msiProcess.ExitCode -notin @(0, 1605) )
                {
                    Write-Error ("{0} {1} repair failed. (Exit code: {2}; MSI: {3})" -f $_.ProductName,$_.ProductVersion,$msiProcess.ExitCode,$_.Path)
                }
                #msiexec /uninstall $_.Path /quiet /l*vx 'D:\uninstall.log'
                $logPath = Join-Path -Path $PSScriptRoot -ChildPath "..\.output\$($_.ProductName).uninstall.log"
                $msiProcess = Start-Process -FilePath "msiexec.exe" `
                                            -ArgumentList "/quiet","/uninstall",('"{0}"' -f $_.Path),'/l!*vx',"""$($logPath)""" `
                                            -NoNewWindow `
                                            -Wait `
                                            -PassThru
                if( $null -ne $msiProcess.ExitCode -and $msiProcess.ExitCode -notin @(0, 1605))
                {
                    Write-Error ("{0} {1} uninstall failed. (Exit code: {2}; MSI: {3})" -f $_.ProductName,$_.ProductVersion,$msiProcess.ExitCode,$_.Path)
                }
            }

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

            [hashtable] $WithParameter = @{}
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
                    Copy-Item -Path $installerPath -Destination $OutFile
            }.GetNewClosure()

            if( -not $WithExpectedChecksum )
            {
                $WithExpectedChecksum = (Get-FileHash -Path $carbonTestInstallerPath).Hash
            }
            $output = Install-TCMsi -Url $AtUrl `
                                -Checksum $WithExpectedChecksum `
                                -ProductName $testInstaller.ProductName `
                                -ProductCode $testInstaller.ProductCode `
                                @WithParameter
            $output | Should -BeNullOrEmpty
            return
        }

        $output = Install-TCMsi -Path $script:carbonTestInstallerPath @WithParameter
        $output | Should -BeNullOrEmpty
    }
}

Describe 'Install-CMsi' {
    BeforeEach {
        Uninstall-CarbonTestInstaller

        $script:testRoot = Join-Path -Path $TestDrive -ChildPath ($script:testNum++)
        New-Item -Path $script:testRoot -ItemType 'Directory' | Out-Null

        $Global:Error.Clear()
    }

    AfterEach {
        Uninstall-CarbonTestInstaller
    }

    It 'should fail when file is not an MSI' {
        Install-TCMsi -Path $PSCommandPath -ErrorAction SilentlyContinue
        $Global:Error.Count | Should -BeGreaterThan 0
        $Global:Error[0] | Should -Match 'Exception opening MSI database'
    }

    It 'should support what if' {
        Assert-CarbonTestInstallerNotInstalled
        Install-TCMsi -Path $carbonTestInstallerPath -WhatIf
        $Global:Error.Count | Should -Be 0
        Assert-CarbonTestInstallerNotInstalled
    }

    It 'should install an MSI' {
        Assert-CarbonTestInstallerNotInstalled
        Install-TCMsi -Path $carbonTestInstallerPath
        Assert-CarbonTestInstallerInstalled
        ThenInstallerSucceeded
    }

    It 'should handle when an installer fails' {
        $logFilePath = $carbonTestInstallerActionsPath | Split-Path -Leaf
        $logFilePath = "$($logFilePath).*.*.log"
        $logFilePath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath $logFilePath
        Get-Item -Path $logFilePath -ErrorAction Ignore | Remove-Item
        $envVarName = 'CARBON_TEST_INSTALLER_THROW_INSTALL_EXCEPTION'
        [Environment]::SetEnvironmentVariable($envVarName, $true.ToString(), 'User')
        try
        {
            Install-TCMsi -Path $carbonTestInstallerActionsPath -ErrorAction SilentlyContinue
            ThenInstallerError 'returned exit code 1603'
            Assert-CarbonTestInstallerNotInstalled $script:testInstallerWithActions
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

    It 'should support wildcards in path to installer' {
        Copy-Item $carbonTestInstallerPath -Destination (Join-Path -Path $script:testRoot -ChildPath 'One.msi')
        Copy-Item $carbonTestInstallerPath -Destination (Join-Path -Path $script:testRoot -ChildPath 'Two.msi')
        Install-TCMsi -Path (Join-Path -Path $script:testRoot -ChildPath '*.msi')
        Assert-CarbonTestInstallerInstalled
        ThenInstallerSucceeded
    }

    It 'should not reinstall an MSI when the MSI is already installed' {
        GivenInstalled
        WhenInstalling -MockInstall
        ThenMsi -Not -Reinstalled
    }

    It 'should re-install an MSI when using the force' {
        GivenInstalled
        WhenInstalling -WithParameter @{ 'Force' = $true } -MockInstall
        ThenMsi -Reinstalled
        ThenInstallerSucceeded
    }

    It 'should handle spaces in MSI path' {
        $newInstaller = Join-Path -Path $script:testRoot -ChildPath 'Installer With Spaces.msi'
        Copy-Item -Path $carbonTestInstallerPath -Destination $newInstaller
        Install-TCMsi -Path $newInstaller
        Assert-CarbonTestInstallerInstalled
        ThenInstallerSucceeded
    }

    It 'should download and install an MSI' {
        WhenInstalling -FromWeb
        ThenMsi -Downloaded -Installed
        ThenMsi -Not -InTempDirectory
        ThenInstallerSucceeded
    }

    It 'should not download or install an MSI if it is already installed' {
        GivenInstalled
        WhenInstalling -FromWeb -MockInstall
        ThenMsi -Not -Downloaded
        ThenMsi -Not -Reinstalled
    }

    It 'should download and install the MSI for an already installed program if using the force' {
        GivenInstalled
        WhenInstalling -FromWeb -MockInstall -WithParameter @{ 'Force' = $true }
        ThenMsi -Downloaded
        ThenMsi -Reinstalled
        ThenInstallerSucceeded
    }

    It 'should download but not install the program when its checksum does not match' {
        WhenInstalling -FromWeb -WithExpectedChecksum 'deadbee' -ErrorAction SilentlyContinue
        ThenMsi -Downloaded
        ThenMsi -Not -Installed
        $Global:Error | Should -Match 'does not match expected checksum'
    }

    It 'should allow customizing the installer display mode' {
        WhenInstalling -WithParameter @{ 'DisplayMode' = 'Passive' }
        ThenMsi -Installed
        ThenInstallerSucceeded
    }

    It 'should allow customizing logging options and log file' {
        $logPath = Join-Path -Path $testRoot -ChildPath 'installer.log'
        WhenInstalling -WithParameter @{ 'LogOption' = '!p' ; 'LogPath' = $logPath ; }
        ThenMsi -Installed
        ThenInstallerSucceeded
        $logPath | Should -Exist
        $logPath | Should -FileContentMatch -ExpectedContent 'Property\(S\)'
        $logPath | Should -Not -FileContentMatch -ExpectedContent 'msiexec.exe'
        $logPath | Should -Not -FileContentMatch -ExpectedContent 'MSI \(s\)'
        $logPath | Should -Not -FileContentMatch -ExpectedContent '^DEBUG  :'
    }

    It 'should allow passing extra arguments to msiexec' {
        WhenInstalling -WithParameter @{ 'ArgumentList' = @('ONE="1"', 'TWO="2"') } -MockInstall
        Assert-MockCalled -CommandName 'Start-Process' `
                          -ModuleName 'Carbon.Windows.Installer' `
                          -ParameterFilter {
                              Write-Debug 'ArgumentList'
                              $ArgumentList | Write-Debug
                              $ArgumentList | Should -Contain 'ONE="1"'
                              $ArgumentList | Should -Contain 'TWO="2"'
                              return $true
                          }
    }
}
