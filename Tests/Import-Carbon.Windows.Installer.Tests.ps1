
#Requires -Version 5.1
Set-StrictMode -Version 'Latest'

& (Join-Path -Path $PSScriptRoot -ChildPath 'Initialize-Test.ps1' -Resolve)

BeforeAll {
    function GivenModuleLoaded
    {
        Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath '..\Carbon.Windows.Installer\Carbon.Windows.Installer.psd1' -Resolve)
        Get-Module -Name 'Carbon.Windows.Installer' | Add-Member -MemberType NoteProperty -Name 'NotReloaded' -Value $true
    }

    function GivenModuleNotLoaded
    {
        Remove-Module -Name 'Carbon.Windows.Installer' -Force -ErrorAction Ignore
    }

    function ThenModuleLoaded
    {
        $module = Get-Module -Name 'Carbon.Windows.Installer'
        $module | Should -Not -BeNullOrEmpty
        $module | Get-Member -Name 'NotReloaded' | Should -BeNullOrEmpty
    }

    function WhenImporting
    {
        $script:importedAt = Get-Date
        Start-Sleep -Milliseconds 1
        & (Join-Path -Path $PSScriptRoot -ChildPath '..\Carbon.Windows.Installer\Import-Carbon.Windows.Installer.ps1' -Resolve)
    }
}

AfterAll {
    Remove-Module -Name 'Carbon.Windows.Installer' -Force
}

Describe 'Import-Carbon.Windows.Installer.when module not loaded' {
    It 'should import the module when not loaded' {
        GivenModuleNotLoaded
        WhenImporting
        ThenModuleLoaded
    }

    It 'should re-import the module when loaded' {
        GivenModuleLoaded
        WhenImporting
        ThenModuleLoaded
    }
}
