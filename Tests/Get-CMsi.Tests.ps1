
#Requires -Version 5.1
Set-StrictMode -Version 'Latest'

& (Join-Path -Path $PSScriptRoot -ChildPath 'Initialize-Test.ps1' -Resolve)

$isPwsh6 = $PSVersionTable['PSVersion'].Major -eq 6
if( $isPwsh6 )
{
    return
}

BeforeAll {
    $script:msiRootPath = Join-Path -Path $PSScriptRoot -ChildPath 'MSI' -Resolve
    $script:testInstallerPath = Join-Path -Path $msiRootPath -ChildPath 'CarbonTestInstaller.msi' | Resolve-Path -Relative
    $script:result = $null
    $script:testRoot = $null
    $script:testNum = 0
    $script:wwwMsiUrl = 'https://the.earth.li/~sgtatham/putty/0.74/w64/putty-64bit-0.74-installer.msi'
    $script:wwwMsiMembers = @{
        'Manufacturer' = 'Simon Tatham';
        'ProductName' = 'PuTTY release 0.74 (64-bit)';
        'ProductCode' = [Guid]::Parse('127b996b-5308-4012-865b-9446451ea326');
        'ProductLanguage' = 1033;
        'ProductVersion' = '0.74.0.0'
    }

    # Tables in test installer.
    $script:tableNames = @(
        '_Validation',
        'AdminExecuteSequence',
        'AdminUISequence',
        'AdvtExecuteSequence',
        'Property',
        'Feature',
        'Component',
        'Directory',
        'Control',
        'Dialog',
        'ControlCondition',
        'ControlEvent',
        'CustomAction',
        'EventMapping',
        'FeatureComponents',
        'InstallExecuteSequence',
        'InstallUISequence',
        'Media',
        'ModuleSignature',
        'RadioButton',
        'TextStyle',
        'UIText',
        'Upgrade'
    )

    function Assert-CarbonMsi
    {
        param(
            $msi
        )

        $script:result = $msi
        ThenMsiObjectReturned
    }

    function ThenMsiObjectReturned
    {
        param(
            [String[]] $WithTableName = @(),

            [hashtable] $WithMember
        )

        $script:result | Should -Not -BeNullOrEmpty
        if( -not $WithMember )
        {
            $WithMember = @{
                'Manufacturer' = 'Carbon';
                'ProductName' = 'Carbon Test Installer';
                'ProductCode' = [Guid]::Parse('e1724abc-a8d6-4d88-bbed-2e077c9ae6d2');
                'ProductLanguage' = 1033;
                'ProductVersion' = '1.0.0'
            }
        }

        foreach( $msi in $script:result )
        {
            $msi | Should -HaveCount 1
            $msi | Should -Not -BeNullOrEmpty
            $msi.pstypenames | Should -Contain 'Carbon.Windows.Installer.MsiInfo'

            foreach( $propertyName in $WithMember.Keys )
            {
                $msi |
                    Get-Member -Name $propertyName |
                    Should -Not -BeNullOrEmpty -Because "MSI should have ""$($propertyName)"" property"
                $msi.$propertyName | Should -Be $WithMember[$propertyName]
            }

            $msi.Property | Should -Not -BeNullOrEmpty
            $msi.Property.Count | Should -BeGreaterThan 5
            $msi.TableNames | Should -Not -BeNullOrEmpty
            # $msi.TableNames | Should -Contain 'Condition'
            $msi.TableNames | Should -Contain 'Property'

            $WithTableName = & {
                    'Property'
                    'Feature'
                    $WithTableName | Write-Output
                } |
                Select-Object -Unique
            foreach( $tableName in $msi.TableNames )
            {
                $msi.Tables | Get-Member -Name $tableName | Should -Not -BeNullOrEmpty
                $table = $msi.Tables.$tableName
                if( $WithTableName -contains $tableName )
                {
                    $table| Should -Not -BeNullOrEmpty -Because "should return contents of table ""$($tableName)"""
                    $table |
                        ForEach-Object { $_.pstypenames[0] | Should -Be "Carbon.Windows.Installer.Records.$($tableName)" }
                }
                else
                {
                    $table | Should -BeNullOrEmpty
                }
            }
        }
    }

    function WhenGettingMsi
    {
        param(
            [hashtable] $WithParameter = @{}
        )

        if( -not $WithParameter.ContainsKey('Path') -and -not $WithParameter.ContainsKey('Url') )
        {
            $WithParameter['Path'] = $testInstallerPath;
        }

        $script:result = Get-TCMsi @WithParameter
    }
}

Describe 'Get-CMsi' {
    BeforeEach {
        $Global:Error.Clear()

        $script:result = $null
        Get-ChildItem -Path ([IO.Path]::GetTempPath()) -Filter '*.msi' | Remove-Item -ErrorAction Ignore

        $script:testRoot = Join-Path -Path $TestDrive -ChildPath $script:testNum
        $script:testNum += 1
        Write-Information "BeforeEach"
        Write-Information "    $($TestDrive)"
        New-Item -Path $script:testRoot -ItemType 'Directory' | Out-Null
    }

    It 'should always return Property and Feature tables' {
        WhenGettingMsi
        ThenMsiObjectReturned -WithTableName 'Property', 'Feature'
    }

    It 'should accept pipeline input' {
        $msi = Get-ChildItem -Path $script:msiRootPath -Filter '*.msi' | Get-TCMsi
        $msi | Should -Not -BeNullOrEmpty
        $msi | ForEach-Object {
            $script:result = $_
            ThenMsiObjectReturned -WithMember @{
                'Manufacturer' = 'Carbon';
                'ProductVersion' = '1.0.0';
                'ProductLanguage' = '1033';
            }
        }
    }

    It 'should accept array of strings' {
        $path = Join-Path -Path $msiRootPath -ChildPath 'CarbonTestInstaller.msi'

        $msi = Get-TCMsi -Path @( $path, $path )
        ,$msi | Should -BeOfType ([object[]])
        foreach( $item in $msi )
        {
            Assert-CarbonMsi $item
        }
    }

    It 'should accept array of file info' {
        $path = Join-Path -Path $msiRootPath -ChildPath 'CarbonTestInstaller.msi'

        $item = Get-Item -Path $path
        $msi = Get-TCMsi -Path @( $item, $item )

        ,$msi | Should -BeOfType ([object[]])
        foreach( $item in $msi )
        {
            Assert-CarbonMsi $item
        }
    }

    It 'should support wildcards' {
        $msi = Get-TCMsi -Path (Join-Path -Path $msiRootPath -ChildPath '*.msi')
        ,$msi | Should -BeOfType ([object[]])
        foreach( $item in $msi )
        {
            $script:result = $item
            ThenMsiObjectReturned -WithMember @{
                'Manufacturer' = 'Carbon';
                'ProductVersion' = '1.0.0';
                'ProductLanguage' = '1033';
            }
        }
    }

    It "should return a single included table" {
        WhenGettingMsi -WithParameter @{ 'IncludeTable' = $tableNames[0] }
        ThenMsiObjectReturned -WithTable $tableNames[0]
    }

    It 'should support wildcard matching on table names' {
        WhenGettingMsi -WithParameter @{ 'IncludeTable' = 'C*' }
        ThenMsiObjectReturned -WithTableName ($tableNames | Where-Object { $_ -like 'C*' })
    }

    It 'should including multiple tables' {
        $expectedTables = $tableNames | Select-Object -First 3
        WhenGettingMsi -WithParameter @{ 'IncludeTable' = $expectedTables }
        ThenMsiObjectReturned -WithTableName $expectedTables
    }

    It 'should download MSI files to temp directory' {
        WhenGettingMsi -WithParameter @{ 'Url' = $wwwMsiUrl }
        ThenMsiObjectReturned -WithMember $wwwMsiMembers
        $expectedFilePath = ([uri]$wwwMsiUrl).segments[-1]
        $expectedFilePath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath $expectedFilePath
        $expectedFilePath | Should -Exist
        Get-Item -Path $expectedFilePath | Should -HaveCount 1
        $script:result.Path | Should -Be $expectedFilePath
    }

    It 'should download MSI to custom directory' {
        WhenGettingMsi -WithParameter @{ 'Url' = $wwwMsiUrl ; 'OutputPath' = $testRoot ; }
        ThenMsiObjectReturned -WithMember $wwwMsiMembers
        $expectedFilePath = ([uri]$wwwMsiUrl).segments[-1]
        $expectedFilePath = Join-Path -Path $testRoot -ChildPath $expectedFilePath
        $expectedFilePath | Should -Exist
        Get-Item -Path $expectedFilePath | Should -HaveCount 1
        $script:result.Path | Should -Be $expectedFilePath
    }

    It 'should download MSI to custom file' {
        $outputPath = Join-Path -Path $testRoot -ChildPath 'customfile.msi'
        WhenGettingMsi -WithParameter @{ 'Url' = $wwwMsiUrl ; 'OutputPath' = $outputPath }
        ThenMsiObjectReturned -WithMember $wwwMsiMembers
        $outputPath | Should -Exist
        Get-Item -Path $outputPath | Should -HaveCount 1
        $script:result.Path | Should -Be $outputPath
        $script:result = $null
    }
}