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

#Requires -Version 5.1
Set-StrictMode -Version 'Latest'

& (Join-Path -Path $PSScriptRoot -ChildPath 'Initialize-Test.ps1' -Resolve)

$msiRootPath = Join-Path -Path $PSScriptRoot -ChildPath 'MSI' -Resolve
$testInstallerPath = Join-Path -Path $msiRootPath -ChildPath 'CarbonTestInstaller.msi' | Resolve-Path -Relative
$result = $null
$testRoot = $null
$testNum = 0
$wwwMsiUrl = 'https://the.earth.li/~sgtatham/putty/0.74/w64/putty-64bit-0.74-installer.msi'
$wwwMsiMembers = @{
    'Manufacturer' = 'Simon Tatham';
    'ProductName' = 'PuTTY release 0.74 (64-bit)';
    'ProductCode' = [Guid]::Parse('127b996b-5308-4012-865b-9446451ea326');
    'ProductLanguage' = 1033;
    'ProductVersion' = '0.74.0.0'
}

function Assert-CarbonMsi
{
    param(
        $msi
    )

    $script:result = $msi
    ThenMsiObjectReturned
}

$isPwsh6 = $PSVersionTable['PSVersion'].Major -eq 6
if( $isPwsh6 )
{
    return
}

function Init
{
    $script:result = $null
    Get-ChildItem -Path ([IO.Path]::GetTempPath()) -Filter '*.msi' | Remove-Item -Verbose -ErrorAction Ignore

    $script:testRoot = Join-Path -Path $TestDrive.FullName -ChildPath ($script:testNum++)
    New-Item -Path $script:testRoot -ItemType 'Directory' | Out-Null
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

Describe 'Get-Msi.when getting MSI with default tables' {
    It 'should get msi with Property and Feature tables' {
        Init
        WhenGettingMsi
        ThenMsiObjectReturned -WithTableName 'Property', 'Feature'
    }
}

Describe 'Get-Msi' {
    It 'should accept pipeline input' {
        $msi = Get-ChildItem -Path $msiRootPath -Filter '*.msi' | Get-TCMsi
        $msi | Should -Not -BeNullOrEmpty
        $msi | ForEach-Object { ThenMsiObjectReturned @{ 'Manufacturer' = 'Carbon' ; 'ProductVersion' = '1.0.0' ; 'ProductLanguage' = '1033' ;} }
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
            ThenMsiObjectReturned @{ 'Manufacturer' = 'Carbon' ; 'ProductVersion' = '1.0.0' ; 'ProductLanguage' = '1033' ;}            
        }
    }
}

# These tables aren't empty.
$tableNames = @(
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

Describe 'Get-Msi.when including a table' {
    It "should return the table" {
        Init
        WhenGettingMsi -WithParameter @{ 'IncludeTable' = $tableNames[0] }
        ThenMsiObjectReturned -WithTable $tableNames[0]
    }
}

Describe 'Get-Msi.when using wildcard' {
    It 'should return all tables that match the wildcard' {
        Init
        WhenGettingMsi -WithParameter @{ 'IncludeTable' = 'C*' }
        ThenMsiObjectReturned -WithTableName ($tableNames | Where-Object { $_ -like 'C*' })
    }
}

Describe 'Get-Msi.when getting multiple tables' {
    It 'should return those tables' {
        Init
        $expectedTables = $tableNames | Select-Object -First 3
        WhenGettingMsi -WithParameter @{ 'IncludeTable' = $expectedTables }
        ThenMsiObjectReturned -WithTableName $expectedTables
    }
}

Describe 'Get-Msi.when downloading MSI' {
    It 'should download file to temp directory' {
        Init
        WhenGettingMsi -WithParameter @{ 'Url' = $wwwMsiUrl }
        ThenMsiObjectReturned -WithMember $wwwMsiMembers
        $expectedFilePath = ([uri]$wwwMsiUrl).segments[-1]
        $expectedFilePath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath $expectedFilePath
        $expectedFilePath | Should -Exist
        Get-Item -Path $expectedFilePath | Should -HaveCount 1
        $script:result.Path | Should -Be $expectedFilePath
    }
}

Describe 'Get-Msi.when downloading to a directory' {
    It 'should save file MSI in that directory' {
        Init
        WhenGettingMsi -WithParameter @{ 'Url' = $wwwMsiUrl ; 'OutputPath' = $testRoot ; }
        ThenMsiObjectReturned -WithMember $wwwMsiMembers
        $expectedFilePath = ([uri]$wwwMsiUrl).segments[-1]
        $expectedFilePath = Join-Path -Path $testRoot -ChildPath $expectedFilePath
        $expectedFilePath | Should -Exist
        Get-Item -Path $expectedFilePath | Should -HaveCount 1
        $script:result.Path | Should -Be $expectedFilePath
    }
}

Describe 'Get-Msi.when downloading to file' {
    It 'should save MSI to that file' {
        Init
        $outputPath = Join-Path -Path $testRoot -ChildPath 'customfile.msi'
        WhenGettingMsi -WithParameter @{ 'Url' = $wwwMsiUrl ; 'OutputPath' = $outputPath }
        ThenMsiObjectReturned -WithMember $wwwMsiMembers
        $outputPath | Should -Exist
        Get-Item -Path $outputPath | Should -HaveCount 1
        $script:result.Path | Should -Be $outputPath
    }
}