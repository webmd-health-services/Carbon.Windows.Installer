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
$result = $null

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
}

function ThenMsiObjectReturned
{
    param(
        [String[]] $WithTableName = @()
    )

    $script:result | Should -Not -BeNullOrEmpty
    foreach( $msi in $script:result )
    {
        $msi | Should -HaveCount 1
        $msi | Should -Not -BeNullOrEmpty
        $msi.pstypenames | Should -Contain 'Carbon.Windows.Installer.MsiInfo'
        $msi.Manufacturer | Should -Be 'Carbon'
        $msi.ProductName | Should -BeLike 'Carbon *'
        $msi.ProductCode | Should -Not -BeNullOrEmpty
        ([Guid]::Empty) | Should -Not -Be $msi.ProductCode
        $msi.ProductLanguage | Should -Be 1033
        $msi.ProductVersion | Should -Be '1.0.0'
        $msi.Property | Should -Not -BeNullOrEmpty
        $msi.Property.Count | Should -BeGreaterThan 5
        $msi.TableNames | Should -Not -BeNullOrEmpty
        $msi.TableNames | Should -Contain 'Condition'
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
        [String] $Named = 'CarbonTestInstaller.msi',

        [Object] $IncludingTable
    )

    $optionalParams = @{}
    if( $IncludingTable )
    {
        $optionalParams['IncludeTable'] = $IncludingTable
    }
    $script:result = Get-CMsi -Path (Join-Path -Path $msiRootPath -ChildPath $Named -Resolve) @optionalParams
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
        $msi = Get-ChildItem -Path $msiRootPath -Filter '*.msi' | Get-CMsi
        $msi | Should -Not -BeNullOrEmpty
        $msi | ForEach-Object {  Assert-CarbonMsi $_ }
    }

    It 'should accept array of strings' {
        $path = Join-Path -Path $msiRootPath -ChildPath 'CarbonTestInstaller.msi'

        $msi = Get-CMsi -Path @( $path, $path )
        ,$msi | Should -BeOfType ([object[]])
        foreach( $item in $msi )
        {
            Assert-CarbonMsi $item
        }
    }

    It 'should accept array of file info' {
        $path = Join-Path -Path $msiRootPath -ChildPath 'CarbonTestInstaller.msi'

        $item = Get-Item -Path $path
        $msi = Get-CMsi -Path @( $item, $item )

        ,$msi | Should -BeOfType ([object[]])
        foreach( $item in $msi )
        {
            Assert-CarbonMsi $item
        }
    }

    It 'should support wildcards' {
        $msi = Get-CMsi -Path (Join-Path -Path $msiRootPath -ChildPath '*.msi')
        ,$msi | Should -BeOfType ([object[]])
        foreach( $item in $msi )
        {
            Assert-CarbonMsi $item
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
        WhenGettingMsi -IncludingTable $tableNames[0]
        ThenMsiObjectReturned -WithTable $tableNames[0]
    }
}

Describe 'Get-Msi.when using wildcard' {
    It 'should return all tables that match the wildcard' {
        Init
        WhenGettingMsi -IncludingTable 'C*'
        ThenMsiObjectReturned -WithTableName ($tableNames | Where-Object { $_ -like 'C*' })
    }
}

Describe 'Get-Msi.when getting multiple tables' {
    It 'should return those tables' {
        Init
        $expectedTables = $tableNames | Select-Object -First 3
        WhenGettingMsi -IncludingTable $expectedTables
        ThenMsiObjectReturned -WithTableName $expectedTables
    }
}