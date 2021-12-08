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

function Assert-CarbonMsi
{
    param(
        $msi
    )

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
}

$isPwsh6 = $PSVersionTable['PSVersion'].Major -eq 6

Describe 'Get-Msi' {
    It 'should get msi' -Skip:$isPwsh6 {
        $msi = Get-CMsi -Path (Join-Path -Path $msiRootPath -ChildPath 'CarbonTestInstaller.msi' -Resolve)
        Assert-CarbonMsi $msi
    }

    It 'should accept pipeline input' -Skip:$isPwsh6 {
        $msi = Get-ChildItem -Path $msiRootPath -Filter '*.msi' | Get-CMsi
        $msi | Should -Not -BeNullOrEmpty
        $msi | ForEach-Object {  Assert-CarbonMsi $_ }
    }

    It 'should accept array of strings' -Skip:$isPwsh6 {
        $path = Join-Path -Path $msiRootPath -ChildPath 'CarbonTestInstaller.msi'

        $msi = Get-CMsi -Path @( $path, $path )
        ,$msi | Should -BeOfType ([object[]])
        foreach( $item in $msi )
        {
            Assert-CarbonMsi $item
        }
    }

    It 'should accept array of file info' -Skip:$isPwsh6 {
        $path = Join-Path -Path $msiRootPath -ChildPath 'CarbonTestInstaller.msi'

        $item = Get-Item -Path $path
        $msi = Get-CMsi -Path @( $item, $item )

        ,$msi | Should -BeOfType ([object[]])
        foreach( $item in $msi )
        {
            Assert-CarbonMsi $item
        }
    }

    It 'should support wildcards' -Skip:$isPwsh6 {
        $msi = Get-CMsi -Path (Join-Path -Path $msiRootPath -ChildPath '*.msi')
        ,$msi | Should -BeOfType ([object[]])
        foreach( $item in $msi )
        {
            Assert-CarbonMsi $item
        }
    }
}
