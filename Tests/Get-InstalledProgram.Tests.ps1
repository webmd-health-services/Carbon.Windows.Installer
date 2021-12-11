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

#Requires -Version 4
Set-StrictMode -Version 'Latest'

& (Join-Path -Path $PSScriptRoot -ChildPath 'Initialize-Test.ps1' -Resolve)

Describe 'Get-InstalledProgram' {
    BeforeEach {
        $Global:Error.Clear()
    }

    It 'should get all installed programs' {
       $programs = Get-TCInstalledProgram
       $programs | Should -Not -BeNullOrEmpty
    }

    It ('should get information about each program') {
        $programs = Get-TCInstalledProgram
        foreach( $program in $programs )
        {
            Write-Verbose -Message $program.DisplayName
            $program | Should -Not -BeNullOrEmpty
            [Microsoft.Win32.RegistryKey]$key = $program.Key
            
            foreach( $property in (Get-Member -InputObject $program -MemberType Property) )
            {
                $propertyName = $property.Name
                Write-Verbose -Message ('  {0}' -f $propertyName)
                if( $propertyName -eq 'Version' )
                {
                    Write-Verbose 'BREAK'
                }
    
                if( $propertyName -eq 'Key' )
                {
                    continue
                }
    
                $keyValue = $key.GetValue( $propertyName )
                $propertyValue = $program.$propertyName
    
                if( $propertyName -eq 'ProductCode' )
                {
                    $propertyValue = Split-Path -Leaf -Path $key.Name
                    [Guid]$guid = [Guid]::Empty
                    [Guid]::TryParse( $propertyValue, [ref]$guid )
                    $propertyValue = $guid
                    $keyValue = $guid
                }
                elseif( $propertyName -eq 'User' )
                {
                    if( $key.Name -match 'HKEY_USERS\\([^\\]+)\\' )
                    {
                        $sddl = $Matches[1]
                        $sid = New-Object 'Security.Principal.SecurityIdentifier' $sddl
                        try
                        {
                            $propertyValue = $sid.Translate([Security.Principal.NTAccount]).Value
                        }
                        catch
                        {
                            $propertyValue = $sid.ToString()
                        }
                        $keyValue = $propertyValue
                    }
                }
    
                $typeName = $program.GetType().GetProperty($propertyName).PropertyType.Name
                if( $keyValue -eq $null )
                {
                    if( $typeName -eq 'Int32' )
                    {
                        $keyValue = 0
                    }
                    elseif( $typeName -eq 'Version' )
                    {
                        $keyValue = $null
                    }
                    elseif( $typeName -eq 'DateTime' )
                    {
                        $keyValue = [DateTime]::MinValue
                    }
                    elseif( $typeName -eq 'Boolean' )
                    {
                        $keyValue = $false
                    }
                    elseif( $typeName -eq 'Guid' )
                    {
                        $keyValue = [Guid]::Empty
                    }
                    else
                    {
                        $keyValue = ''
                    }
                }
                else
                {
                    if( $typeName -eq 'DateTime' )
                    {
                        [DateTime]$dateTime = [DateTime]::MinValue
    
                        if( -not ([DateTime]::TryParse($keyValue,[ref]$dateTime)) )
                        {
                            [DateTime]::TryParseExact( $keyValue, 'yyyyMMdd', [Globalization.CultureInfo]::CurrentCulture, [Globalization.DateTimeStyles]::None, [ref]$dateTime)
                        }
                        $keyValue = $dateTime
                    }
                    elseif( $typeName -eq 'Int32' )
                    {
                        $intValue = 0
                        $keyValue = [Int32]::TryParse($keyValue, [ref] $intValue)
                        $keyValue = $intValue
                    }
                    elseif( $typeName -eq 'Version' )
                    {
                        [int]$intValue = 0
                        if( $keyValue -isnot [int32] -and [int]::TryParse($keyValue,[ref]$intValue) )
                        {
                            $keyValue = $intValue
                        }

                        if( $keyValue -is [int32] )
                        {
                            $major = $keyValue -shr 24   # First 8 bits
                            $minor = ($keyValue -band 0x00ff0000) -shr 16  # bits 9 - 16
                            $build = $keyValue -band 0x0000ffff   # last 8 bits
                            $keyValue = New-Object 'Version' $major,$minor,$build
                        }
                        else
                        {
                            [Version]$version = $null
                            if( [Version]::TryParse($keyValue, [ref]$version) )
                            {
                                $keyValue = $version
                            }
                        }
                    }
                }
    
                if( $keyValue -eq $null )
                {
                    $propertyValue | Should -BeNullOrEmpty
                }
                else
                {
                    $propertyValue | Should -Be $keyValue
                }
            }
        }
    }

    It 'should get a program by its name' {
        $p = Get-TCInstalledProgram | Select-Object -First 1
        $p2 = Get-TCInstalledProgram $p.DisplayName
        $p2 | Should -Not -BeNullOrEmpty
        Compare-Object -ReferenceObject $p -DifferenceObject $p2 | Should -BeNullOrEmpty
    }

    It 'should fail when a program does not exist' {
        Get-TCInstalledProgram -Name 'CwiFubarSnafu' -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
        $Global:Error | Should -Match '"CwiFubarSnafu" is not installed'
    }

    It 'should not fail when a program does not exist and ignoring errors' {
        Get-TCInstalledProgram -Name 'CwiFubarSnafu' -ErrorAction Ignore | Should -BeNullOrEmpty
        $Global:Error | Should -BeNullOrEmpty
    }

    It 'should find programs with wildcards' {
        $p = Get-TCInstalledProgram | Select-Object -First 1

        $wildcard = $p.DisplayName.Substring(0,$p.DisplayName.Length - 1)
        $wildcard = '{0}*' -f $wildcard
        $p2 = Get-TCInstalledProgram $wildcard
    
            $p2 | Should -Not -BeNullOrEmpty
        Compare-Object -ReferenceObject $p -DifferenceObject $p2 | Should -BeNullOrEmpty
    }

    It 'should not fail when wildcard does not match any programs' {
        Get-TCInstalledProgram -Name 'CwiFubarSnafu*' | Should -BeNullOrEmpty
        $Global:Error | Should -BeNullOrEmpty
    }

    It 'should ignore invalid integer version' {
        
        $program = Get-TCInstalledProgram | Select-Object -First 1

        $regKeyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Carbon.Windows.Installer'
        if( -not (Test-Path -Path $regKeyPath ) )
        {
            New-Item -Path $regKeyPath -ItemType RegistryKey -Force
        }

        try
        {
            $programName = 'Carbon+Get-CInstalledProgram'
            New-ItemProperty -Path $regKeyPath -Name 'DisplayName' -Value $programName -PropertyType 'String'
            New-ItemProperty -Path $regKeyPath -Name 'Version' -Value 0xff000000 -PropertyType 'DWord'

            $program = Get-TCInstalledProgram -Name $programName
        
            $program | Should -Not -BeNullOrEmpty
            $program.Version | Should -BeNullOrEmpty
        }
        finally
        {
            Remove-Item -Path $regKeyPath -Recurse
        }
    }
}
