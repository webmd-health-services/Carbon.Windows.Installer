
function Get-InstalledProgram
{
    <#
    .SYNOPSIS
    Gets information about the programs installed on the computer.
    
    .DESCRIPTION
    The `Get-CInstalledProgram` function is the PowerShell equivalent of the Programs and Features/Apps and Features
    settings UI. It inspects the registry to determine what programs are installed. When running as an administrator, it
    returns programs installed for *all* users, not just the current user.

    The function looks in the following registry keys for install information:

    * HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall
    * HKEY_LOCAL_MACHINE\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
    * HKEY_USERS\*\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'

    A key is skipped if:

    * it doesn't have a `DisplayName` value.
    * it has a `ParentKeyName` value.
    * it has a `SystemComponent` value and its value is `1`.

    `Get-CInstalledProgram` tries its best to get accurate data. The following properties either aren't stored
    consistently, is in strange formats, can't be parsed, etc.

    * The `ProductCode` property is set to `[Guid]::Empty` if the software doesn't have a product code.
    * The `User` property will only be set for software installed for specific users. For global software, the `User`
      property will be `[String]::Empty`.
    * The `InstallDate` property is set to `[DateTime]::MinValue` if the install date can't be determined.
    * The `Version` property is `$null` if the version can't be parsed.
    
    .EXAMPLE
    Get-CInstalledProgram | Sort-Object 'DisplayName'

    Demonstrates how to get a list of all the installed programs, similar to what the Programs and Features settings UI
    shows. The returned objects are not sorted, so you'll usually want to pipe the output to `Sort-Object`.

    .EXAMPLE
    Get-CInstalledProgram -Name 'Google Chrome'

    Demonstrates how to get a specific program. If the specific program isn't found, `$null` is returned.

    .EXAMPLE
    Get-CInstalledProgram -Name 'Microsoft*'

    Demonstrates that you can use wildcards to search for programs.
    #>
    [CmdletBinding()]
    param(
        # The name of a specific program to get. Wildcards supported.
        [String] $Name
    )

    Set-StrictMode -Version 'Latest'
    Use-CallerPreference -Cmdlet $PSCmdlet -Session $ExecutionContext.SessionState

    function Get-KeyStringValue
    {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [Microsoft.Win32.RegistryKey] $Key,

            [Parameter(Mandatory)]
            [String] $ValueName
        )

        $value = $Key.GetValue($ValueName)
        if( $null -eq $value )
        {
            return ''
        }
        return $value.ToString()
    }

    function Get-KeyIntValue
    {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [Microsoft.Win32.RegistryKey] $Key,

            [Parameter(Mandatory)]
            [String] $ValueName
        )

        [int] $value = 0
        $rawValue = $Key.GetValue($ValueName)
        if( [int]::TryParse([Convert]::ToString($rawValue), [ref]$value) )
        {
            return $value
        }

        return 0
    }

    if( -not (Test-Path -Path 'hku:\') )
    {
        $null = New-PSDrive -Name 'HKU' -PSProvider Registry -Root 'HKEY_USERS' -WhatIf:$false
    }

    $keys = & {
        Get-ChildItem -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
        Get-ChildItem -Path 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
        Get-ChildItem -Path 'hku:\*\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction Ignore
    }

    $programs = $null
    & {
        foreach( $key in $keys )
        {
            $valueNames = [Collections.Generic.Hashset[String]]::New($key.GetValueNames())

            if( -not $valueNames.Contains('DisplayName') )
            {
                continue
            }

            $displayName = $key.GetValue('DisplayName')
            if( $Name -and $displayName -notlike $Name )
            {
                continue
            }

            if( $valueNames.Contains('ParentKeyName') )
            {
                continue
            }

            if( $valueNames.Contains('SystemComponent') )
            {
                continue
            }

            $systemComponent = $key.GetValue('SystemComponent')
            if( $systemComponent -eq 1 )
            {
                continue
            }

            $info = [pscustomobject]@{
                Comments = Get-KeyStringValue -Key $key -ValueName 'Comments';
                Contact = Get-KeyStringValue -Key $key -ValueName 'Contact';
                DisplayName = $displayName;
                DisplayVersion = Get-KeyStringValue -Key $key -ValueName 'DisplayVersion';
                EstimatedSize = Get-KeyIntValue -Key $key -ValueName 'EstimatedSize';
                HelpLink = Get-KeyStringValue -Key $key -ValueName 'HelpLink';
                HelpTelephone = Get-KeyStringValue -Key $key -ValueName 'HelpTelephone';
                InstallDate = $null;
                InstallLocation = Get-KeyStringValue -Key $key -ValueName 'InstallLocation';
                InstallSource = Get-KeyStringValue -Key $key -ValueName 'InstallSource';
                Key = $key;
                Language = Get-KeyIntValue -Key $key -ValueName 'Language';
                ModifyPath = Get-KeyStringValue -Key $key -ValueName 'ModifyPath';
                Path = Get-KeyStringValue -Key $key -ValueName 'Path';
                ProductCode = $null;
                Publisher = Get-KeyStringValue -Key $key -ValueName 'Publisher';
                Readme = Get-KeyStringValue -Key $key -ValueName 'Readme';
                Size = Get-KeyStringValue -Key $key -ValueName 'Size';
                UninstallString = Get-KeyStringValue -Key $key -ValueName 'UninstallString';
                UrlInfoAbout = Get-KeyStringValue -Key $key -ValueName 'URLInfoAbout';
                UrlUpdateInfo = Get-KeyStringValue -Key $key -ValueName 'URLUpdateInfo';
                User = $null;
                Version = $null;
                VersionMajor = Get-KeyIntValue -Key $key -ValueName 'VersionMajor';
                VersionMinor = Get-KeyIntValue -Key $key -ValueName 'VersionMinor';
                WindowsInstaller = $false;
            }
            $info | Add-Member -Name 'Name' -MemberType AliasProperty -Value 'DisplayName'
            
            $installDateValue = Get-KeyStringValue -Key $key -ValueName 'InstallDate'
            [DateTime] $installDate = [DateTime]::MinValue
            if( [DateTime]::TryParse($installDateValue, [ref]$installDate) -or
                [DateTime]::TryParseExact($installDateValue, 'yyyyMMdd', [cultureinfo]::CurrentCulture,
                                            [Globalization.DateTimeStyles]::None, [ref]$installDate)
            )
            {
                $info.InstallDate = $installDate
            }

            [Guid]$productCode = [Guid]::Empty
            $keyName = [IO.Path]::GetFileName($key.Name)
            if( [Guid]::TryParse($keyName, [ref]$productCode) )
            {
                $info.ProductCode = $productCode
            }

            if( $key.Name -match '^HKEY_USERS\\([^\\]+)\\')
            {
                $info.User = $Matches[1]
                $numErrors = $Global:Error.Count
                try
                {
                    $sid = [Security.Principal.SecurityIdentifier]::New($user)
                    if( $sid.IsValidTargetType([Security.Principal.NTAccount]))
                    {
                        $ntAccount = $sid.Translate([Security.Principal.NTAccount])
                        if( $ntAccount )
                        {
                            $info.User = $ntAccount.Value
                        }
                    }
                }
                catch
                {
                    for( $idx = $numErrors; $idx -lt $Global:Error.Count; ++$idx )
                    {
                        $Global:Error.RemoveAt(0)
                    }
                }
            }

            $intVersion = Get-KeyIntValue -Key $key -ValueName 'Version'
            if( $intVersion )
            {
                $major = $intVersion -shr 24 # first 8 bits are major version number
                $minor = ($intVersion -band 0x00ff0000) -shr 16 # bits 9 - 16 are the minor version number
                $build = $intVersion -band 0x0000ffff # last 16 bits are the build number
                $rawVersion = "$($major).$($minor).$($build)"
            }
            else
            {
                $rawVersion = Get-KeyStringValue -Key $key -ValueName 'Version'
            }

            [Version]$version = $null
            if( [Version]::TryParse($rawVersion, [ref]$version) )
            {
                $info.Version = $version
            }

            $windowsInstallerValue = Get-KeyIntValue -Key $key -ValueName 'WindowsInstaller'
            $info.WindowsInstaller = ($windowsInstallerValue -gt 0)

            $info.pstypenames.Insert(0, 'Carbon.Windows.Installer.ProgramInfo')
            $info | Write-Output
        } 
    } |
    Tee-Object -Variable 'programs' |
    Sort-Object -Property 'DisplayName'

    if( $Name -and -not [wildcardpattern]::ContainsWildcardCharacters($Name) -and -not $programs )
    {
        $msg = "Program ""$($Name)"" is not installed."
        Write-Error -Message $msg -ErrorAction $ErrorActionPreference
    }
}
