
function Get-Msi
{
    <#
    .SYNOPSIS
    Gets information from an MSI.

    .DESCRIPTION
    The `Get-CMsi` function uses the `WindowsInstaller.Installer` COM API to read properties from an MSI file. Pass the
    path to the MSI file or files to the `Path`, or pipe file objects to `Get-CMsi`. An object is returned that exposes
    the internal metadata of the MSI file:
    
    * `ProductName`, the value of the MSI's `ProductName` property.
    * `ProductCode`, the value of the MSI's `ProductCode` property as a `Guid`.
    * `ProduceLanguage`, the value of the MSI's `ProduceLanguage` property, as an integer.
    * `Manufacturer`, the value of the MSI's `Manufacturer` property.
    * `ProductVersion`, the value of the MSI's `ProductVersion` property, as a `Version`
    * `Path`, the path of the MSI file.
    * `TableNames`: the names of all the tables in the MSI's internal database
    * `Tables`: records from tables in the MSI's internal database

    The function can also return the records from the MSI's internal database tables. Tables included are returned as
    properties on the return object's `Tables` property. It is expensive to read all the records in all the database
    tables, so by default, `Get-CMsi` only returns the records from the `Property` and `Feature` tables. The `Property`
    table contains program metadata like product name, product code, etc. The `Feature` table contains the feature names
    of any optional features you might want to install. When installing, these feature names would get passed to the
    `msiexec` install command as a comma-separated list as the `ADDLOCAL` property, e.g. ADDLOCAL="Feature1,Feature2".

    To return the records from additional tables, pass the table name or names to the `IncludeTable` parameter.
    Wildcards supported. Records from the `Property` and `Feature` tables are *always* returned. The `TableNames`
    property on returned objects is the list of all tables in the MSI's database.

    Because this function uses the Windows Installer COM API, it requires Windows PowerShell 5.1 or PowerShell 7.1+ on
    Windows.

    .LINK
    https://msdn.microsoft.com/en-us/library/aa370905.aspx

    .EXAMPLE
    Get-CMsi -Path MyCool.msi

    Demonstrates how to get information from an MSI file.

    .EXAMPLE
    Get-ChildItem *.msi -Recurse | Get-CMsi

    Demonstrates that you can pipe file info objects into `Get-CMsi`.

    .EXAMPLE
    Get-CMsi -Path example.msi -IncludeTable 'Component'

    Demonstrates how to return records from one of an MSI's internal database tables by passing the table name to the
    `IncludeTable` parameter. Wildcards supported.
    #>
    [CmdletBinding()]
    param(
        # Path to the MSI file whose information to retrieve. Wildcards supported.
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [String[]] $Path,

        # Extra tables to read from the MSI and return. By default, only the installer's Property and Feature tables
        # are returned. Wildcards supported. See https://docs.microsoft.com/en-us/windows/win32/msi/database-tables for
        # the list of all MSI database tables.
        [String[]] $IncludeTable
    )

    begin
    {
        $timer = [Diagnostics.Stopwatch]::StartNew()
        $lastWrite = [Diagnostics.Stopwatch]::New()
        function Debug
        {
            param(
                [String] $Message
            )
    
            $msg = "[$([Math]::Round($timer.Elapsed.TotalMinutes))m $($timer.Elapsed.Seconds.toString('00'))s " +
                   "$($timer.Elapsed.Milliseconds.ToString('000'))ms]  " +
                   "[$([Math]::Round($lastWrite.Elapsed.TotalSeconds).ToString('00'))s " +
                   "$($lastWrite.Elapsed.Milliseconds.ToString('000'))ms]  $($Message)"
            Microsoft.PowerShell.Utility\Write-Debug -Message $msg
            $lastWrite.Restart()
        }

        $IncludeTable = & {
            'Feature'
            'Property'
            $IncludeTable | Write-Output
        } | Select-Object -Unique
    }

    process 
    {
        Set-StrictMode -Version 'Latest'
        Use-CallerPreference -Cmdlet $PSCmdlet -Session $ExecutionContext.SessionState

        $Path = Resolve-Path -Path $Path | Select-Object -ExpandProperty 'ProviderPath'
        if( -not $Path )
        {
            return
        }

        foreach( $msiPath in $Path )
        {
            $info = [pscustomobject]@{
                Manufacturer = $null;
                Path = $null;
                ProductCode = $null;
                ProductLanguage = $null;
                ProductName = $null;
                ProductVersion = $null;
                TableNames = @();
                Tables = [pscustomobject]@{};
            }
            $info |
                Add-Member -Name 'Name' -MemberType AliasProperty -Value 'ProductName' -PassThru |
                Add-Member -Name 'Property' -MemberType 'ScriptProperty' -Value { $this.Tables.Property } -PassThru |
                Add-Member -Name 'GetPropertyValue' -MemberType 'ScriptMethod' -Value {
                    param(
                        [Parameter(Mandatory)]
                        [String] $Name
                    )

                    if( -not $this.Property )
                    {
                        return
                    }

                    $this.Property | Where-Object 'Property' -eq $Name | Select-Object -ExpandProperty 'Value'
                }

            $installer = New-Object -ComObject 'WindowsInstaller.Installer'

            $database = $null;
            Debug "[$($PSCmdlet.MyInvocation.MyCommand.Name)]  Opening ""$($msiPath)""."
            try
            {
                $database = $installer.OpenDatabase($msiPath, 0)
                if( -not $database )
                {
                    $msg = "$($msiPath): failed to open database."
                    Write-Error -Message $msg -ErrorAction $ErrorActionPreference
                    continue
                }
            }
            catch
            {
                $msg = "Exception opening MSI database in file ""$($msiPath)"": $($_)"
                Write-Error -Message $msg -ErrorAction $ErrorActionPreference
                continue
            }

            try
            {
                Debug '    _Tables'
                $tables = Read-MsiTable -Database $database -Name '_Tables' -MsiPath $msiPath
                $info.TableNames = $tables | Select-Object -ExpandProperty 'Name'

                foreach( $tableName in $info.TableNames )
                {
                    $info.Tables | Add-Member -Name $tableName -MemberType NoteProperty -Value @()
                    if( $IncludeTable -and -not ($IncludeTable | Where-Object { $tableName -like $_ }) )
                    {
                        Debug "  ! $($tableName)"
                        continue
                    }

                    Debug "    $($tableName)"
                    $info.Tables.$tableName = Read-MsiTable -Database $database -Name $tableName -MsiPath $msiPath
                }

                [Guid] $productCode = [Guid]::Empty
                [String] $rawProductCode = $info.GetPropertyValue('ProductCode')
                if( [Guid]::TryParse($rawProductCode, [ref]$productCode) )
                {
                    $info.ProductCode = $productCode
                }

                [int] $langID = 0
                [String] $rawLangID = $info.GetPropertyValue('ProductLanguage')
                if( [int]::TryParse($rawLangID, [ref]$langID) )
                {
                    $info.ProductLanguage = $langID
                }

                $info.Path = $msiPath;
                $info.Manufacturer = $info.GetPropertyValue('Manufacturer')
                $info.ProductName = $info.GetPropertyValue('ProductName')
                $info.ProductVersion = $info.GetPropertyValue('ProductVersion')

                [void]$info.pstypenames.Insert(0, 'Carbon.Windows.Installer.MsiInfo')
            }
            finally
            {
                if( $database )
                {
                    [void][Runtime.InteropServices.Marshal]::ReleaseCOMObject($database)
                }
            }

            $info | Write-Output
        }
    }

    end
    {
        [GC]::Collect()
    }
}
