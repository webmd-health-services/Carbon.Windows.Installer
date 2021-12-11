
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

    .EXAMPLE
    Get-CMsi -Url 'https://example.com/example.msi'

    Demonstrates how to download an MSI file to read its metadata. The file is saved to the current user's temp
    directory with the same name as the file name in the URL. The return object will have the path to the MSI file.

    .EXAMPLE
    Get-CMsi -Url 'https://example.com/example.msi' -OutputPath '~\Downloads'

    Demonstrates how to download an MSI file and save it to a directory using the name of the file from the download
    URL as the filename. In this case, the file will be saved to `~\Downloads\example.msi`. The return object's `Path`
    property will contain the full path to the downloaded MSI file.

    .EXAMPLE
    Get-CMsi -Url 'https://example.com/example.msi' -OutputPath '~\Downloads\new_example.msi'

    Demonstrates how to use a custom file name for the downloaded file by making `OutputPath` be a path to an item that
    doesn't exist or the path to an existing file.
    #>
    [CmdletBinding(DefaultParameterSetName='ByPath')]
    param(
        # Path to the MSI file whose information to retrieve. Wildcards supported.
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName='ByPath',
                   Position=0)]
        [Alias('FullName')]
        [String[]] $Path,

        # The URL to the MSI file to get. The file will be downloaded to the current user's temp directory. Use the
        # `OutputPath` parameter to save it somewhere else or use the `Path` property on the returned object to copy the
        # downloaded file somewhere else.
        [Parameter(Mandatory, ParameterSetName='ByUrl')]
        [Uri] $Url,

        # The path where the downloaded MSI file should be saved. By default, the file is downloaded to the current
        # user's temp directory. If `OutputPath` is a directory, the file will be saved to that directory with the 
        # same name as file's name in the `Url`. Otherwise, `OutputPath` is considered to be the path to the file where
        # the downloaded MSI should be saved. Any existing file will be overwritten.
        [Parameter(ParameterSetName='ByUrl')]
        [String] $OutputPath,

        # Extra tables to read from the MSI and return. By default, only the installer's Property and Feature tables
        # are returned. Wildcards supported. See https://docs.microsoft.com/en-us/windows/win32/msi/database-tables for
        # the list of all MSI database tables.
        [String[]] $IncludeTable
    )

    begin
    {
        Set-StrictMode -Version 'Latest'
        Use-CallerPreference -Cmdlet $PSCmdlet -Session $ExecutionContext.SessionState

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

        if( $PSCmdlet.ParameterSetName -eq 'ByUrl' )
        {
            $msiFileName = $Url.Segments[-1]
            if( $OutputPath )
            {
                if( (Test-Path -Path $OutputPath -PathType Container) )
                {
                    $OutputPath = Join-Path -Path $OutputPath -ChildPath $msiFileName
                }
            }
            else
            {
                $OutputPath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath $msiFileName
            }
            $ProgressPreference = [Management.Automation.ActionPreference]::SilentlyContinue
            Invoke-WebRequest -Uri $Url -OutFile $OutputPath | Out-Null
            Get-Item -LiteralPath $OutputPath | Get-Msi -IncludeTable $IncludeTable
            return
        }

        $IncludeTable = & {
            'Feature'
            'Property'
            $IncludeTable | Write-Output
        } | Select-Object -Unique
    }

    process
    {
        if( $PSCmdlet.ParameterSetName -eq 'ByUrl' )
        {
            return
        }

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
                $collect = $false
                if( $database )
                {
                    [void][Runtime.InteropServices.Marshal]::ReleaseCOMObject($database)
                    $collect = $true
                }

                if( $installer )
                {
                    [void][Runtime.InteropServices.Marshal]::ReleaseCOMObject($installer)
                    $collect = $true
                }

                if( $collect )
                {
                    # ReleaseCOMObject still leaves the MSI file open. The only way to close the file handle is to run
                    # garbage collection, and even then it takes a few seconds. :(
                    Debug "[GC]::Collect()  START"
                    [GC]::Collect()
                    Debug "[GC]::Collect()  END"
                }
            }

            # It can take some milliseconds for the COM file handles to get closed. In my testing, about 10 to 30
            # milliseconds. I give it 100ms just to be safe. But don't keep trying because something else might 
            # legitimately have the file open. 100ms is the longest something can take without a human wondering what's
            # taking so long.
            $timer = [Diagnostics.Stopwatch]::StartNew()
            $numAttempts = 1
            $maxTime = [TimeSpan]::New(0, 0, 0, 0, 100)
            while( $timer.Elapsed -lt $maxTime )
            {
                $numErrors = $Global:Error.Count
                try
                {
                    # Wait until the file handle held by the WindowsInstaller COM objects is closed.
                    [IO.File]::Open($msiPath, 'Open', 'Read', 'None').Close()
                    break
                }
                catch
                {
                    ++$numAttempts
                    for( $numErrors; $numErrors -lt $Global:Error.Count; ++$numErrors )
                    {
                        $Global:Error.RemoveAt(0)
                    }
                    Start-Sleep -Milliseconds 1
                }
            }
            $timer.Stop()
            $msg = "Took $($numAttempts) attempt(s) in " +
                   "$($timer.Elapsed.TotalSeconds.ToString('0.000'))s for handle to ""$($msiPath)"" to close."
            Debug $msg
            $info | Write-Output
        }
    }
}
