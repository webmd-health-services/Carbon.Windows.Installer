
function Read-CMsiTable
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [Object] $Database,

        [Parameter(Mandatory)]
        [String] $Name,

        [Parameter(Mandatory)]
        [String] $MsiPath
    )

    Set-StrictMode -Version 'Latest'
    Use-CallerPreference -Cmdlet $PSCmdlet -Session $ExecutionContext.SessionState

    $numErrors = $Global:Error.Count
    $view = $null

    try
    {
        $view = $Database.OpenView("select * from ``$($Name)``")
        if( -not $view )
        {
            $msg = "$($msiPath): failed to query $($Name) table."
            Write-Error -Message $msg -ErrorAction $ErrorActionPreference
            return
        }
    }
    catch
    {
        for( $idx = $numErrors; $idx -lt $Global:Error.Count; ++$idx )
        {
            $Global:Error.RemoveAt(0)
        }
        $msg = "Exception opening table ""$($Name)"" from MSI ""$($MsiPath)"": $($_)"
        Write-Debug -Message $msg
        return
    }

    $numErrors = $Global:Error.Count
    try
    {
        [void]$view.Execute()

        $colIdxToName = [Collections.ArrayList]::New()

        for( $idx = 0; $idx -le $view.ColumnInfo(0).FieldCount(); ++$idx )
        {
            $numErrors = $Global:Error.Count
            $columnName = $view.ColumnInfo(0).StringData($idx)
            Write-Debug "    [$($idx)] $columnName"
            [void]$colIdxToName.Add($columnName)
        }

        $msiRecord = $view.Fetch()
        while( $msiRecord )
        {
            $record = [pscustomobject]@{};
            Write-Debug '    +-----+'
            for( $idx = 0; $idx -lt $colIdxToName.Count; ++$idx )
            {
                $fieldName = $colIdxToName[$idx]
                if( -not $fieldName )
                {
                    continue
                }

                $fieldValue = $msiRecord.StringData($idx)
                Write-Debug "    [$($idx)][$($fieldName)]  $($fieldValue)"
                $record | Add-Member -Name $fieldName -MemberType NoteProperty -Value $fieldValue
            }
            $record.pstypenames.Insert(0, "Carbon.Windows.Installer.Records.$($Name)")
            $record | Write-Output
            $msiRecord = $view.Fetch()
        }
    }
    catch
    {
        $msg = "Exception reading ""$($Name)"" table record data from MSI ""$($MsiPath)"": " +
                "$($_)"
        Write-Debug -Message $msg
    }
    finally
    {
        if( $view )
        {
            [void]$view.Close()
        }
    }
}
