
function Get-Msi
{
    <#
    .SYNOPSIS
    Gets information from an MSI.

    .DESCRIPTION
    The `Get-CMsi` function uses the `WindowsInstaller.Installer` COM API to read properties from an MSI file. It
    returns a `Carbon.Windows.Installer.MsiInfo` object with the following properties:

    * `ProductName`, the value of the MSI's `ProductName` property.
    * `ProductCode`, the value of the MSI's `ProductCode` property as a `Guid`.
    * `ProduceLanguage`, the value of the MSI's `ProduceLanguage` property, as an integer.
    * `Manufacturer`, the value of the MSI's `Manufacturer` property.
    * `ProductVersion`, the value of the MSI's `ProductVersion` property, as a `Version`
    * `Property`, an array of `Carbon.Windows.Installer.MsiInfo.Records.Property` objects for each property in the MSI.
      Each property has `Property` and `Value` properties for the name and value of the property, respectively.
    * `Path`, the path of the MSI file.

    .LINK
    https://msdn.microsoft.com/en-us/library/aa370905.aspx

    .EXAMPLE
    Get-CMsi -Path MyCool.msi

    Demonstrates how to get information from an MSI file.

    .EXAMPLE
    Get-ChildItem *.msi -Recurse | Get-CMsi

    Demonstrates that you can pipe file info objects into `Get-CMsi`.
    #>
    [CmdletBinding()]
    param(
        # Path to the MSI file whose information to retrieve. Wildcards supported.
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [String[]] $Path
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
                Tables = @{};
            }
            $info |
                Add-Member -Name 'Name' -MemberType AliasProperty -Value 'ProductName' -PassThru |
                Add-Member -Name 'Property' -MemberType 'ScriptProperty' -Value { $this.Tables['Property'] } -PassThru |
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
                $info.Tables['Property'] = Read-MsiTable -Database $database -Name 'Property' -MsiPath $msiPath

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
