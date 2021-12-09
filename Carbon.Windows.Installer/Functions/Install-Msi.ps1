
function Install-Msi
{
    <#
    .SYNOPSIS
    Installs an MSI.

    .DESCRIPTION
    `Install-CMsi` installs software from an MSI file, without displaying any user interface. Pass the path to the MSI
    to install to the `Path` property. The `Install-CMsi` function reads the product name code from the MSI file, and
    does nothing if a program with that product code is already installed. Otherwise, the function runs the installer in
    quiet mode (i.e. no UI is visible) with `msiexec`. All the program's features will be installed with their default
    values. You can control the installer's display mode with the `DisplayMode` parameter: set it to `Passive` to show a
    UI with just a progress bar, or `Full` to show the UI as-if the user double-clicked the MSI file.

    `Install-CMsi` can also download an MSI and install it. Pass the URL to the MSI file to the `Url` parameter. Pass
    the MSI file's SHA256 checksum to the `Checksum` parameter. (Use PowerShell's `Get-FileHash` cmdlet to get the
    checksum.) In order avoid downloading an MSI that is already installed, you must also pass the MSI's product name to
    the `ProductName` parameter and its product code to the `ProductCode` parameter. Use this module's `Get-CMsi`
    function to get an MSI file's product metadata.

     If the install fails, the function writes an error and leaves a debug-level log file in the current user's temp
     directory. The log file name begins with the name of the MSI file name, then has a `.`, then a random file name
     (e.g. `xxxxxxxx.xxx`), then ends with a `.log` extension. You can customize the location of the log file with the
     `LogPath` parameter. You can customize logging options with the `LogOption` parameter. Default log options are
    `!*vx` (log all messages, all verbose message, all debug messages, and flush each line to the log file as it is
    written).

    If you want to install the MSI even if it is already installed, use the `-Force` switch. For downloaded MSI files,
    this will cause the file to be downloaded every time `Install-CMsi` is run.

    You can pass additional arguments to `msiexec.exe` when installing the MSI file with the `ArgumentList` parameter.

    Requires Windows PowerShell 5.1 or PowerShell 7.1+ on Windows.

    .EXAMPLE
    Install-CMsi -Path '.\Path\to\installer.msi'
    
    Demonstrates how to install a program with its MSI.

    .EXAMPLE
    Get-ChildItem *.msi | Install-CMsi

    Demonstrates that you can pipe file objects to `Install-CMsi`.

    .EXAMPLE
    Install-CMsi -Path 'installer.msi' -Force

    Demonstrates how to re-install an MSI file even if it's already installed.

    .EXAMPLE
    Install-CMsi -Url 'https://example.com/installer.msi' -Checksum '63c34def9153659a825757ec475a629dff5be93d0f019f1385d07a22a1df7cde' -ProductName 'Carbon Test Installer' -ProductCode 'e1724abc-a8d6-4d88-bbed-2e077c9ae6d2'

    Demonstrates that `Install-CMsi` can download and install MSI files.
    #>
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName='ByPath')]
    param(
        # The path to the installer to run. Wildcards supported.
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName='ByPath')]
        [Alias('FullName')]
        [string[]] $Path,

        # The URL to an installer to download and install. Requires the `Checksum` parameter to ensure the correct file
        # was downloaded.
        [Parameter(Mandatory, ParameterSetName='ByUrl')]
        [Uri] $Url,

        # Used with the `Url` parameter. The SHA256 hash the downloaded installer should have. Case-insensitive.
        [Parameter(Mandatory, ParameterSetName='ByUrl')]
        [String] $Checksum,

        # The product name of the downloaded MSI. Used to determine if the program is installed or not. Used with the
        # `Url` parameter. The installer is only downloaded if the product is not installed or the `-Force` switch is
        # used. Use the `Get-CMsi` function to get the product code of an MSI.
        [Parameter(Mandatory, ParameterSetName='ByUrl')]
        [String] $ProductName,
        
        # The product code of the downloaded MSI. Used to determine if the program is installed or not. Used with the
        # `Url` parameter. The installer is only downloaded if the product is not installed or the `-Force` switch is
        # used. Use the `Get-CMsi` function to get the product code from an MSI.
        [Parameter(Mandatory, ParameterSetName='ByUrl')]
        [Guid] $ProductCode,
        
        # Install the MSI even if it has already been installed. Will cause a repair/reinstall to run.
        [Switch] $Force,

        # Controls how the MSI UI is displayed to the user. The default is `Quiet`, meaning no UI is shown. Valid values
        # are `Passive`, a UI showing a progress bar is shown, or `Full`, the UI is displayed to the user as if they
        # double-clicked the MSI file.
        [ValidateSet('Quiet', 'Passive', 'Full')]
        [String] $DisplayMode = 'Quiet',

        # The logging options to use. The default is to log all information (`*`), log verbose output (`v`), log exta
        # debugging information (`x`), and to flush each line to the log (`!`).
        [String] $LogOption = '!*vx',

        # The path to the log file. The default is to log to a file in the temporary directory and delete the log file
        # unless the installation fails. The default log file name begins with the name of the MSI file name, then
        # has a `.`, then a random file name (e.g. `xxxxxxxx.xxx`), then ends with a `.log` extension.
        [String] $LogPath,

        # Extra arguments to pass to the installer. These are passed directly after the install option and path to the
        # MSI file. Do not pass any install option, display option, or logging option in this parameter. Instead, use
        # the `DisplayMode` parameter to control display options, the `LogOption` parameter to control logging options,
        # and the `LogPath` parameter to control where the installation log file is saved.
        [String[]] $ArgumentList
    )

    process
    {
        function Test-ProgramInstalled
        {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [String] $Name,

                [Guid] $Code
            )

            $DebugPreference = 'SilentlyContinue'
            $installInfo = Get-CInstalledProgram -Name $Name -ErrorAction Ignore
            if( -not $installInfo )
            {
                return $false
            }

            $installed = $installInfo.ProductCode -eq $Code
            if( $installed )
            {
                $msg = "$($msgPrefix)[$($installInfo.DisplayName)]  Installed $($installInfo.InstallDate)."
                Write-Verbose -Message $msg
                return $true
            }

            return $false
        }

        function Invoke-Msiexec
        {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory, ValueFromPipeline)]
                [Object] $Msi,

                [String] $From
            )

            process
            {
                $target = $Msi.ProductName
                if( $Msi.Manufacturer )
                {
                    $target = "$($Msi.Manufacturer)'s ""$($target)"""
                }

                if( $Msi.ProductVersion )
                {
                    $target = "$($target) $($Msi.ProductVersion)"
                }

                $deleteLog = $false
                if( -not $LogPath )
                {
                    $LogPath = "$($Msi.Path | Split-Path -Leaf).$([IO.Path]::GetRandomFileName()).log"
                    $LogPath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath $LogPath
                    $deleteLog = $true
                }
                $logParentDir = $LogPath | Split-Path -Parent
                if( $logParentDir -and -not (Test-Path -Path $logParentDir) )
                {
                    New-Item -Path $logParentDir -Force -ItemType 'Directory' | Out-Null
                }
                if( -not (Test-Path -Path $LogPath) )
                {
                    New-Item -Path $LogPath -ItemType 'File' | Out-Null
                }

                if( -not $From )
                {
                    $From = $Msi.Path
                }

                $displayOptions = @{
                    'Quiet' = '/quiet';
                    'Passive' = '/passive';
                    'Full' = '';
                }

                $ArgumentList = & {
                    '/i'
                    # Must surround with double quotes. Single quotes are interpreted as part of the path.
                    """$($msi.Path)"""
                    $displayOptions[$DisplayMode]
                    $ArgumentList | Write-Output
                    "/l$($LogOption)",
                    # Must surround with double quotes. Single quotes are interpreted as part of the path.
                    """$($LogPath)"""
                } | Where-Object { $_ }

                $action = 'Install'
                $verb = 'Installing'
                if( $Force )
                {
                    $action = 'Repair'
                    $verb = 'Repairing'
                }

                if( $PSCmdlet.ShouldProcess( $From, $action ) )
                {
                    Write-Information -Message "$($msgPrefix)$($verb) $($target) from ""$($From)"""
                    Write-Debug -Message "msiexec.exe $($ArgumentList -join ' ')"
                    $msiProcess = Start-Process -FilePath 'msiexec.exe' `
                                                -ArgumentList $ArgumentList `
                                                -NoNewWindow `
                                                -PassThru `
                                                -Wait

                    if( $null -ne $msiProcess.ExitCode -and $msiProcess.ExitCode -ne 0 )
                    {
                        $msg = "$($target) $($action.ToLowerInvariant()) failed. Installer ""$($msi.Path)"" returned " +
                               "exit code $($msiProcess.ExitCode). See the installation log file ""$($LogPath)"" for " +
                               'more information and https://docs.microsoft.com/en-us/windows/win32/msi/error-codes ' +
                               'for a description of the exit code.'
                        Write-Error $msg -ErrorAction $ErrorActionPreference
                        return
                    }
                }

                if( $deleteLog -and (Test-Path -Path $LogPath) )
                {
                    Remove-Item -Path $LogPath -ErrorAction Ignore
                }
            }
        }

        Set-StrictMode -Version 'Latest'
        Use-CallerPreference -Cmdlet $PSCmdlet -Session $ExecutionContext.SessionState

        $msgPrefix = "[$($MyInvocation.MyCommand.Name)]  "
        Write-Debug "$($msgPrefix)+"
        
        if( $Path )
        {
            Get-CMsi -Path $Path |
                Where-Object {
                    $msiInfo = $_

                    $installed = Test-ProgramInstalled -Name $msiInfo.ProductName -Code $msiInfo.ProductCode
                    if( $installed )
                    {
                        return $Force.IsPresent
                    }

                    # Not installed so $Force has no meaning. $Force also controls whether the action is "Install" or
                    # "Repair". We're always installing if not installed so set $Force to $false.
                    $Force = $false
                    return $true
                } |
                Invoke-Msiexec
            return
        }

        # If the program we are going to download is already installed, don't re-download it.
        $installed = Test-ProgramInstalled -Name $ProductName -Code $ProductCode
        if( $installed -and -not $Force )
        {
            return
        }
        # Make sure action is properly reported as Install or Repair.
        if( -not $installed )
        {
            $Force = $false
        }

        $msi = Get-Msi -Url $Url
        if( -not $msi )
        {
            return
        }

        $actualChecksum = Get-FileHash -LiteralPath $msi.Path
        if( $actualChecksum.Hash -ne $Checksum )
        {
            $msg = "Install failed: checksum ""$($actualChecksum.Hash.ToLowerInvariant())"" for installer " +
                   "downloaded from ""$($Url)"" does not match expected checksum ""$($Checksum.ToLowerInvariant())""."
            Write-Error -Message $msg -ErrorAction $ErrorActionPreference
            return
        }

        $msi | Invoke-Msiexec -From $Url

        Write-Debug "$($msgPrefix)-"
    }
}
