
function Install-Msi
{
    <#
    .SYNOPSIS
    Installs an MSI.

    .DESCRIPTION
    `Install-CMsi` installs software from an MSI file, without displaying any user interface. If the install fails, the
    function writes an error and leaves a debug-level log file in the current user's temp directory. The log file name
    uses the pattern `INSTALLER_FILE_NAME.RANDOM_FILENAME.RANDOM_EXTENSION.log`.

    Pass the path to the MSI to install to the `Path` property. The `Install-CMsi` function reads the product code from
    the MSI file, and does nothing if a program with that product code is already installed. Otherwise, it runs the
    installer in quiet mode with `msiexec`. All the program's features will be installed with their default values.

    `Install-CMsi` can also download an MSI and install it. Pass the URL to the MSI file to the `Url` parameter. Pass
    the MSI file's SHA256 checksum to the `Checksum` parameter. (Use PowerShell's `Get-FileHash` cmdlet to get the
    checksum.) In order avoid downloading an MSI that is already installed, you must also pass the MSI's product name to
    the `ProductName` parameter and its product code to the `ProductCode` parameter. Use this module's `Get-CMsi`
    function to get an MSI file's product metadata.

    If you want to install the MSI even if it is already installed, use the `-Force` switch. For downloaded MSI files,
    this will cause the file to be downloaded every time `Install-CMsi` is run.

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
        [Switch] $Force
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

                $installerLogFile = "$($Msi.Path | Split-Path -Leaf).$([IO.Path]::GetRandomFileName()).log"
                $installerLogFile = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath $installerLogFile
                New-Item -Path $installerLogFile -ItemType 'File' -WhatIf:$false | Out-Null

                if( -not $From )
                {
                    $From = $Msi.Path
                }

                $argumentList = @(
                    '/quiet',
                    '/i',
                    # Must surround with double quotes. Single quotes are interpreted as part of the path.
                    """$($msi.Path)""",
                    # Log EVERYTHING and flush after every line.
                    "/L!*VX",
                    # Must surround with double quotes. Single quotes are interpreted as part of the path.
                    """$($installerLogFile)"""
                )

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
                    $msiProcess = Start-Process -FilePath 'msiexec.exe' `
                                                -ArgumentList $argumentList `
                                                -NoNewWindow `
                                                -PassThru `
                                                -Wait

                    if( $null -ne $msiProcess.ExitCode -and $msiProcess.ExitCode -ne 0 )
                    {
                        $msg = "$($target) $($action.ToLowerInvariant()) failed. Installer ""$($msi.Path)"" returned " +
                               "exit code $($msiProcess.ExitCode). See the installation log file " +
                               """$($installerLogFile)"" for more information and " +
                               'https://docs.microsoft.com/en-us/windows/win32/msi/error-codes for a description of ' +
                               'the exit code.'
                        Write-Error $msg -ErrorAction $ErrorActionPreference
                        return
                    }
                }

                if( (Test-Path -Path $installerLogFile) )
                {
                    Remove-Item -Path $installerLogFile -ErrorAction Ignore -WhatIf:$false
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

        $outFile = $Url.Segments[-1]
        if( -not $outFile -or $outFile -eq '/' )
        {
            $fsReplaceRegex = [IO.Path]::GetInvalidFileNameChars() | ForEach-Object { [regex]::Escape($_) }
            $fsReplaceRegex = $fsReplaceRegex -join '|'
            $outFile = $Url.ToString() -replace $fsReplaceRegex, '_'
        }
        $outFile = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath $outFile
        $ProgressPreference = [Management.Automation.ActionPreference]::SilentlyContinue
        Invoke-WebRequest -Uri $Url -OutFile $outFile -UseBasicParsing | Out-Null

        $actualChecksum = Get-FileHash -LiteralPath $outFile
        if( $actualChecksum.Hash -ne $Checksum )
        {
            $msg = "Install failed: checksum ""$($actualChecksum.Hash.ToLowerInvariant())"" for installer " +
                   "downloaded from ""$($Url)"" does not match expected checksum ""$($Checksum.ToLowerInvariant())""."
            Write-Error -Message $msg -ErrorAction $ErrorActionPreference
            return
        }

        try
        {
            Get-CMsi -Path $outFile | Invoke-Msiexec -From $Url
        }
        finally
        {
            if( (Test-Path -LiteralPath $outFile -PathType Leaf) )
            {
                Remove-Item -LiteralPath $outFile -Force -ErrorAction Ignore
            }
        }
        Write-Debug "$($msgPrefix)-"
    }
}
