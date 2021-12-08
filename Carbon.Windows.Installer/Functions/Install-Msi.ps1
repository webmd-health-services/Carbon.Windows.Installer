
function Install-Msi
{
    <#
    .SYNOPSIS
    Installs an MSI.

    .DESCRIPTION
    `Install-CMsi` installs software from an MSI file, without displaying any user interface. If the install fails, the
    function writes an error and leaves a debug-level log file in the current user's temp directory using the
    installer's file name, a random filename, and the .log extension.

    Pass the path to the MSI to install to the `Path` property. The `Install-CMsi` function reads the product code from
    the MSI file, and does nothing if a program with that product code is already installed. Otherwise, it runs the
    installer in quiet mode with `msiexec`. All the program's features will be installed with their default values.

    If you want to install the program even if it is already installed, use the `-Force` switch.

    Requires Windows PowerShell 5.1 or PowerShell 7.1+ on Windows.

    .EXAMPLE
    Install-CMsi -Path '.\Path\to\installer.msi'
    
    Demonstrates how to install a program with its MSI.

    .EXAMPLE
    Get-ChildItem *.msi | Install-CMsi

    Demonstrates that you can pipe file objects to `Install-CMsi`.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        # The path to the installer to run. Wildcards supported.
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [string[]] $Path,
        
        # Install the MSI even if it has already been installed. Will cause a repair/reinstall to run.
        [Switch] $Force
    )

    process
    {
        Set-StrictMode -Version 'Latest'
        Use-CallerPreference -Cmdlet $PSCmdlet -Session $ExecutionContext.SessionState

        $msgPrefix = "[$($MyInvocation.MyCommand.Name)]  "
        Write-Debug "$($msgPrefix)+"
        
        Get-CMsi -Path $Path |
            Where-Object {
                $msiInfo = $_

                if( $Force )
                {
                    return $true
                }

                $installInfo = Get-CInstalledProgram -Name $msiInfo.ProductName -ErrorAction Ignore
                if( -not $installInfo )
                {
                    return $true
                }

                $installed = $installInfo.ProductCode -eq $msiInfo.ProductCode
                if( $installed )
                {
                    $msg = "$($msgPrefix)[$($installInfo.DisplayName)]  Installed $($installInfo.InstallDate)."
                    Write-Verbose -Message $msg
                    return $false
                }

                return $true
            } |
            ForEach-Object {
                $msi = $_
                $target = $msi.ProductName
                if( $msi.Manufacturer )
                {
                    $target = "$($msi.Manufacturer)'s $($target)"
                }

                if( $msi.ProductVersion )
                {
                    $target = "$($target) version $($msi.ProductVersion)"
                }

                $installerLogFile = "$($msi.Path | Split-Path -Leaf).$([IO.Path]::GetRandomFileName()).log"
                $installerLogFile = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath $installerLogFile
                New-Item -Path $installerLogFile -ItemType 'File' -WhatIf:$false | Out-Null

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
                if( $PSCmdlet.ShouldProcess( "$($target) from '$($msi.Path)'", 'Install' ) )
                {
                    Write-Information -Message "$($msgPrefix)Installing $($target) from ""$($msi.Path)"""
                    $msiProcess = Start-Process -FilePath 'msiexec.exe' `
                                                -ArgumentList $argumentList `
                                                -NoNewWindow `
                                                -PassThru `
                                                -Wait

                    if( $null -ne $msiProcess.ExitCode -and $msiProcess.ExitCode -ne 0 )
                    {
                        $msg = "$($target) installation failed. Installer ""$($msi.Path)"" returned exit code " +
                               "$($msiProcess.ExitCode). See the installation log file ""$($installerLogFile)"" for " +
                               'more information and https://docs.microsoft.com/en-us/windows/win32/msi/error-codes ' +
                               'for a description of the exit code.'
                        Write-Error $msg -ErrorAction $ErrorActionPreference
                        return
                    }
                }

                if( (Test-Path -Path $installerLogFile) )
                {
                    Remove-Item -Path $installerLogFile -ErrorAction Ignore -WhatIf:$false
                }
            }

        Write-Debug "$($msgPrefix)-"
    }
}
