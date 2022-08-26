# Copyright WebMD Health Services
#
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
# limitations under the License

@{

    # Script module or binary module file associated with this manifest.
    RootModule = 'Carbon.Windows.Installer.psm1'

    # Version number of this module.
    ModuleVersion = '2.0.0'

    # ID used to uniquely identify this module
    GUID = '57af6e60-d9f5-4523-8aa3-3e3480fd91cc'

    # Author of this module
    Author = 'WebMD Health Services'

    # Company or vendor of this module
    CompanyName = 'WebMD Health Services'

    # If you want to support .NET Core, add 'Core' to this list.
    CompatiblePSEditions = @( 'Desktop', 'Core' )

    # Copyright statement for this module
    Copyright = '(c) WebMD Health Services.'

    # Description of the functionality provided by this module
    Description = @'
The Carbon.Windows.Installer module is a Windows-only module that has functions for reading and installing Windows MSI
files/packages, and for replicating Windows' "Programs and Features"/"Apps and Features"
graphical user interface as objects and text.

Functions include:

* `Get-CInstalledProgram`: reads the registry to determine all the programs installed on the local computer. Returns
an object for each program. Is an object-based and text-based version of Windows' "Programs and Features"/"Apps and
Features" GUI.
* `Get-CMsi`: reads an MSI file and returns an object exposing the MSI's internal tables, like product name, product
code, product version, etc. Let's you inspect an MSI for its metadata without installing it. Can also download the MSI
file first. This function requires Windows PowerShell or PowerShell 7.1+ on Windows.
* `Install-CMsi`: installs a program from an MSI file, or other file that can be installed by Windows. Can also download
the MSI file to install. This function requires Windows PowerShell or PowerShell 7.1+ on Windows.

System Requirements:

* Windows PowerShell 5.1 on .NET 4.5.2+
* PowerShell 6.2, 7.1, or 7.2
* Windows Server 2012 R2+ or Windows 8.1+
'@

    # Minimum version of the Windows PowerShell engine required by this module
    PowerShellVersion = '5.1'

    # Name of the Windows PowerShell host required by this module
    # PowerShellHostName = ''

    # Minimum version of the Windows PowerShell host required by this module
    # PowerShellHostVersion = ''

    # Minimum version of Microsoft .NET Framework required by this module
    # DotNetFrameworkVersion = ''

    # Minimum version of the common language runtime (CLR) required by this module
    # CLRVersion = ''

    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = ''

    # Modules that must be imported into the global environment prior to importing this module
    # RequiredModules = @()

    # Assemblies that must be loaded prior to importing this module
    # RequiredAssemblies = @( )

    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    # ScriptsToProcess = @()

    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()

    # Format files (.ps1xml) to be loaded when importing this module
    FormatsToProcess = @(
        'Formats\MsiInfo.ps1xml',
        'Formats\ProgramInfo.ps1xml',
        'Formats\Records.Feature.ps1xml'
    )

    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    # NestedModules = @()

    # Functions to export from this module. Only list public function here.
    FunctionsToExport = @(
        'Get-CMsi',
        'Get-CInstalledProgram',
        'Install-CMsi'
    )

    # Cmdlets to export from this module. By default, you get a script module, so there are no cmdlets.
    # CmdletsToExport = @()

    # Variables to export from this module. Don't export variables except in RARE instances.
    VariablesToExport = @()

    # Aliases to export from this module. Don't create/export aliases. It can pollute your user's sessions.
    AliasesToExport = @()

    # DSC resources to export from this module
    # DscResourcesToExport = @()

    # List of all modules packaged with this module
    # ModuleList = @()

    # List of all files packaged with this module
    # FileList = @()

    # HelpInfo URI of this module
    # HelpInfoURI = ''

    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    # DefaultCommandPrefix = ''

    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData = @{

        PSData = @{

            # Tags applied to this module. These help with module discovery in online galleries.
            Tags = @(
                'Desktop',
                'msi',
                'package',
                'programs',
                'features',
                'installer',
                'feature',
                'component',
                'msiexec'
            )

            # A URL to the license for this module.
            LicenseUri = 'http://www.apache.org/licenses/LICENSE-2.0'

            # A URL to the main website for this project.
            ProjectUri = 'https://github.com/webmd-health-services/Carbon.Windows.Installer'

            # A URL to an icon representing this module.
            # IconUri = ''

            Prerelease = ''

            # ReleaseNotes of this module
            ReleaseNotes = 'https://github.com/webmd-health-services/Carbon.Windows.Installer/blob/main/CHANGELOG.md'

        } # End of PSData hashtable

    } # End of PrivateData hashtable
}
