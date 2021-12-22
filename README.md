# Overview

The Carbon.Windows.Installer module is a Windows-only module that has functions for reading and installing Windows MSI
files/packages, and for replicating Windows' "Programs and Features"/"Apps and Features" graphical user interface but
with objects and text.

# System Requirements

* Windows PowerShell 5.1 on .NET 4.5.2+
* PowerShell 6.2, 7.1, or 7.2
* Windows Server 2012 R2+ or Windows 8.1+

# Installing

To install globally:

```powershell
Install-Module -Name 'Carbon.Windows.Installer'
Import-Module -Name 'Carbon.Windows.Installer'
```

To install privately:

```powershell
Save-Module -Name 'Carbon.Windows.Installer' -Path '.'
Import-Module -Name '.\Carbon.Windows.Installer'
```

# Commands

By default, commands have a `C` prefix. This can be changed with `Import-Module` cmdlet's `Prefix` parameter.

* `Get-CInstalledProgram`: reads the registry to determine all the programs installed on the local computer. Returns
an object for each program. Is an object-based and text-based version of Windows' "Programs and Features"/"Apps and
Features" GUI.
* `Get-CMsi`: reads an MSI file and returns an object exposing the MSI's internal tables, like product name, product
code, product version, etc. Let's you inspect an MSI for its metadata without installing it. Can also download the MSI
file first. This function requires Windows PowerShell or PowerShell 7.1+ on Windows.
* `Install-CMsi`: installs a program from an MSI file, or other file that can be installed by Windows. Can also 
download the MSI file to install. This function requires Windows PowerShell or PowerShell 7.1+ on Windows.
