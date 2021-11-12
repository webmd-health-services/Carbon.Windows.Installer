# Overview

The "Carbon.Windows.Installer" module 

# System Requirements

* Windows Server 2012 R2 or later, including corresponding desktop/consumer operating systems
* Windows PowerShell 5.1 or later
* PowerShell Core 7 or later

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

* `Get-CMsi`: reads an MSI file and returns an object exposing the MSI file's properties like product name, code,
  manufacturer, etc.
* `Get-CInstalledProgram`: reads the registry and returns objects for each installed program. Mimics the Programs and
  Features/Add and Features GUIs.
* `Install-CMsi`: runs an MSI to install a program.
