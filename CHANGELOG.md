
# 1.0.0

The `Carbon.Windows.Installer` module was created from functions in the `Carbon` module. These release notes assume
you're migrating from [Carbon](http://get-carbon.org). If you're not migrating, you can ignore these release notes.

## Migration Instructions

If migrating from Carbon, you'll need to make the following changes:

### General

* `Carbon.Windows.Installer` requires Windows PowerShell 5.1 or PowerShell 6.2+ on Windows.

### Get-CMsi

* Rename usages of `Get-Msi` to `Get-CMsi` (the `C` prefix is now required).
* Rename `Properties` property usages to `Property` on objects returned from `Get-CMsi`. The `Properties` property on
  objects returned from `Get-CMsi` was renamed to `Property`.
* Change property lookups on the objects returned from `Get-CMsi` from `$msiInfo.Properties[KEY]` to
  `$msiInfo.GetPropertyValue(KEY)`. The `Properties` property is now named `Property` and is an array of
  `Carbon.Windows.Installer.Records.Property` objects (instead of a hashtable) and we added a `GetPropertyValue` method
  to lookup specific property values.
* Remove usages of the `[Carbon.Msi.MsiInfo]` type. It was removed. `Get-CMsi` now returns `[PSObject]` objects with a
  pstypename of `[Carbon.Windows.Installer.MsiInfo]`.

### Get-CProgramInstallInfo

* Rename usages of `Get-CProgramInstallInfo` to `Get-CInstalledProgram`. We renamed `Get-CProgramInstallInfo` to 
  `Get-CInstalledProgram`. The `C` prefix is now required.
* Remove usages of the `[Carbon.Computer.ProgramInstallInfo]` type. `Get-CInstalledProgram` now returns `[PSObject]`
objects with a pstypename of `[Carbon.Windows.Installer.ProgramInfo]`.
* Add `-ErrorAction Ignore` or `-ErrorAction SilentlyContinue` to usages of `Get-CInstalledProgram` that handle if no
object is returned. `Get-CInstalledProgram` now writes an error if a program isn't installed.

### Install-CMsi

* Remove usages of the `-Quiet` switch. The `-Quiet` parameter was removed.

## Release Notes

### Added

* You can now change the default `C` prefix when importing `Carbon.Windows.Installer` with the `Import-Module` cmdlet's
`-Prefix` parameter.
* `Get-CMsi`: returned objects now have a `GetPropertyValue([String])` method for getting property values.
* `Get-CInstalledProgram` (i.e. `Get-CProgramInstallInfo`) now writes an error when a program isn't found.
* `Get-CMsi` can now read records from *all* of an MSI's internal database tables. By default, only records from the
`Product` and `Feature` tables are returned. The objects returned by `Get-CMsi` now have a `TableNames` property, which
is the list of all the tables in the MSI's database, and `Tables`, which contains properties for each included table.
To include a table, pass its name (or a wildcard) to the new `IncludeTable` parameter.
* `Get-CMsi` can now download MSI files. Pass the URL to the MSI file to the new `Url` parameter. Use the `Path` of the
return object to get the file's location. Use the new `OutputPath` parameter to save the installer to a specific
directory or to a specific file.
* When an installer fails, `Install-CMsi` leaves a debug log of the installation in the user's temp directory. The
file's name begins with the installer's file name followed by a random file name and has a `.log` extension.
* `Install-CMsi` can now download an MSI file to install. Pass the URL to the installer to the `Url` parameter, the
file's SHA256 checksum to the `Checksum` parameter, its product name to the `ProductName` parameter, and its 
product code to the `ProductCode` parameter. The MSI file will only be downloaded if it isn't installed. You can get the MSI's checksum with PowerShell's `Get-FileHash` cmdlet. You can get the MSI's product name and code with this module's `Get-CMsi` function.
* Added `ArgumentList` parameter to `Install-CMsi` to pass additional arguments to `msiexec.exe` when installing an MSI.
* Added `DisplayOption`, `LogOption`, and `LogPath` parameters to allow users to control the display and logging options passed to `msiexec.exe`.

### Changed

* Minimum system requirements are now Windows PowerShell 5.1 on .NET 4.5.2+ or PowerShell 6.2+ on Windows. The
`Get-CMsi` and `Install-CMsi` functions don't work in PowerShell 6.2.
* `Get-CMsi`: The `Properties` property on returned objects renamed to `Property`.
* `Get-CMsi`: The `Properties` property on returned objects is no longer a hashtable but an array of
`[Carbon.Windows.Installer.Records.Property]` objects. To find the value of a property, use the
`GetPropertyValue([String])` method.
* `Get-CMsi` now returns `[PSObject]` objects (instead of `[Carbon.Msi.MsiInfo]` objects) with a pstypename of
`[Carbon.Windows.Installer.MsiInfo]`.
* `Get-CProgramInstallInfo` renamed to `Get-CInstalledProgram`.
* `Get-CInstalledProgram` no longer writes errors when not running as an administrator.
* `Get-CInstalledProgram` now returns `[PSObject]` objects (instead of `[Carbon.Computer.ProgramInstallInfo]` objects)
with a pstypename of `[Carbon.Windows.Installer.ProgramInfo]`.

## Removed

* `Get-Msi` backward-compatible shim function. Use `Get-CMsi` instead.
* `Get-ProgramInstallInfo` backward-compatible shim function. Use `Get-CInstalledProgram` instead.
* `[Carbon.Msi.MsiInfo]` .NET object.
* `[Carbon.Computer.ProgramInstallInfo]` .NET object.
* `Install-Msi` backwards-compatible shim function. Use `Install-CMsi` instead.
