
# 1.0.0

## Upgrade Instructions

If migrating from Carbon, you'll need to make the following changes:

### General

* `Carbon.Windows.Installer` requires PowerShell 5.1 or later.

### Get-CMsi

* Rename usages of `Get-Msi` to `Get-CMsi` (the `C` prefix is now required).
* Rename `Properties` property usages to `Property` on objects returned from `Get-CMsi`. The `Properties` property on
  objects returned from `Get-CMsi` was renamed to `Property`.
* Change property lookups on the objects returned from `Get-CMsi` from `$msiInfo.Properties[KEY]` to
  `$msiInfo.GetPropertyValue(KEY)`. The `Properties` property is now an array of
  `Carbon.Windows.Installer.Records.Property` objects (instead of a hashtable) and we added a `GetPropertyValue` method
  to lookup specific property values.
* Remove usages of the `[Carbon.Msi.MsiInfo]` type. It was removed. `Get-CMsi` now returns `[PSObject]` objects with 
  pstypename `[Carbon.Windows.Installer.MsiInfo]`.

### Get-CProgramInstallInfo

* Rename usages of `Get-CProgramInstallInfo` to `Get-CInstalledProgram`. We renamed `Get-CProgramInstallInfo` to 
  `Get-CInstalledProgram`. The `C` prefix is now required.
* Remove usages of the `[Carbon.Computer.ProgramInstallInfo]`. `Get-CInstalledProgram` now returns `[PSObject]` objects
  with pstypename `[Carbon.Windows.Installer.ProgramInfo]`.

### Install-CMsi

* Remove usages of the `-Quiet` switch. The `-Quiet` parameter was removed.

## Changes

### General

* `Carbon.Windows.Installer` requires PowerShell 5.1 or later.
* You can now change the default `C` prefix when importing `Carbon.Windows.Installer`.

### Get-CMsi

* `Get-Msi` backwards-compatible functions removed. You must now use `Get-CMsi`.
* The `Properties` property renamed to `Property` on objects returned from `Get-CMsi`.
* The `Properties` property is no longer a hashtable but an array of `[Carbon.Windows.Installer.Records.Property]` 
  objects. To find an the property value of an installer, use the `GetPropertyValue([String])` method on the object returned by `Get-CMsi`
* Removed the `[Carbon.Msi.MsiInfo]` .NET object. `Get-CMsi` now returns `[PSObject]` objects with pstypename
  `[Carbon.Windows.Installer.MsiInfo]`.

### Get-CInstalledProgram (f.k.a Get-CProgramInstallInfo)

* `Get-CProgramInstallInfo` renamed to `Get-CInstalledProgram`.
* `Get-ProgramInstallInfo` function removed. Use `Get-CInstalledProgram` instead.
* `Get-CInstalledProgram` no longer writes errors when not running as an administrator.
* Removed the `[Carbon.Computer.ProgramInstallInfo]` .NET object. `Get-CInstalledProgram` now returns `[PSObject]`
  objects with pstypename `[Carbon.Windows.Installer.ProgramInfo]`.

### Install-Msi

* `Install-Msi` backwards-compatible function removed. You must now use `Install-CMsi`.
