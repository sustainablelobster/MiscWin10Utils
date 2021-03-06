# MiscWin10Utils Module
## ConvertTo-WSLPath
### Synopsis
Converts a Windows path to its corresponding WSL path.
### Syntax
```powershell

ConvertTo-WSLPath [-Path] <String> [-Full] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Path</nobr> |  | Windows path to convert | true | true \(ByValue\) |  |
| <nobr>Full</nobr> |  | Use the Full switch to get get the full path \(default is relative\) | false | false | False |
### Inputs
 - System.String You can pipe paths to this function.

### Outputs
 - System.String

### Examples
**EXAMPLE 1**
```powershell
ConvertTo-WSLPath -Path "." -Full
```
Get the full WSL path of the current directory.

**EXAMPLE 2**
```powershell
ConvertTo-WSLPath -Path ".\dir\subdir\file.bin"
```
Get the relative WSL path of the given Windows path.

**EXAMPLE 3**
```powershell
(Get-ChildItem -Path "." -Recurse).FullName | ConvertTo-WSLPath
```
Get WSL paths for every item in the current directory.

### Links

 - [https://docs.microsoft.com/en-us/windows/wsl/about](https://docs.microsoft.com/en-us/windows/wsl/about)
## Disable-AutoAccentColor
### Synopsis
Disables automatic selection of UI accent color.
### Syntax
```powershell

Disable-AutoAccentColor [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Disable-AutoAccentColor -RestartExplorer
```
Disable automatic accents and restart explorer.exe so changes take effect immediately.

### Links

 - [Enable-AutoAccentColor](#Enable-AutoAccentColor)
## Disable-CortanaButton
### Synopsis
Disables Cortana button on Taskbar.
### Syntax
```powershell

Disable-CortanaButton [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Disable-CortanaButton -RestartExplorer
```
Disable Cortana button and restart explorer.exe so changes take effect immediately.

### Links

 - [Enable-CortanaButton](#Enable-CortanaButton)
## Disable-DesktopRecycleBin
### Synopsis
Hide Recycle Bin on Desktop.
### Syntax
```powershell

Disable-DesktopRecycleBin [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Disable-DesktopRecycleBin -RestartExplorer
```
Hide Recycle Bin on Desktop and restart explorer.exe so changes take effect immediately.

### Links

 - [Enable-DesktopRecycleBin](#Enable-DesktopRecycleBin)
## Disable-ExplorerQuickAccess
### Synopsis
Disables the Quick Access menu in Explorer system-wide.
### Syntax
```powershell

Disable-ExplorerQuickAccess [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Disable-ExplorerQuickAccess -RestartExplorer
```
Disable Quick Access and restart explorer.exe so changes take effect immediately.

### Links

 - [Enable-ExplorerQuickAccess](#Enable-ExplorerQuickAccess)
## Disable-InkWorkspaceButton
### Synopsis
Hide the Ink Workspace button on the Taskbar.
### Syntax
```powershell

Disable-InkWorkspaceButton [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Disable-InkWorkspaceButton -RestartExplorer
```
Hide Ink Workspace button and restart explorer.exe so changes take effect immediately.

### Links

 - [Enable-InkWorkspaceButton](#Enable-InkWorkspaceButton)
## Disable-TransparencyEffects
### Synopsis
Disable system UI transparency effects.
### Syntax
```powershell

Disable-TransparencyEffects [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Disable-TransparencyEffects -RestartExplorer
```
Disable system UI transparency effects and restart explorer.exe so changes take effect immediately.

### Links

 - [Enable-TransparencyEffects](#Enable-TransparencyEffects)
## Enable-AppDarkTheme
### Synopsis
Set application theme to dark mode.
### Syntax
```powershell

Enable-AppDarkTheme [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Enable-AppDarkTheme -RestartExplorer
```
Set app theme to dark mode and restart explorer.exe so changes take effect immediately.

### Links

 - [Enable-AppLightTheme](#Enable-AppLightTheme)
 - [Enable-SystemLightTheme](#Enable-SystemLightTheme)
 - [Enable-SystemDarkTheme](#Enable-SystemDarkTheme)
## Enable-AppLightTheme
### Synopsis
Set application theme to light mode.
### Syntax
```powershell

Enable-AppLightTheme [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Enable-AppLightTheme -RestartExplorer
```
Set app theme to light mode and restart explorer.exe so changes take effect immediately.

### Links

 - [Enable-AppDarkTheme](#Enable-AppDarkTheme)
 - [Enable-SystemLightTheme](#Enable-SystemLightTheme)
 - [Enable-SystemDarkTheme](#Enable-SystemDarkTheme)
## Enable-AutoAccentColor
### Synopsis
Let Windows pick UI accent color automatically.
### Syntax
```powershell

Enable-AutoAccentColor [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Enable-AutoAccentColor -RestartExplorer
```
Enable automatic accents and restart explorer.exe so changes take effect immediately.

### Links

 - [Disable-AutoAccentColor](#Disable-AutoAccentColor)
## Enable-CortanaButton
### Synopsis
Enables Cortana button on Taskbar.
### Syntax
```powershell

Enable-CortanaButton [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Enable-CortanaButton -RestartExplorer
```
Enable Cortana button and restart explorer.exe so changes take effect immediately.

### Links

 - [Disable-CortanaButton](#Disable-CortanaButton)
## Enable-DesktopRecycleBin
### Synopsis
Show Recycle Bin on Desktop.
### Syntax
```powershell

Enable-DesktopRecycleBin [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Enable-DesktopRecycleBin -RestartExplorer
```
Show Recycle Bin on Desktop and restart explorer.exe so changes take effect immediately.

### Links

 - [Disable-DesktopRecycleBin](#Disable-DesktopRecycleBin)
## Enable-ExplorerQuickAccess
### Synopsis
Enables the Quick Access menu in Explorer system-wide.
### Syntax
```powershell

Enable-ExplorerQuickAccess [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Enable-ExplorerQuickAccess -RestartExplorer
```
Enable Quick Access and restart explorer.exe so changes take effect immediately.

### Links

 - [Disable-ExplorerQuickAccess](#Disable-ExplorerQuickAccess)
## Enable-InkWorkspaceButton
### Synopsis
Show the Ink Workspace button on the Taskbar.
### Syntax
```powershell

Enable-InkWorkspaceButton [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Enable-InkWorkspaceButton -RestartExplorer
```
Show Ink Workspace button and restart explorer.exe so changes take effect immediately.

### Links

 - [Disable-InkWorkspaceButton](#Disable-InkWorkspaceButton)
## Enable-SystemDarkTheme
### Synopsis
Set system theme to dark mode.
### Syntax
```powershell

Enable-SystemDarkTheme [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Enable-SystemDarkTheme -RestartExplorer
```
Set system theme to dark mode and restart explorer.exe so changes take effect immediately.

### Links

 - [Enable-SystemLightTheme](#Enable-SystemLightTheme)
 - [Enable-AppLightTheme](#Enable-AppLightTheme)
 - [Enable-AppDarkTheme](#Enable-AppDarkTheme)
## Enable-SystemLightTheme
### Synopsis
Set system theme to light mode.
### Syntax
```powershell

Enable-SystemLightTheme [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Enable-SystemLightTheme -RestartExplorer
```
Set system theme to light mode and restart explorer.exe so changes take effect immediately.

### Links

 - [Enable-SystemDarkTheme](#Enable-SystemDarkTheme)
 - [Enable-AppLightTheme](#Enable-AppLightTheme)
 - [Enable-AppDarkTheme](#Enable-AppDarkTheme)
## Enable-TransparencyEffects
### Synopsis
Enable system UI transparency effects.
### Syntax
```powershell

Enable-TransparencyEffects [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Enable-TransparencyEffects -RestartExplorer
```
Enable system UI transparency effects and restart explorer.exe so changes take effect immediately.

### Links

 - [Disable-TransparencyEffects](#Disable-TransparencyEffects)
## Install-Boxstarter
### Synopsis
Installs Boxstarter.
### Syntax
```powershell

Install-Boxstarter [<CommonParameters>]





```
### Inputs
 - None

### Outputs
 - None

### Note
Function body is based on Boxstarter's "installing from the web" instructions.

### Examples
**EXAMPLE 1**
```powershell
Install-Boxstarter
```


### Links

 - [https://boxstarter.org/](https://boxstarter.org/)
 - [https://boxstarter.org/installboxstarter#installing-from-the-web](https://boxstarter.org/installboxstarter#installing-from-the-web)
## Install-EdgeExtension
### Synopsis
Installs the given Edge extension.
### Syntax
```powershell

Install-EdgeExtension [-ExtensionUrl] <String> [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>ExtensionUrl</nobr> |  | URL of the extension's Edge Add-ons page or Chrome Web Store page. | true | true \(ByValue\) |  |
### Inputs
 - System.String You can pipe extension URLs to this function.

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Install-EdgeExtension -ExtensionUrl "https://microsoftedge.microsoft.com/addons/detail/dark-reader/ifoakfbpdcdoeenechcleahebpibofpc"
```
Install the "Dark Reader" extension from the Edge Add-ons store.

**EXAMPLE 2**
```powershell
Install-EdgeExtension -ExtensionUrl "https://chrome.google.com/webstore/detail/picture-in-picture-extens/hkgfoiooedgoejojocmhlaklaeopbecg"
```
Install the "Picture-in-Picture Extension \(by Google\)" extension from the Chrome Web Store.

### Links

 - [https://microsoftedge.microsoft.com/addons/Microsoft-Edge-Extensions-Home](https://microsoftedge.microsoft.com/addons/Microsoft-Edge-Extensions-Home)
 - [https://chrome.google.com/webstore/category/extensions](https://chrome.google.com/webstore/category/extensions)
 - [edge://extensions](edge://extensions)
## Install-VSCodeExtension
### Synopsis
Installs the given Visual Studio Code extension.
### Syntax
```powershell

Install-VSCodeExtension [-Extension] <String> [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Extension</nobr> |  | Extension ID or path to local .vsix file | true | true \(ByValue\) |  |
### Inputs
 - System.String You can pipe extension IDs or .vsix paths to this function.

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Install-VSCodeExtension -Extension "ms-python.python"
```
Download and install the "Python" extension by Microsoft.

**EXAMPLE 2**
```powershell
Install-VSCodeExtension -Extension ".\ms-python.python-2022.9.11611009.vsix"
```
Install the "Python" extension by Microsoft from a local .vsix file.

### Links

 - [https://marketplace.visualstudio.com/vscode](https://marketplace.visualstudio.com/vscode)
 - [https://code.visualstudio.com/](https://code.visualstudio.com/)
## Restart-Explorer
### Synopsis
Restarts explorer.exe.
### Syntax
```powershell

Restart-Explorer [<CommonParameters>]





```
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Restart-Explorer
```


### Links

 - [Stop-Process](#Stop-Process)
## Set-AutoAccentColor
### Synopsis
Enable or disables automatic selection of UI accent color.
### Syntax
```powershell

Set-AutoAccentColor [-Value] <Int32> [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Value</nobr> |  | 0 to disable, 1 to enable | true | false | 0 |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Links

 - [Enable-AutoAccentColor](#Enable-AutoAccentColor)
 - [Disable-AutoAccentColor](#Disable-AutoAccentColor)
## Set-CortanaButton
### Synopsis
Enables or disables Cortana button on Taskbar.
### Syntax
```powershell

Set-CortanaButton [-Value] <Int32> [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Value</nobr> |  | ShowCortanaButton value: 1 to enable, 0 to disable | true | false | 0 |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Set-CortanaButton -Value 1 -RestartExplorer
```
Enable Cortana button and restart explorer.exe so changes take effect immediately.

**EXAMPLE 2**
```powershell
Set-CortanaButton -Value 0
```
Disable Cortana button. Changes may not take effect until explorer.exe is restarted.

### Links

 - [Enable-CortanaButton](#Enable-CortanaButton)
 - [Disable-CortanaButton](#Disable-CortanaButton)
## Set-DesktopRecycleBin
### Synopsis
Show or hide Recycle Bin on Desktop.
### Syntax
```powershell

Set-DesktopRecycleBin [-Value] <Int32> [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Value</nobr> |  | 0 to show, 1 to hide | true | false | 0 |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Links

 - [Enable-DesktopRecycleBin](#Enable-DesktopRecycleBin)
 - [Disable-DesktopRecycleBin](#Disable-DesktopRecycleBin)
## Set-ExplorerQuickAccess
### Synopsis
Enables or disables the Quick Access menu in Explorer system-wide.
### Syntax
```powershell

Set-ExplorerQuickAccess [-Value] <Int32> [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Value</nobr> |  | HubMode value: 0 to enable Quick Access, 1 to disable | true | false | 0 |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Set-ExplorerQuickAccess -Value 1 -RestartExplorer
```
Disable Quick Access and restart explorer.exe so changes take effect immediately.

**EXAMPLE 2**
```powershell
Set-ExplorerQuickAccess -Value 0
```
Enable Quick Access. Changes may not take effect until explorer.exe is restarted.

### Links

 - [Enable-ExplorerQuickAccess](#Enable-ExplorerQuickAccess)
 - [Disable-ExplorerQuickAccess](#Disable-ExplorerQuickAccess)
## Set-InkWorkspaceButton
### Synopsis
Show or hide the Ink Workspace button on the Taskbar.
### Syntax
```powershell

Set-InkWorkspaceButton [-Value] <Int32> [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Value</nobr> |  | 1 to show Ink Workspace button, 0 to hide | true | false | 0 |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Set-InkWorkspaceButton -Value 0 -RestartExplorer
```
Hide Ink Workspace button and restart explorer.exe so changes take effect immediately.

**EXAMPLE 2**
```powershell
Set-InkWorkspaceButton -Value 1
```
Show Ink Workspace button. Changes may not take effect until explorer.exe is restarted.

### Links

 - [Enable-InkWorkspaceButton](#Enable-InkWorkspaceButton)
 - [Disable-InkWorkspaceButton](#Disable-InkWorkspaceButton)
## Set-NewsAndInterests
### Synopsis
Change the appearance of "News and Interests" on Taskbar.
### Syntax
```powershell

Set-NewsAndInterests [-Mode] <String> [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Mode</nobr> |  | News and Interest mode: IconAndText, Icon, or Off | true | false |  |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Note
Doesn't appear to work on 21H2.

### Examples
**EXAMPLE 1**
```powershell
Set-NewsAndInterests -Mode "IconAndText" -RestartExplorer
```
Show "News and Interests" icon and text and restart explorer.exe so changes take effect immediately.

**EXAMPLE 2**
```powershell
Set-NewsAndInterests -Mode "Icon" -RestartExplorer
```
Show "News and Interests" icon and restart explorer.exe so changes take effect immediately.

**EXAMPLE 3**
```powershell
Set-NewsAndInterests -Mode "Off"
```
Disable "News and Interests". Changes may not take effect until explorer.exe is restarted.

## Set-PersonlizeKeyValue
### Synopsis
Changes Personlization settings via Registry.
### Syntax
```powershell

Set-PersonlizeKeyValue [-Name] <String> [-Value] <Object> [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Name</nobr> |  | Name of value | true | false |  |
| <nobr>Value</nobr> |  | Value data | true | false |  |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Links

 - [Enable-AppLightTheme](#Enable-AppLightTheme)
 - [Enable-AppDarkTheme](#Enable-AppDarkTheme)
 - [Enable-SystemLightTheme](#Enable-SystemLightTheme)
 - [Enable-SystemDarkTheme](#Enable-SystemDarkTheme)
 - [Enable-TransparencyEffects](#Enable-TransparencyEffects)
 - [Disable-TransparencyEffects](#Disable-TransparencyEffects)
## Set-TaskbarSearch
### Synopsis
Change the appearance of the Taskbar Search feature.
### Syntax
```powershell

Set-TaskbarSearch [-Mode] <String> [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Mode</nobr> |  | Taskbar Search mode: Hidden, Icon, or Bar | true | false |  |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Set-TaskbarSearch -Mode "Hidden" -RestartExplorer
```
Hide search and restart explorer.exe so changes take effect immediately.

**EXAMPLE 2**
```powershell
Set-TaskbarSearch -Mode "Icon" -RestartExplorer
```
Show search icon on Taskbar and restart explorer.exe so changes take effect immediately.

**EXAMPLE 3**
```powershell
Set-TaskbarSearch -Mode "Bar"
```
Show search bar on Taskbar. Changes may not take effect until explorer.exe is restarted.

## Set-Wallpaper
### Synopsis
Applies a specified wallpaper to the current user's desktop
### Syntax
```powershell

Set-Wallpaper [-Image] <String> [[-Style] <String>] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Image</nobr> |  | Provide the exact path to the image | true | false |  |
| <nobr>Style</nobr> |  | Provide wallpaper style \(Example: Fill, Fit, Stretch, Tile, Center, or Span\) | false | false |  |
### Outputs
 - System.Void

### Note
Author: Jose Espitia Date: 2020-08-11

### Examples
**EXAMPLE 1**
```powershell
Set-Wallpaper -Image "C:\Wallpaper\Default.jpg"
```


**EXAMPLE 2**
```powershell
Set-Wallpaper -Image "C:\Wallpaper\Background.jpg" -Style Fit
```


### Links

 - [https://www.joseespitia.com/2017/09/15/set-wallpaper-powershell-function/](https://www.joseespitia.com/2017/09/15/set-wallpaper-powershell-function/)
## Set-WallpaperQuality
### Synopsis
Sets the quality of desktop wallpaper.
### Syntax
```powershell

Set-WallpaperQuality [-Quality] <Int32> [-RestartExplorer] [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Quality</nobr> |  | Wallpaper quality \(60 - 100\) | true | false | 0 |
| <nobr>RestartExplorer</nobr> |  | Restart explorer.exe to allow changes to take effect | false | false | False |
### Inputs
 - None

### Outputs
 - None

### Examples
**EXAMPLE 1**
```powershell
Set-WallpaperQuality -Value 100 -RestartExplorer
```
Set wallpaper quality to max value and restart explorer.exe so changes take effect immediately.

### Links

 - [https://www.howtogeek.com/277808/windows-10-compresses-your-wallpaper-but-you-can-make-them-high-quality-again/](https://www.howtogeek.com/277808/windows-10-compresses-your-wallpaper-but-you-can-make-them-high-quality-again/)
## Test-Command
### Synopsis
Determines if a command exists.
### Syntax
```powershell

Test-Command [-Name] <String> [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Name</nobr> |  | Command to check | true | true \(ByValue\) |  |
### Inputs
 - System.String You can pipe command names to this function.

### Outputs
 - System.Boolean

### Examples
**EXAMPLE 1**
```powershell
Test-Command -Name "foo"
```
Check if the command "foo" exists.

### Links

 - [Get-Command](#Get-Command)
## Test-EdgeExtensionUrl
### Synopsis
Determines whether an Edge extension URL is valid.
### Syntax
```powershell

Test-EdgeExtensionUrl [-Url] <String> [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Url</nobr> |  | URL to test | true | true \(ByValue\) |  |
### Inputs
 - System.String You can pipe extension URLs to this function.

### Outputs
 - System.Boolean

### Examples
**EXAMPLE 1**
```powershell
Test-EdgeExtensionUrl -Url "https://microsoftedge.microsoft.com/addons/detail/dark-reader/ifoakfbpdcdoeenechcleahebpibofpc"
```


### Links

 - [Test-Url](#Test-Url)
 - [https://microsoftedge.microsoft.com/addons/Microsoft-Edge-Extensions-Home](https://microsoftedge.microsoft.com/addons/Microsoft-Edge-Extensions-Home)
 - [https://chrome.google.com/webstore/category/extensions](https://chrome.google.com/webstore/category/extensions)
## Test-Url
### Synopsis
Determines whether a URL is reachable.
### Syntax
```powershell

Test-Url [-Url] <String> [<CommonParameters>]





```
### Parameters
| Name  | Alias  | Description | Required? | Pipeline Input | Default Value |
| - | - | - | - | - | - |
| <nobr>Url</nobr> |  | URL to test | true | true \(ByValue\) |  |
### Inputs
 - System.String You can pipe URLs to this function.

### Outputs
 - System.Boolean

### Examples
**EXAMPLE 1**
```powershell
Test-Url -Url "https://zombo.com"
```


## Test-Virtualization
### Synopsis
Determines if system supports virtualization.
### Syntax
```powershell

Test-Virtualization [<CommonParameters>]





```
### Inputs
 - None

### Outputs
 - System.Boolean

### Examples
**EXAMPLE 1**
```powershell
Test-Virtualization
```


