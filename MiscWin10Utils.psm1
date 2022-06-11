function Install-Boxstarter {
    [System.Net.ServicePointManager]::SecurityProtocol = 
            [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://boxstarter.org/bootstrapper.ps1'))
    Get-Boxstarter -Force
}


function Install-EdgeExtension {
    [CmdletBinding()]
    
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [string] $ExtensionURL
    )

    begin {
        $ExtensionsRootKey = "HKCU:\SOFTWARE\Microsoft\Edge\Extensions"
    }

    process {
        $ExtensionID = $ExtensionURL.Split("/")[-1]
        $ExtensionKey = $ExtensionsRootKey + "\" + $ExtensionID

        if ($ExtensionURL -match "chrome`.google`.com") {
            $UpdateURL = "https://clients2.google.com/service/update2/crx"
        } else {
            $UpdateURL = "https://edge.microsoft.com/extensionwebstorebase/v1/crx"
        }

        New-Item -Path $ExtensionKey -Force
        Set-ItemProperty -Path $ExtensionKey -Name "update_url" -Value $UpdateURL -Force
    }
    
}


function Install-VSCodeExtension {
    [CmdletBinding()]

    param(
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [string] $Extension
    )

    process {
        code --install-extension $Extension
    }
}


function ConvertTo-WSLPath {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,
        [Parameter(Mandatory = $false)]
        [switch] $Full
    )

    if ($Full) {
        $Path = (Get-Item -Path $Path).FullName
    }

    wsl wslpath $Path.Replace("\", "\\")
}


function Test-Command {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Name
    )

    $SavedErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "stop"
    $CommandInstalled = $true

    try {
        Get-Command -Name $Name | Out-Null
    } catch {
        $CommandInstalled = $false
    }

    $ErrorActionPreference = $SavedErrorActionPreference
    $CommandInstalled
}


function Test-Virtualization {
    (Get-CimInstance -ClassName "win32_processor").VirtualizationFirmwareEnabled
}


function Enable-WindowsFeaturesFromJson {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Json,
        [Parameter(Mandatory = $false)]
        [switch] $Restart = $false
    )

    $Features = (Get-Content -Path $Json | ConvertFrom-Json)
    $VirtualizationEnabled = Test-Virtualization

    foreach ($Feature in $Features) {
        if ($Feature.requiresVirtualization -and -not $VirtualizationEnabled) {
            continue
        }

        Enable-WindowsOptionalFeature -FeatureName $Feature.name -Online -All -NoRestart
    }

    if ($Restart) {
        Restart-Computer -Force
    }
}


function Set-ExplorerQuickAccess {
    param(
        [Parameter(Mandatory = $false)]
        [switch] $Enable = $false,
        [Parameter(Mandatory = $false)]
        [switch] $Disable = $false,
        [Parameter(Mandatory = $false)]
        [switch] $RestartExplorer
    )

    $ExplorerKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"
    
    $HubModeValue = 0 
    if ($Disable) {
        $HubModeValue = 1
    } 
    
    Set-ItemProperty -Path $ExplorerKey -Name "HubMode" -Value $HubModeValue -Force

    if ($RestartExplorer) {
        Restart-Explorer
    }
}


function Restart-Explorer {
    Stop-Process -Name "explorer"
}


function New-ThisPCFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Name,
        [Parameter(Mandatory = $true)]
        [string] $Path,
        [Parameter(Mandatory = $false)]
        [string] $Infotip = "",
        [Parameter(Mandatory = $false)]
        [string] $Icon = "$env:SystemRoot\system32\shell32.dll,3"
    )

    $Guid = (New-Guid).Guid
    $TemplateKey = "HKLM:\SOFTWARE\Classes\CLSID\{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}" # Videos
    $CLSIDKey = "HKCU:\SOFTWARE\Classes\CLSID\{$Guid}"
    $NamespaceKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{$GUID}"
    $InstanceCLSID = "{0AFACED1-E828-11D1-9187-B532F1E9575D}" # Folder shortcut

    Copy-Item -Path $TemplateKey -Destination $CLSIDKey -Recurse -Force
    Set-ItemProperty -Path $CLSIDKey -Name "(Default)" -Value $Name -Force
    Set-ItemProperty -Path $CLSIDKey -Name "Infotip" -Value $Infotip -Force
    Set-ItemProperty -Path $CLSIDKey -Name "CreatedBy" -Value "MyWin10" -Force
    Set-ItemProperty -Path "$CLSIDKey\DefaultIcon" -Name "(Default)" -Value $Icon -Force
    Set-ItemProperty -Path "$CLSIDKey\Instance" -Name "CLSID" -Value $InstanceCLSID -Force
    Set-ItemProperty -Path "$CLSIDKey\Instance\InitPropertyBag" -Name "Target" -Value $Path -Type ExpandString `
            -Force
    Set-ItemProperty -Path "$CLSIDKey\Instance\InitPropertyBag" -Name "Attributes" -Value 0x15 -Force
    Remove-ItemProperty -Path "$CLSIDKey\Instance\InitPropertyBag" -Name "TargetKnownFolder" -Force
    Remove-ItemProperty -Path "$CLSIDKey\ShellFolder" -Name "FolderValueFlags" -Force
    Remove-ItemProperty -Path "$CLSIDKey\ShellFolder" -Name "SortOrderIndex" -Force
    New-Item -Path $NamespaceKey -Force
}


function Set-CortanaButton {
    param(
        [Parameter(Mandatory = $false)]
        [switch] $Enable = $false,
        [Parameter(Mandatory = $false)]
        [switch] $Disable = $false,
        [Parameter(Mandatory = $false)]
        [switch] $RestartExplorer
    )

    $ExplorerAdvancedKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

    $ShowCortanaButtonValue = 1
    if ($Disable) {
        $ShowCortanaButtonValue = 0
    }

    Set-ItemProperty -Path $ExplorerAdvancedKey -Name "ShowCortanaButton" -Value $ShowCortanaButtonValue -Force

    if ($RestartExplorer) {
        Restart-Explorer
    }
}


function Set-TaskbarSearch {
    param(
        [Parameter(Mandatory = $false)]
        [switch] $ShowBar = $false,
        [Parameter(Mandatory = $false)]
        [switch] $ShowIcon = $false,
        [Parameter(Mandatory = $false)]
        [switch] $Disable,
        [Parameter(Mandatory = $false)]
        [switch] $RestartExplorer
    )

    $SearchKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search"

    $TaskbarModeValue = 2
    if ($Disable) {
        $TaskbarModeValue = 0
    } elseif ($ShowIcon) {
        $TaskbarModeValue = 1
    }

    Set-ItemProperty -Path $SearchKey -Name "SearchboxTaskbarMode" -Value $TaskbarModeValue -Force

    if ($RestartExplorer) {
        Restart-Explorer
    }
}


function Set-NewsAndInterests {
    param(
        [Parameter(Mandatory = $false)]
        [switch] $ShowIconAndText = $false,
        [Parameter(Mandatory = $false)]
        [switch] $ShowIconOnly = $false,
        [Parameter(Mandatory = $false)]
        [switch] $Disable = $false,
        [Parameter(Mandatory = $false)]
        [switch] $RestartExplorer
    )

    $FeedsKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Feeds"

    $ViewModeValue = 0
    if ($ShowIconOnly) {
        $ViewModeValue = 1
    } elseif ($Disable) {
        $ViewModeValue = 2
    }

    Set-ItemProperty -Path $FeedsKey -Name "ShellFeedsTaskbarViewMode" -Value $ViewModeValue -Force

    if ($RestartExplorer) {
        Restart-Explorer
    }
}


function Set-InkWorkspaceButton {
    param(
        [Parameter(Mandatory = $false)]
        [switch] $Enable = $false,
        [Parameter(Mandatory = $false)]
        [switch] $Disable,
        [Parameter(Mandatory = $false)]
        [switch] $RestartExplorer
    )

    $PenWorkspaceKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\PenWorkspace"

    $VisibilityValue = 1
    if ($Disable) {
        $VisibilityValue = 0
    }

    Set-ItemProperty -Path $PenWorkspaceKey -Name "PenWorkspaceButtonDesiredVisibility" -Value $VisibilityValue `
            -Force

    if ($RestartExplorer) {
        Restart-Explorer
    }
}


function Set-WallpaperQuality {
    param(
        [Parameter(Mandatory = $true)]
        [int] $Quality,
        [Parameter(Mandatory = $false)]
        [switch] $RestartExplorer
    )

    $DesktopKey = "HKCU:\SOFTWARE\Control Panel\Desktop"

    if ($Quality -lt 60) {
        $Quality = 60
    } elseif ($Quality -gt 100) {
        $Quality = 100
    }

    Set-ItemProperty -Path $DesktopKey -Name "JPEGImportQuality" -Value $Quality -Force

    if ($RestartExplorer) {
        Restart-Explorer
    }
}


function Set-AutoAccentColor {
    param(
        [Parameter(Mandatory = $false)]
        [switch] $Enable = $false,
        [Parameter(Mandatory = $false)]
        [switch] $Disable = $false,
        [Parameter(Mandatory = $false)]
        [switch] $RestartExplorer
    )

    $DesktopKey = "HKCU:\Control Panel\Desktop"

    $AutoColorizationValue = 0
    if ($Enable) {
        $AutoColorizationValue = 1
    }

    Set-ItemProperty -Path $DesktopKey -Name "AutoColorization" -Value $AutoColorizationValue -Force

    if ($RestartExplorer) {
        Restart-Explorer
    }
}


function Set-ColorSettings {
    param(
        [Parameter(Mandatory = $false)]
        [switch] $EnableAppLightTheme = $false,
        [Parameter(Mandatory = $false)]
        [switch] $EnableAppDarkTheme = $false,
        [Parameter(Mandatory = $false)]
        [switch] $EnableSystemLightTheme = $false,
        [Parameter(Mandatory = $false)]
        [switch] $EnableSystemDarkTheme = $false,
        [Parameter(Mandatory = $false)]
        [switch] $EnableTransparancyEffects = $false,
        [Parameter(Mandatory = $false)]
        [switch] $DisableTransparancyEffects = $false,
        [Parameter(Mandatory = $false)]
        [switch] $EnableAutoAccentColor = $false,
        [Parameter(Mandatory = $false)]
        [switch] $DisableAutoAccentColor = $false,
        [Parameter(Mandatory = $false)]
        [switch] $RestartExplorer = $false
    )

    $PersonalizeKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"

    if ($EnableAppLightTheme) {
        Set-ItemProperty -Path $PersonalizeKey -Name "AppsUseLightTheme" -Value 1 -Force
    } elseif ($EnableAppDarkTheme) {
        Set-ItemProperty -Path $PersonalizeKey -Name "AppsUseLightTheme" -Value 0 -Force
    }

    if ($EnableSystemLightTheme) {
        Set-ItemProperty -Path $PersonalizeKey -Name "SystemUsesLightTheme" -Value 1 -Force
    } elseif ($EnableSystemDarkTheme) {
        Set-ItemProperty -Path $PersonalizeKey -Name "SystemUsesLightTheme" -Value 0 -Force
    }

    if ($EnableTransparancyEffects) {
        Set-ItemProperty -Path $PersonalizeKey -Name "EnableTransparancy" -Value 1 -Force
    } elseif ($DisableTransparancyEffects) {
        Set-ItemProperty -Path $PersonalizeKey -Name "EnableTransparancy" -Value 0 -Force
    }

    if ($EnableAutoAccentColor) {
        Set-AutoAccentColor -Enable
    } elseif ($DisableAutoAccentColor) {
        Set-AutoAccentColor -Disable
    }

    if ($RestartExplorer) {
        Restart-Explorer
    }
}


function Test-Ubuntu {
    if (-not (Test-Command -Name "ubuntu")) {
        $false
    } else {
        ubuntu run cat /dev/null
        $?
    }
}


function Set-Wallpaper {
    <#
        .SYNOPSIS
            Applies a specified wallpaper to the current user's desktop

        .PARAMETER Image
            Provide the exact path to the image

        .PARAMETER Style
            Provide wallpaper style (Example: Fill, Fit, Stretch, Tile, Center, or Span)

        .EXAMPLE
            Set-Wallpaper -Image "C:\Wallpaper\Default.jpg"
            Set-Wallpaper -Image "C:\Wallpaper\Background.jpg" -Style Fit

        .NOTES
            Author: Jose Espitia
            Date: 2020-08-11
            Page: https://www.joseespitia.com/2017/09/15/set-wallpaper-powershell-function/
    #>
     
    param (
        [parameter(Mandatory=$True)]
        # Provide path to image
        [string] $Image,
        # Provide wallpaper style that you would like applied
        [parameter(Mandatory=$False)]
        [ValidateSet('Fill', 'Fit', 'Stretch', 'Tile', 'Center', 'Span')]
        [string] $Style
    )
     
    $WallpaperStyle = switch ($Style) {
        "Fill" {"10"}
        "Fit" {"6"}
        "Stretch" {"2"}
        "Tile" {"0"}
        "Center" {"0"}
        "Span" {"22"}
    }
     
    if ($Style -eq "Tile") {
     
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String `
                -Value $WallpaperStyle -Force
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 1 `
                -Force
     
    } else {
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String `
                -Value $WallpaperStyle -Force
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 0 `
                -Force
    }
     
    Add-Type -TypeDefinition @" 
        using System; 
        using System.Runtime.InteropServices;        
        public class Params { 
            [DllImport("User32.dll",CharSet=CharSet.Unicode)] 
            public static extern int SystemParametersInfo (Int32 uAction, Int32 uParam, String lpvParam, 
                    Int32 fuWinIni);
        }
"@ 
      
    $SPI_SETDESKWALLPAPER = 0x0014
    $UpdateIniFile = 0x01
    $SendChangeEvent = 0x02
    $fWinIni = $UpdateIniFile -bor $SendChangeEvent
    [Params]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $Image, $fWinIni)
}


function Set-DesktopRecycleBin {
    param(
        [Parameter(Mandatory = $false)]
        [switch] $Enable = $false,
        [Parameter(Mandatory = $false)]
        [switch] $Disable = $false,
        [Parameter(Mandatory = $false)]
        [switch] $RestartExplorer = $false
    )

    $NewStartPanelKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
    $RecycleBinGuid = "{645FF040-5081-101B-9F08-00AA002F954E}"

    if ($Enable) {
        Remove-ItemProperty -Path $NewStartPanelKey -Name $RecycleBinGuid
    } elseif ($Disable) {
        Set-ItemProperty -Path $NewStartPanelKey -Name $RecycleBinGuid -Value 1
    }

    if ($RestartExplorer) {
        Restart-Explorer
    }
}
