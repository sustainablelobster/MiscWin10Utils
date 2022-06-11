function Install-Boxstarter {
    <#
        .SYNOPSIS
            Installs Boxstarter.

        .DESCRIPTION
            Downloads and installs Boxstarter, a tool "to automate the installation of software and create 
            repeatable, scripted Windows environments."

        .INPUTS
            None

        .OUTPUTS
            None

        .EXAMPLE
            Install-Boxstarter

        .NOTES
            Function body is based on Boxstarter's "installing from the web" instructions.

        .LINK
            https://boxstarter.org/
            https://boxstarter.org/installboxstarter#installing-from-the-web
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param()

    process {
        $SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        [System.Net.ServicePointManager]::SecurityProtocol = $SecurityProtocol
        $BootstrapperUrl = "https://boxstarter.org/bootstrapper.ps1"
        Invoke-Expression -Command ((New-Object System.Net.WebClient).DownloadString($BootstrapperUrl))
        Get-Boxstarter -Force
    }
}


function Install-EdgeExtension {
    <#
        .SYNOPSIS
            Installs the given Edge extension.

        .DESCRIPTION
            Installs a Microsoft Edge extension from the given Edge Add-ons / Chrome Web Store URL. The extension 
            must be manually activated in Edge's Extensions panel (edge://extensions). "Allow extensions from other
            stores" must be enabled in the Extensions panel to allow extensions from the Chrome Web Store.

        .INPUTS
            System.String
                You can pipe extension URLs to this function.
        
        .OUTPUTS
            None

        .EXAMPLE
            Install-EdgeExtension -ExtensionUrl "https://microsoftedge.microsoft.com/addons/detail/dark-reader/ifoakfbpdcdoeenechcleahebpibofpc"

            Install the "Dark Reader" extension from the Edge Add-ons store.

        .EXAMPLE
            Install-EdgeExtension -ExtensionUrl "https://chrome.google.com/webstore/detail/picture-in-picture-extens/hkgfoiooedgoejojocmhlaklaeopbecg"

            Install the "Picture-in-Picture Extension (by Google)" extension from the Chrome Web Store.

        .LINK
            https://microsoftedge.microsoft.com/addons/Microsoft-Edge-Extensions-Home
            https://chrome.google.com/webstore/category/extensions
            edge://extensions
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param(
        # URL of the extension's Edge Add-ons page or Chrome Web Store page.
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [ValidateScript({ Test-EdgeExtensionUrl -Url $_ })]
        [String]
        $ExtensionUrl
    )

    begin {
        $ErrorActionPreference = "Stop"
    }

    process {
        $ExtensionID = $ExtensionUrl.Split("/")[-1]
        if ($ExtensionID.Contains("?")) {
            $ExtensionID = $ExtensionID.Substring(0, $ExtensionID.LastIndexOf("?"))
        }
        

        $ExtensionKey = Join-Path -Path "HKCU:\SOFTWARE\Microsoft\Edge\Extensions" -ChildPath $ExtensionID

        $UpdateURL = switch -Wildcard ($ExtensionUrl) {
            "https://chrome.google.com/webstore/detail/*" { "https://clients2.google.com/service/update2/crx" }
            Default { "https://edge.microsoft.com/extensionwebstorebase/v1/crx" }
        }

        New-Item -Path $ExtensionKey -Force | Out-Null
        Set-ItemProperty -Path $ExtensionKey -Name "update_url" -Value $UpdateURL -Force
    }
}


function Test-EdgeExtensionUrl {
    <#
        .SYNOPSIS
            Determines whether an Edge extension URL is valid.

        .DESCRIPTION
            Determines whether an Edge extension URL is valid by:
                1. Testing if the URL matches the pattern of a an Edge or Chrome extension page.
                2. Testing if the URL is reachable (using Test-Url)

            Returns $True if the URL is valid, else $False.

        .INPUTS
            System.String
                You can pipe extension URLs to this function.
        
        .OUTPUTS
            System.Boolean

        .EXAMPLE
            Test-EdgeExtensionUrl -Url "https://microsoftedge.microsoft.com/addons/detail/dark-reader/ifoakfbpdcdoeenechcleahebpibofpc"

        .LINK
            Test-Url
            https://microsoftedge.microsoft.com/addons/Microsoft-Edge-Extensions-Home
            https://chrome.google.com/webstore/category/extensions
    #>

    [CmdletBinding()]
    [OutputType([Boolean])]
    param(
        # URL to test
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [AllowEmptyString()]
        [String] $Url
    )

    process {
        $ExtUrlRegex = "https://(?:microsoftedge`.microsoft`.com/addons)|(?:chrome`.google`.com/webstore)/detail/.+?/[a-p]{32}"

        if ($Url -match $ExtUrlRegex) {
            Test-Url -Url $Url
        } else {
            $False
        }
    }
}


function Test-Url {
    [CmdletBinding()]
    [OutputType([Boolean])]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [AllowEmptyString()]
        [String] $Url
    )

    begin {
        $ProgressPreference = "SilentlyContinue"
    }

    process {
        $InvokeWebRequestParams = @{
            Uri = $Url
            Method = "Head"
            UseBasicParsing = $True
            DisableKeepAlive = $True
            ErrorAction = "Stop"
        }

        try {
            (Invoke-WebRequest @InvokeWebRequestParams).StatusCode -eq 200
        } catch {
            $False
        }
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
