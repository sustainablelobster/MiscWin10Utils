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

        .LINK
            https://boxstarter.org/installboxstarter#installing-from-the-web
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param ()

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

        .LINK
            https://chrome.google.com/webstore/category/extensions

        .LINK
            edge://extensions
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
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

            Returns $true if the URL is valid, else $false.

        .INPUTS
            System.String
                You can pipe extension URLs to this function.
        
        .OUTPUTS
            System.Boolean

        .EXAMPLE
            Test-EdgeExtensionUrl -Url "https://microsoftedge.microsoft.com/addons/detail/dark-reader/ifoakfbpdcdoeenechcleahebpibofpc"

        .LINK
            Test-Url
        
        .LINK
            https://microsoftedge.microsoft.com/addons/Microsoft-Edge-Extensions-Home

        .LINK
            https://chrome.google.com/webstore/category/extensions
    #>

    [CmdletBinding()]
    [OutputType([Boolean])]
    param (
        # URL to test
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [AllowEmptyString()]
        [String]
        $Url
    )

    process {
        $ExtUrlRegex = "https://(?:microsoftedge`.microsoft`.com/addons)|(?:chrome`.google`.com/webstore)/detail/.+?/[a-p]{32}"

        if ($Url -match $ExtUrlRegex) {
            Test-Url -Url $Url
        } else {
            $false
        }
    }
}


function Test-Url {
    <#
        .SYNOPSIS
            Determines whether a URL is reachable.

        .DESCRIPTION
            Determines whether a URL is reachable by attempting an HTTP HEAD request. Returns $true if the request
            is successful, else $false.

        .INPUTS
            System.String
                You can pipe URLs to this function.
        
        .OUTPUTS
            System.Boolean

        .EXAMPLE
            Test-Url -Url "https://zombo.com"
    #>

    [CmdletBinding()]
    [OutputType([Boolean])]
    param (
        # URL to test
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [AllowEmptyString()]
        [String]
        $Url
    )

    begin {
        $ProgressPreference = "SilentlyContinue"
    }

    process {
        $InvokeWebRequestParams = @{
            Uri = $Url
            Method = "Head"
            UseBasicParsing = $true
            DisableKeepAlive = $true
            ErrorAction = "Stop"
        }

        try {
            (Invoke-WebRequest @InvokeWebRequestParams).StatusCode -eq 200
        } catch {
            $false
        }
    }
}


function Install-VSCodeExtension {
    <#
        .SYNOPSIS
            Installs the given Visual Studio Code extension.

        .DESCRIPTION
            Installs the given Visual Studio Code extension from either extension ID or path to a local .vsix file.
            Visual Studio Code must be installed.

        .INPUTS
            System.String
                You can pipe extension IDs or .vsix paths to this function.
        
        .OUTPUTS
            None

        .EXAMPLE
            Install-VSCodeExtension -Extension "ms-python.python"

            Download and install the "Python" extension by Microsoft.

        .EXAMPLE
            Install-VSCodeExtension -Extension ".\ms-python.python-2022.9.11611009.vsix"

            Install the "Python" extension by Microsoft from a local .vsix file.

        .LINK
            https://marketplace.visualstudio.com/vscode

        .LINK
            https://code.visualstudio.com/
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Extension ID or path to local .vsix file
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [String]
        $Extension
    )

    begin {
        if (-not (Test-Command -Name "code")) {
            throw "Could not find 'code' command. Check if Visual Studio Code is installed."
        }

        $ErrorRecordType = [System.Management.Automation.ErrorRecord]
    }

    process {
        $CodeOutput = code --install-extension $Extension 2>&1
        if ($CodeOutput[1].GetType() -eq $ErrorRecordType -and $CodeOutput[1] -match "not found") {
            Write-Error $CodeOutput[1]
        } 
    }
}


function ConvertTo-WSLPath {
    <#
        .SYNOPSIS
            Converts a Windows path to its corresponding WSL path.

        .DESCRIPTION
            Converts the given Windows path to its corresponding WSL path using the "wslpath" command. The Windows
            Subsystem for Linux feature must be enabled.

        .INPUTS
            System.String
                You can pipe paths to this function.
        
        .OUTPUTS
            System.String

        .EXAMPLE
            ConvertTo-WSLPath -Path "." -Full

            Get the full WSL path of the current directory.

        .EXAMPLE
            ConvertTo-WSLPath -Path ".\dir\subdir\file.bin"

            Get the relative WSL path of the given Windows path.

        .EXAMPLE
            (Get-ChildItem -Path "." -Recurse).FullName | ConvertTo-WSLPath

            Get WSL paths for every item in the current directory.

        .LINK
            https://docs.microsoft.com/en-us/windows/wsl/about
    #>

    [CmdletBinding()]
    [OutputType([String])]
    param (
        # Windows path to convert
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [String]
        $Path,
        # Use the Full switch to get get the full path (default is relative)
        [Parameter(Mandatory = $false)]
        [Switch]
        $Full
    )

    begin {
        if (-not (Test-Command -Name "wsl")) {
            throw "Could not find 'wsl' command. Check if Windows Subsystem for Linux is enabled."
        }
    }

    process {
        $Path = $Path.Replace("\", "\\")

        if ($Full) {
            wsl wslpath -a $Path
        } else {
            wsl wslpath $Path
        }
    }
}


function Test-Command {
    <#
        .SYNOPSIS
            Determines if a command exists.

        .DESCRIPTION
            Determines if a command exists by checking if Get-Command can find it. Returns $true if command is
            found, else $false.

        .INPUTS
            System.String
                You can pipe command names to this function.
        
        .OUTPUTS
            System.Boolean

        .EXAMPLE
            Test-Command -Name "foo"

            Check if the command "foo" exists.

        .LINK
            Get-Command
    #>

    [CmdletBinding()]
    [OutputType([Boolean])]
    param (
        # Command to check
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [String]
        $Name
    )

    begin {
        $ErrorActionPreference = "Stop"
    }
    
    process {
        try {
            $null -ne (Get-Command -Name $Name)
        } catch {
            $false
        }
    }
}


function Test-Virtualization {
    <#
        .SYNOPSIS
            Determines if system supports virtualization.

        .DESCRIPTION
            Determines if system supports virtualization by checking if virtualization is enabled in firmware or
            if Hyper-V's "vmcompute" service is enabled.

        .INPUTS
            None
        
        .OUTPUTS
            System.Boolean

        .EXAMPLE
            Test-Virtualization
    #>

    [CmdletBinding()]
    [OutputType([Boolean])]
    param ()

    begin {
        $ErrorActionPreference = "Stop"
    }

    process {
        $VirtFWEnabled = (Get-CimInstance -ClassName "win32_processor").VirtualizationFirmwareEnabled

        try {
            $VMComputeFound = $null -ne (Get-Service -Name "vmcompute")
        } catch {
            $VMComputeFound = $false
        }

        $VirtFWEnabled -or $VMComputeFound
    }
}


function Enable-ExplorerQuickAccess {
    <#
        .SYNOPSIS
            Enables the Quick Access menu in Explorer system-wide.

        .DESCRIPTION
            Enables the Quick Access menu in Explorer system-wide. Changes may not take effect until Explorer is
            restarted. Must be run as Administrator.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Enable-ExplorerQuickAccess -RestartExplorer

            Enable Quick Access and restart explorer.exe so changes take effect immediately.

        .LINK
            Disable-ExplorerQuickAccess
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-ExplorerQuickAccess -Value 0 -RestartExplorer:$RestartExplorer
    }
}


function Disable-ExplorerQuickAccess {
    <#
        .SYNOPSIS
            Disables the Quick Access menu in Explorer system-wide.

        .DESCRIPTION
            Disables the Quick Access menu in Explorer system-wide. Changes may not take effect until Explorer is
            restarted. Must be run as Administrator.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Disable-ExplorerQuickAccess -RestartExplorer

            Disable Quick Access and restart explorer.exe so changes take effect immediately.

        .LINK
            Enable-ExplorerQuickAccess
    #>
    
    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-ExplorerQuickAccess -Value 1 -RestartExplorer:$RestartExplorer
    }
}


function Set-ExplorerQuickAccess {
    <#
        .SYNOPSIS
            Enables or disables the Quick Access menu in Explorer system-wide.

        .DESCRIPTION
            Enables or disables the Quick Access menu in Explorer system-wide. Changes may not take effect until 
            Explorer is restarted. Must be run as Administrator.

            Not intended to be used directly; use the functions in the related links instead.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Set-ExplorerQuickAccess -Value 1 -RestartExplorer

            Disable Quick Access and restart explorer.exe so changes take effect immediately.

        .EXAMPLE
            Set-ExplorerQuickAccess -Value 0

            Enable Quick Access. Changes may not take effect until explorer.exe is restarted.

        .LINK
            Enable-ExplorerQuickAccess

        .LINK
            Disable-ExplorerQuickAccess
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # HubMode value: 0 to enable Quick Access, 1 to disable
        [Parameter(Mandatory = $true)]
        [ValidateRange(0, 1)]
        [Int]
        $Value,
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        $SetItemPropertyParams = @{
            Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"
            Name = "HubMode"
            Value = $Value
            Force = $true
        }

        Set-ItemProperty @SetItemPropertyParams

        if ($RestartExplorer) {
            Restart-Explorer
        }
    }
}


function Restart-Explorer {
    <#
        .SYNOPSIS
            Restarts explorer.exe.

        .DESCRIPTION
            Restarts explorer.exe.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Restart-Explorer

        .LINK
            Stop-Process
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param ()

    process {
        Stop-Process -Name "explorer"
    }
}


function Enable-CortanaButton {
    <#
        .SYNOPSIS
            Enables Cortana button on Taskbar.

        .DESCRIPTION
            Enables the Cortana button on the Taskbar. Changes may not take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Enable-CortanaButton -RestartExplorer

            Enable Cortana button and restart explorer.exe so changes take effect immediately.

        .LINK
            Disable-CortanaButton
    #>
    
    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-CortanaButton -Value 1 -RestartExplorer:$RestartExplorer
    }
}


function Disable-CortanaButton {
    <#
        .SYNOPSIS
            Disables Cortana button on Taskbar.

        .DESCRIPTION
            Disables the Cortana button on the Taskbar. Changes may not take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Disable-CortanaButton -RestartExplorer

            Disable Cortana button and restart explorer.exe so changes take effect immediately.

        .LINK
            Enable-CortanaButton
    #>
    
    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-CortanaButton -Value 0 -RestartExplorer:$RestartExplorer
    }
}


function Set-CortanaButton {
    <#
        .SYNOPSIS
            Enables or disables Cortana button on Taskbar.

        .DESCRIPTION
            Enables or disabels the Cortana button on the Taskbar. Changes may not take effect until Explorer is 
            restarted.

            Not intended to be used directly; use the functions in the related links instead.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Set-CortanaButton -Value 1 -RestartExplorer

            Enable Cortana button and restart explorer.exe so changes take effect immediately.

        .EXAMPLE
            Set-CortanaButton -Value 0

            Disable Cortana button. Changes may not take effect until explorer.exe is restarted.

        .LINK
            Enable-CortanaButton

        .LINK
            Disable-CortanaButton
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # ShowCortanaButton value: 1 to enable, 0 to disable
        [Parameter(Mandatory = $true)]
        [ValidateRange(0, 1)]
        [Int] 
        $Value,
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch] 
        $RestartExplorer
    )

    process {
        $ExplorerAdvancedKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $ExplorerAdvancedKey -Name "ShowCortanaButton" -Value $Value -Force

        if ($RestartExplorer) {
            Restart-Explorer
        }
    }
}


function Set-TaskbarSearch {
    <#
        .SYNOPSIS
            Change the appearance of the Taskbar Search feature.

        .DESCRIPTION
            Change the appearance of the Taskbar Search feature to:
                "Hidden" - Disable search bar/icon
                "Icon" - Show search icon on Taskbar
                "Bar" - Show search bar on Taskbar

            Changes may not take effect until explorer.exe is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Set-TaskbarSearch -Mode "Hidden" -RestartExplorer

            Hide search and restart explorer.exe so changes take effect immediately.

        .EXAMPLE
            Set-TaskbarSearch -Mode "Icon" -RestartExplorer

            Show search icon on Taskbar and restart explorer.exe so changes take effect immediately.

        .EXAMPLE
            Set-TaskbarSearch -Mode "Bar"

            Show search bar on Taskbar. Changes may not take effect until explorer.exe is restarted.
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Taskbar Search mode: Hidden, Icon, or Bar
        [Parameter(Mandatory = $true)]
        [ValidateSet("Hidden", "Icon", "Bar")]
        [String]
        $Mode,
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        $Value = switch ($Mode) {
            "Hidden" { 0 }
            "Icon" { 1 }
            Default { 2 }   # Bar
        }
    
        $SearchKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search"
        Set-ItemProperty -Path $SearchKey -Name "SearchboxTaskbarMode" -Value $Value -Force
    
        if ($RestartExplorer) {
            Restart-Explorer
        }
    }
}


function Set-NewsAndInterests {
    <#
        .SYNOPSIS
            Change the appearance of "News and Interests" on Taskbar.

        .DESCRIPTION
            Change the appearance of "News and Interests" on the Taskbar to:
                "IconAndText" - Show icon and text
                "Icon" - Show icon only
                "Off" - Disable "News and Interests"

            Changes may not take effect until explorer.exe is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Set-NewsAndInterests -Mode "IconAndText" -RestartExplorer

            Show "News and Interests" icon and text and restart explorer.exe so changes take effect immediately.

        .EXAMPLE
            Set-NewsAndInterests -Mode "Icon" -RestartExplorer

            Show "News and Interests" icon and restart explorer.exe so changes take effect immediately.

        .EXAMPLE
            Set-NewsAndInterests -Mode "Off"

            Disable "News and Interests". Changes may not take effect until explorer.exe is restarted.

        .NOTES
            Doesn't appear to work on 21H2.
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # News and Interest mode: IconAndText, Icon, or Off
        [Parameter(Mandatory = $true)]
        [ValidateSet("IconAndText", "Icon", "Off")]
        [String]
        $Mode,
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        $Value = switch ($Mode) {
            "Off" { 2 }
            "Icon" { 1 }
            Default { 0 }   # IconAndText
        }

        $FeedsKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Feeds"
        Set-ItemProperty -Path $FeedsKey -Name "ShellFeedsTaskbarViewMode" -Value $Value -Force

        if ($RestartExplorer) {
            Restart-Explorer
        }
    }
}


function Enable-InkWorkspaceButton {
    <#
        .SYNOPSIS
            Show the Ink Workspace button on the Taskbar.

        .DESCRIPTION
            Show the Ink Workspace button on the Taskbar. Changes may not take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Enable-InkWorkspaceButton -RestartExplorer

            Show Ink Workspace button and restart explorer.exe so changes take effect immediately.

        .LINK
            Disable-InkWorkspaceButton
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-InkWorkspaceButton -Value 1 -RestartExplorer:$RestartExplorer
    }
}


function Disable-InkWorkspaceButton {
    <#
        .SYNOPSIS
            Hide the Ink Workspace button on the Taskbar.

        .DESCRIPTION
            Hide the Ink Workspace button on the Taskbar. Changes may not take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Disable-InkWorkspaceButton -RestartExplorer

            Hide Ink Workspace button and restart explorer.exe so changes take effect immediately.

        .LINK
            Enable-InkWorkspaceButton
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-InkWorkspaceButton -Value 0 -RestartExplorer:$RestartExplorer
    }
}


function Set-InkWorkspaceButton {
    <#
        .SYNOPSIS
            Show or hide the Ink Workspace button on the Taskbar.

        .DESCRIPTION
            Show or Hide the Ink Workspace button on the Taskbar.
                Value = 0 hides the button
                Value = 1 shows the button
                
            Changes may not take effect until Explorer is restarted.

            Not intended to be used directly; use the functions in the related links instead.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Set-InkWorkspaceButton -Value 0 -RestartExplorer

            Hide Ink Workspace button and restart explorer.exe so changes take effect immediately.

        .EXAMPLE
            Set-InkWorkspaceButton -Value 1

            Show Ink Workspace button. Changes may not take effect until explorer.exe is restarted.

        .LINK
            Enable-InkWorkspaceButton

        .LINK
            Disable-InkWorkspaceButton
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # 1 to show Ink Workspace button, 0 to hide
        [Parameter(Mandatory = $true)]
        [ValidateRange(0, 1)]
        [Int]
        $Value,
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        $PenWorkspaceKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\PenWorkspace"
        Set-ItemProperty -Path $PenWorkspaceKey -Name "PenWorkspaceButtonDesiredVisibility" -Value $Value -Force

        if ($RestartExplorer) {
            Restart-Explorer
        }
    }
}


function Set-WallpaperQuality {
    <#
        .SYNOPSIS
            Sets the quality of desktop wallpaper.

        .DESCRIPTION
            Sets the quality of desktop wallpaper anywhere from 60% to 100%. By default, Windows 10 compresses
            wallpapers to 85% of their original quality. Changes may not take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Set-WallpaperQuality -Value 100 -RestartExplorer

            Set wallpaper quality to max value and restart explorer.exe so changes take effect immediately.

        .LINK
            https://www.howtogeek.com/277808/windows-10-compresses-your-wallpaper-but-you-can-make-them-high-quality-again/
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Wallpaper quality (60 - 100)
        [Parameter(Mandatory = $true)]
        [ValidateRange(60, 100)]
        [Int] 
        $Quality,
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        $DesktopKey = "HKCU:\SOFTWARE\Control Panel\Desktop"
        Set-ItemProperty -Path $DesktopKey -Name "JPEGImportQuality" -Value $Quality -Force

        if ($RestartExplorer) {
            Restart-Explorer
        }
    }
}


function Enable-AutoAccentColor {
    <#
        .SYNOPSIS
            Let Windows pick UI accent color automatically.

        .DESCRIPTION
            Let Windows pick UI accent color automatically based on your wallpaper. Changes may not take effect
            until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Enable-AutoAccentColor -RestartExplorer

            Enable automatic accents and restart explorer.exe so changes take effect immediately.

        .LINK
            Disable-AutoAccentColor
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-AutoAccentColor -Value 1 -RestartExplorer:$RestartExplorer
    }
}


function Disable-AutoAccentColor {
    <#
        .SYNOPSIS
            Disables automatic selection of UI accent color.

        .DESCRIPTION
            Prevents Windows from picking UI accent color automatically based on your wallpaper. Changes may not
            take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Disable-AutoAccentColor -RestartExplorer

            Disable automatic accents and restart explorer.exe so changes take effect immediately.

        .LINK
            Enable-AutoAccentColor
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-AutoAccentColor -Value 0 -RestartExplorer:$RestartExplorer
    }
}


function Set-AutoAccentColor {
    <#
        .SYNOPSIS
            Enable or disables automatic selection of UI accent color.

        .DESCRIPTION
            Allow or prevent Windows from picking UI accent color automatically based on your wallpaper. 
                Value = 0 to disable
                Value = 1 to enable

            Changes may not take effect until Explorer is restarted.

            Not intended to be used directly; use the functions in the related links instead.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .LINK
            Enable-AutoAccentColor

        .LINK
            Disable-AutoAccentColor
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # 0 to disable, 1 to enable
        [Parameter(Mandatory = $true)]
        [ValidateRange(0, 1)]
        [Int]
        $Value,
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        $DesktopKey = "HKCU:\Control Panel\Desktop"
        Set-ItemProperty -Path $DesktopKey -Name "AutoColorization" -Value $Value -Force

        if ($RestartExplorer) {
            Restart-Explorer
        }
    }
}


function Enable-AppLightTheme {
    <#
        .SYNOPSIS
            Set application theme to light mode.

        .DESCRIPTION
            Set application theme to light mode. Changes may not take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Enable-AppLightTheme -RestartExplorer

            Set app theme to light mode and restart explorer.exe so changes take effect immediately.

        .LINK
            Enable-AppDarkTheme

        .LINK
            Enable-SystemLightTheme

        .LINK
            Enable-SystemDarkTheme
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-PersonlizeKeyValue -Name "AppsUseLightTheme" -Value 1 -RestartExplorer:$RestartExplorer
    }
}


function Enable-AppDarkTheme {
    <#
        .SYNOPSIS
            Set application theme to dark mode.

        .DESCRIPTION
            Set application theme to dark mode. Changes may not take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Enable-AppDarkTheme -RestartExplorer

            Set app theme to dark mode and restart explorer.exe so changes take effect immediately.

        .LINK
            Enable-AppLightTheme

        .LINK
            Enable-SystemLightTheme

        .LINK
            Enable-SystemDarkTheme
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-PersonlizeKeyValue -Name "AppsUseLightTheme" -Value 0 -RestartExplorer:$RestartExplorer
    }
}


function Enable-SystemLightTheme {
    <#
        .SYNOPSIS
            Set system theme to light mode.

        .DESCRIPTION
            Set system theme to light mode. Changes may not take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Enable-SystemLightTheme -RestartExplorer

            Set system theme to light mode and restart explorer.exe so changes take effect immediately.

        .LINK
            Enable-SystemDarkTheme

        .LINK
            Enable-AppLightTheme

        .LINK
            Enable-AppDarkTheme
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-PersonlizeKeyValue -Name "SystemUsesLightTheme" -Value 1 -RestartExplorer:$RestartExplorer
    }
}


function Enable-SystemDarkTheme {
    <#
        .SYNOPSIS
            Set system theme to dark mode.

        .DESCRIPTION
            Set system theme to dark mode. Changes may not take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Enable-SystemDarkTheme -RestartExplorer

            Set system theme to dark mode and restart explorer.exe so changes take effect immediately.

        .LINK
            Enable-SystemLightTheme

        .LINK
            Enable-AppLightTheme

        .LINK
            Enable-AppDarkTheme
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-PersonlizeKeyValue -Name "SystemUsesLightTheme" -Value 0 -RestartExplorer:$RestartExplorer
    }
}


function Enable-TransparencyEffects {
    <#
        .SYNOPSIS
            Enable system UI transparency effects.

        .DESCRIPTION
            Enable system UI transparency effects. Changes may not take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Enable-TransparencyEffects -RestartExplorer

            Enable system UI transparency effects and restart explorer.exe so changes take effect immediately.

        .LINK
            Disable-TransparencyEffects
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-PersonlizeKeyValue -Name "EnableTransparency" -Value 1 -RestartExplorer:$RestartExplorer
    }
}


function Disable-TransparencyEffects {
    <#
        .SYNOPSIS
            Disable system UI transparency effects.

        .DESCRIPTION
            Disable system UI transparency effects. Changes may not take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Disable-TransparencyEffects -RestartExplorer

            Disable system UI transparency effects and restart explorer.exe so changes take effect immediately.

        .LINK
            Enable-TransparencyEffects
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-PersonlizeKeyValue -Name "EnableTransparency" -Value 0 -RestartExplorer:$RestartExplorer
    }
}


function Set-PersonlizeKeyValue {
    <#
        .SYNOPSIS
            Changes Personlization settings via Registry.

        .DESCRIPTION
            Changes Personlization settings by modifying values of the Registry key 
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize".

            Not intended to be used directly; use the functions in the related links instead.

        .INPUTS
            None

        .OUTPUTS
            None

        .LINK
            Enable-AppLightTheme

        .LINK
            Enable-AppDarkTheme

        .LINK
            Enable-SystemLightTheme

        .LINK
            Enable-SystemDarkTheme

        .LINK
            Enable-TransparencyEffects

        .LINK
            Disable-TransparencyEffects
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Name of value
        [Parameter(Mandatory = $true)]
        [String]
        $Name,
        # Value data
        [Parameter(Mandatory = $true)]
        $Value,
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        $PersonalizeKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        Set-ItemProperty -Path $PersonalizeKey -Name $Name -Value $Value -Force

        if ($RestartExplorer) {
            Restart-Explorer
        }
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

        .EXAMPLE
            Set-Wallpaper -Image "C:\Wallpaper\Background.jpg" -Style Fit

        .NOTES
            Author: Jose Espitia
            Date: 2020-08-11

        .LINK
            https://www.joseespitia.com/2017/09/15/set-wallpaper-powershell-function/
    #>
    
    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Provide path to image
        [parameter(Mandatory=$true)]
        [ValidateScript({ Test-Path -Path $_ })]
        [String]
        $Image,
        # Provide wallpaper style that you would like applied
        [parameter(Mandatory=$false)]
        [ValidateSet('Fill', 'Fit', 'Stretch', 'Tile', 'Center', 'Span')]
        [String] 
        $Style
    )
     
    process {
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
}


function Enable-DesktopRecycleBin {
    <#
        .SYNOPSIS
            Show Recycle Bin on Desktop.

        .DESCRIPTION
            Show Recycle Bin on Desktop. Changes may not take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Enable-DesktopRecycleBin -RestartExplorer

            Show Recycle Bin on Desktop and restart explorer.exe so changes take effect immediately.

        .LINK
            Disable-DesktopRecycleBin
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-DesktopRecycleBin -Value 0 -RestartExplorer:$RestartExplorer
    }
}


function Disable-DesktopRecycleBin {
    <#
        .SYNOPSIS
            Hide Recycle Bin on Desktop.

        .DESCRIPTION
            Hide Recycle Bin on Desktop. Changes may not take effect until Explorer is restarted.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .EXAMPLE
            Disable-DesktopRecycleBin -RestartExplorer

            Hide Recycle Bin on Desktop and restart explorer.exe so changes take effect immediately.

        .LINK
            Enable-DesktopRecycleBin
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    process {
        Set-DesktopRecycleBin -Value 1 -RestartExplorer:$RestartExplorer
    }
}


function Set-DesktopRecycleBin {
    <#
        .SYNOPSIS
            Show or hide Recycle Bin on Desktop.

        .DESCRIPTION
            Show or hide Recycle Bin on Desktop.
                Value = 0 to show
                Value = 1 to hide

            Changes may not take effect until Explorer is restarted.

            Not intended to be used directly; use the functions in the related links instead.

        .INPUTS
            None
        
        .OUTPUTS
            None

        .LINK
            Enable-DesktopRecycleBin

        .LINK
            Disable-DesktopRecycleBin
    #>

    [CmdletBinding()]
    [OutputType([Void])]
    param (
        # 0 to show, 1 to hide
        [Parameter(Mandatory = $true)]
        [ValidateRange(0, 1)]
        [Int]
        $Value,
        # Restart explorer.exe to allow changes to take effect
        [Parameter(Mandatory = $false)]
        [Switch]
        $RestartExplorer
    )

    $NewStartPanelKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
    $RecycleBinGuid = "{645FF040-5081-101B-9F08-00AA002F954E}"
    Set-ItemProperty -Path $NewStartPanelKey -Name $RecycleBinGuid -Value $Value

    if ($RestartExplorer) {
        Restart-Explorer
    }
}
