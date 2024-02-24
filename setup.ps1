function Write-ScriptMessage {
  param (
      [String]$Message
  )

  Write-Host "[SETUP SCRIPT] $Message" -ForegroundColor Green
}

function Set-RegistryValue {
  param (
    [String]$Path,
    [String]$Name,
    $Value,
    [String]$Type
  )

  $key = Get-Item -Path $Path -ErrorAction SilentlyContinue
  if ($null -eq $key) {
    New-Item -Path $Path -ItemType Key -Force
  }

  $reg = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
  if ($null -eq $reg) {
    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force
  } else {
    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Force
  }
}

function Invoke-Code {
  param(
      [string]$Code,
      [bool]$RunAs = $false,
      [bool]$ExitNoError = $false
  )

  $arguments = ""
  if ($ExitNoError) {
    $arguments += " -ErrorAction SilentlyContinue"
  }
  if ($RunAs) {
    $arguments += " -Verb RunAs"
  }
  $arguments += " -NoExit"

  $scriptBlock = [scriptblock]::Create($Code)

  Start-Process powershell.exe -ArgumentList "$arguments -Command `"$scriptBlock`"" -PassThru
}

function Wait-AndCloseWhenReady {
  param(
      [string]$ProgramName
  )
  do {
      Start-Sleep -Seconds 1
      $process = Get-Process -Name $ProgramName -ErrorAction SilentlyContinue
  } until ($process -ne $null)

  if ($process -ne $null) {
      $process | Stop-Process -Force
      Write-ScriptMessage "$ProgramName is running, stopping it."
  } else {
      Write-ScriptMessage "It seems that $ProgramName is not running."
  }
}

function Install-ApplicationFromURL {
  param(
      [string]$ApplicationName,
      [string]$DownloadURL,
      [bool]$MSIX = $false
  )
  $downloadPath = "$env:TEMP\$ApplicationName.exe"
  try {
      Invoke-WebRequest -Uri $DownloadURL -OutFile $downloadPath -ErrorAction Stop
      Write-ScriptMessage "Download of $ApplicationName completed."
  } catch {
      Write-ScriptMessage "An error occurred while downloading $ApplicationName : $_"
      return
  }
  try {
    if ($MSIX) {
      Add-AppxPackage -Path $downloadPath -ErrorAction Stop
    } else {
      Start-Process -FilePath $downloadPath -Wait -ErrorAction Stop
    }
    Write-ScriptMessage "$ApplicationName has been successfully installed."
  } catch {
    Write-ScriptMessage "An error occurred while installing $ApplicationName : $_"
  }
  Remove-Item -Path $downloadPath -ErrorAction SilentlyContinue
}

function Unpin-AllFromTaskbar {
  $shell = New-Object -ComObject "Shell.Application"
  $taskbar = $shell.NameSpace('shell:::{0}' -f (0x1))
  if ($taskbar -eq $null) {
    Write-ScriptMessage "Failed to retrieve the Taskbar."
      return
  }

  $taskbarItems = $taskbar.Items()
  if ($taskbarItems -eq $null) {
    Write-ScriptMessage "Failed to retrieve items from the Taskbar."
      return
  }

  foreach ($item in $taskbarItems) {
      $taskbar.InvokeVerb("Unpin from taskbar", $item)
  }
  Write-ScriptMessage "All items have been unpinned from the taskbar."
}

function Pin-ToTaskbar {
  param (
      [string]$ProgramPath
  )
  $shell = New-Object -ComObject "Shell.Application"
  $taskbar = $shell.NameSpace('shell:::{0}' -f (0x1))
  $folderItem = $shell.Namespace((Get-Item $ProgramPath).DirectoryName).ParseName((Get-Item $ProgramPath).Name)
  $taskbar.InvokeVerb("Pin to taskbar", $folderItem)
  Write-ScriptMessage "$ProgramPath has been pinned to the taskbar."
}

function Pin-MicrosoftProgramToTaskbar {
  param (
      [string]$ProgramName
  )

  try {
    $shell = New-Object -ComObject "Shell.Application"
    $taskbar = $shell.Namespace("shell:::{0}" -f (0xa))

    $startMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu"
    $programPath = Get-ChildItem -Path $startMenuPath -Recurse -Include "$ProgramName.lnk" -ErrorAction Stop | Select-Object -First 1 -ExpandProperty FullName

    if ($programPath) {
        $folderItem = $shell.Namespace((Get-Item $programPath).DirectoryName).ParseName((Get-Item $programPath).Name)
        $taskbar.InvokeVerb("Pin to taskbar", $folderItem)
        Write-ScriptMessage "$ProgramName has been pinned to the taskbar."
    } else {
      Write-ScriptMessage "Program '$ProgramName' not found in the Start menu."
    }
  } catch {
    Write-ScriptMessage "An error occurred while pinning $ProgramName to the taskbar: $_"
  }
}

function Create-FileOnDesktop {
    param(
        [string]$FileName,
        [string]$FileContent
    )

    $desktopPath = [System.Environment]::GetFolderPath('Desktop')
    $filePath = Join-Path -Path $desktopPath -ChildPath $FileName

    try {
        Set-Content -Path $filePath -Value $FileContent
        Write-ScriptMessage "Le fichier $FileName a été créé sur le bureau avec succès."
    } catch {
      Write-ScriptMessage "Une erreur s'est produite lors de la création du fichier $FileName : $_"
    }
}

if (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
  Write-ScriptMessage "This script not requires administrator privileges. Starting without it."
  Exit
}

Start-Process powershell.exe -ArgumentList "-NoProfile -File `"restore-point.ps1`" -NoExit" -Verb RunAs

Write-ScriptMessage "PART 1 : Fixing Windows 11"
Write-ScriptMessage "Disabling Mouse acceleration"
Set-RegistryValue -Path "HKCU:Control Panel\Mouse" -Name "MouseThreshold1" -Value "0" -Type "String"
Set-RegistryValue -Path "HKCU:Control Panel\Mouse" -Name "MouseThreshold2" -Value "0" -Type "String"
Set-RegistryValue -Path "HKCU:Control Panel\Mouse" -Name "MouseSpeed" -Value "0" -Type "String"

Write-ScriptMessage "Disabling Recycle bin on all drives"
Get-ChildItem "HKCU:Software\Microsoft\Windows\CurrentVersion\Explorer\BitBucket\Volume" |
Foreach-Object { Set-RegistryValue -Path "HKCU:Software\Microsoft\Windows\CurrentVersion\Explorer\BitBucket\Volume\$(Split-Path $_ -Leaf)" -Name "NukeOnDelete" -Value 1 -Type "DWord" }

Write-ScriptMessage "Removing Taskbar default icons"
Set-RegistryValue -Path "HKCU:Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTaskViewButton" -Value 0 -Type "DWord"
Set-RegistryValue -Path "HKCU:Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarMn" -Value 0 -Type "DWord"
Set-RegistryValue -Path "HKCU:Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 0 -Type "DWord"

Write-ScriptMessage "Enabling dark mode"
Get-ChildItem "HKCU:Software\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops\Desktops" |
Foreach-Object { Set-RegistryValue -Path "HKCU:Software\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops\Desktops\$(Split-Path $_ -Leaf)" -Name "Wallpaper" -Value "C:\Windows\web\wallpaper\Windows\img19.jpg" -Type "String" }
Set-RegistryValue -Path "HKCU:Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers" -Name "BackgroundHistoryPath0" -Value "C:\Windows\web\wallpaper\Windows\img19.jpg" -Type "String"
Set-RegistryValue -Path "HKCU:Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers" -Name "BackgroundHistoryPath1" -Value "C:\Windows\web\wallpaper\Windows\img0.jpg" -Type "String"
Set-RegistryValue -Path "HKCU:Software\Microsoft\Windows\CurrentVersion\Themes" -Name "CurrentTheme" -Value "C:\Windows\resources\Themes\dark.theme" -Type "String"
Set-RegistryValue -Path "HKCU:Software\Microsoft\Windows\CurrentVersion\Themes" -Name "CurreThemeMRUntTheme" -Value "C:\Windows\resources\Themes\dark.theme;C:\Windows\resources\Themes\aero.theme;" -Type "String"
Set-RegistryValue -Path "HKCU:Software\Microsoft\Windows\CurrentVersion\Themes\HighContrast" -Name "Pre-High Contrast Scheme" -Value "C:\Windows\resources\Themes\dark.theme" -Type "String"
Set-RegistryValue -Path "HKCU:Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0 -Type "DWord"
Set-RegistryValue -Path "HKCU:Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 0 -Type "DWord"

Write-ScriptMessage "Removing desktop shortcuts"
$desktop = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
Get-ChildItem $desktop\*.lnk | ForEach-Object {
  Write-ScriptMessage "Removing desktop shortcut $_.FullName"
  Remove-Item -Path $_.FullName
}

Write-ScriptMessage "Part 2 : Installation"
if (-not (Test-Path $env:USERPROFILE\scoop)) {
  Invoke-RestMethod get.scoop.sh | Invoke-Expression
}
Write-ScriptMessage "Installing scoop buckets"

function Test-ScoopBucket {
  param (
      [string]$BucketName
  )
  $BucketsDir = Join-Path $env:USERPROFILE "scoop/buckets"
  $BucketExists = Test-Path (Join-Path $BucketsDir $BucketName)
  return $BucketExists
}

$scoopBuckets = @(
  "extras",
  "games",
  "nerd-fonts",
  "nonportable",
  "versions"
)
foreach ($scoopBucket in $scoopBuckets) {
  if (-not (Test-ScoopBucket $scoopBucket)) {
    Write-ScriptMessage "Adding scoop bucket $scoopBucket"
    scoop bucket add $scoopBucket
  } else {
    Write-ScriptMessage "Scoop bucket $scoopBucket already exists"
  }
}

Write-ScriptMessage "Installing scoop packages"
$packageList = scoop list
function Test-ScoopPackage {
  param (
        [string]$PackageName
    )
    $Package = $PackageName -replace "^[^/]+/", ""
    $PackageExists = $packageList | Select-String -Pattern $Package
    return [bool]$PackageExists
}
$scoopPackages = @(
  'main/python',
  'main/nodejs-lts',
  'main/git',
  'main/maven',
  'main/pwsh',
  'extras/opera',
  'extras/notion',
  'extras/telegram',
  'extras/termius',
  'extras/winrar',
  'extras/everything-lite',
  'extras/jetbrains-toolbox',
  'extras/vscode',
  'extras/discord',
  'games/epic-games-launcher',
  'games/goggalaxy',
  'games/steam',
  'games/battlenet',
  'extras/gog-galaxy-plugin-downloader',
  'nonportable/office-365-apps-np',
  'versions/ubisoftconnect',
  'nerd-fonts/CascadiaCode-NF-Mono -g'
)
foreach ($scoopPackage in $scoopPackages) {
  if (-not (Test-ScoopPackage $scoopPackage)) {
    Write-ScriptMessage "Installing scoop package $scoopPackage"
    scoop install $scoopPackage
  } else {
    Write-ScriptMessage "Scoop package $scoopPackage already installed but will be updated"
    scoop update $scoopPackage
  }
}
Start-Process pwsh -Verb runAs -Args "-ExecutionPolicy Bypass $($MyInvocation.Line) -SecondStep"
Start-Process DISM -Args "/online /disable-feature /featurename:WindowsMediaPlayer"

Write-ScriptMessage "Installing Vencord"
Start-Process -FilePath "c:\Users\antoi\AppData\Local\Discord\Update.exe" -ArgumentList "--processStart Discord.exe"
Wait-AndCloseWhenReady "discord"
Invoke-WebRequest 'https://raw.githubusercontent.com/Vencord/Installer/main/install.ps1' -UseBasicParsing | Invoke-Expression

Write-ScriptMessage "Installing WSL"
Invoke-Code "Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux -NoRestart" -RunAs $true
Invoke-WebRequest -Uri "https://aka.ms/wslubuntu2004" -OutFile "$env:USERPROFILE\Ubuntu.appx" -UseBasicParsing
Add-AppxPackage "$env:USERPROFILE\Ubuntu.appx"
Remove-Item "$env:USERPROFILE\Ubuntu.appx"

Install-ApplicationFromURL "TradingView" "https://tvd-packages.tradingview.com/stable/latest/win32/TradingView.msix" -MSIX $true
Install-ApplicationFromURL "Quantower" "https://updates.quantower.com/Quantower/x64/latest/Quantower.exe"
Install-ApplicationFromURL "1password" "https://downloads.1password.com/win/1PasswordSetup-latest.exe"
Install-ApplicationFromURL "SteelSeries GG" "https://engine.steelseriescdn.com/SteelSeriesGG57.0.0Setup.exe"
Install-ApplicationFromURL "NZXT CAM" "https://nzxt-app.nzxt.com/NZXT-CAM-Setup.exe"

Write-ScriptMessage "Part 3: Post installation"
Write-ScriptMessage "Copying Git config"
Copy-Item -Force .\.gitconfig $env:USERPROFILE

Write-ScriptMessage "Setting up Powershell Core"
Write-ScriptMessage "Copying profile"
Copy-Item -Force -Recurse .\.config\powershell $env:USERPROFILE\.config
New-Item -Force -ItemType Directory -Path $env:USERPROFILE\Documents\PowerShell
Write-Output ". `$env:USERPROFILE\.config\powershell\user_profile.ps1" | Out-File -FilePath $env:USERPROFILE\Documents\PowerShell\Microsoft.PowerShell_profile.ps1 -Encoding ASCII -Append
Write-ScriptMessage "Installing Starship"
winget install --accept-source-agreements --id Starship.Starship
Write-ScriptMessage "Installing Terminal-Icons"
Install-Module -Scope CurrentUser Terminal-Icons -Force

Write-ScriptMessage "Setting up Windows Terminal"
Write-ScriptMessage "Copying Windows Terminal settings"
Copy-Item -Force .\windows_terminal\settings.json $env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState
Write-ScriptMessage "Setting Windows Terminal as default terminal application"
Set-RegistryValue -Path "HKCU:Console\%%Startup" -Name "DelegationConsole" -Value "{2EACA947-7F5F-4CFA-BA87-8F7FBEEFBE69}" -Type "String"
Set-RegistryValue -Path "HKCU:Console\%%Startup" -Name "DelegationTerminal" -Value "{E12CFF52-A866-4C77-9A90-F570A7AA2C6B}" -Type "String"

Write-ScriptMessage "Setting VSCode"
reg import "$env:USERPROFILE\scoop\apps\vscode\current\install-associations.reg"
Write-ScriptMessage "Copying up VSCode settings"
New-Item -Force -ItemType Directory -Path $env:USERPROFILE\scoop\persist\vscode\data\user-data\User\
Copy-Item -Force .\vscode\settings.json $env:USERPROFILE\scoop\persist\vscode\data\user-data\User\

Write-ScriptMessage "Installing VSCode extensions"
$code_extensions = @(
  "tobiasalthoff.atom-material-theme",
  "ms-vscode.cpptools",
  "ms-vscode.cpptools-extension-pack",
  "ms-vscode.cpptools-themes",
  "ms-vscode.makefile-tools",
  "ms-vscode.powershell",
  "ms-azuretools.vscode-docker",
  "ms-vscode-remote.remote-wsl",
  "ben.epiheader",
  "github.vscode-github-actions",
  "github.copilot",
  "github.copilot-chat",
  "tal7aouy.rainbow-bracket",
  "shardulm94.trailing-spaces"
)
foreach ($extension in $code_extensions) {
  code --install-extension $extension
}

Write-ScriptMessage "Setting up Taskbar"
Unpin-AllFromTaskbar
Pin-ToTaskbar -ProgramPath "C:\Users\antoi\AppData\Local\Programs\Opera\launcher.exe"
Pin-ToTaskbar -ProgramPath "C:\Users\antoi\AppData\Local\Discord\Update.exe --processStart Discord.exe"
Pin-ToTaskbar -ProgramPath "C:\Users\antoi\scoop\apps\goggalaxy\current\GalaxyClient.exe"
Pin-MicrosoftProgramToTaskbar -ProgramName "TradingView"
Pin-ToTaskbar -ProgramPath "C:\Users\antoi\Quantower\TradingPlatform\v1.137.14\Starter.exe"
Pin-ToTaskbar -ProgramPath "C:\Users\antoi\scoop\apps\notion\current\Notion.exe"
Pin-ToTaskbar -ProgramPath "C:\Users\antoi\AppData\Local\1Password\app\8\1Password.exe"

Write-ScriptMessage "Setting up Gog Galaxy"
gog-plugins-downloader.exe -p steam,battlenet

Write-ScriptMessage "Part 4: Installing drivers"
Install-ApplicationFromURL "AMD Radeon Software" "https://drivers.amd.com/drivers/installer/23.40/whql/amd-software-adrenalin-edition-24.1.1-combined-minimalsetup-240122_web.exe"
Install-ApplicationFromURL "DriversCloud" "https://dcdrivers.driverscloud.com/applis/DriversCloudx64_12_0_18.exe"

$fileContent = @"
# Welcome to your new computer!

TODO:
- [ ] Add games on Epic Games Launcher
- [ ] Add games on Steam
- [ ] Add games on Battle.net
- [ ] Add games on Ubisoft Connect
- [ ] Add games on GOG Galaxy
- [ ] Start OneDrive sync
- [ ] Update drivers with DriversCloud
- [ ] Fix ssh with 1Password

## Software installed
- [x] Opera
- [x] Notion
- [x] Telegram
- [x] Termius
- [x] WinRAR
- [x] Everything Lite
- [x] JetBrains Toolbox
- [x] Visual Studio Code
- [x] Epic Games Launcher
- [x] GOG Galaxy
- [x] Steam
- [x] Battle.net
- [x] GOG Galaxy Plugin Downloader
- [x] Office 365 Apps
- [x] Ubisoft Connect
- [x] Cascadia Code Nerd Font
- [x] Python
- [x] Node.js LTS
- [x] Git

## Software to configure
- [ ] Vencord

## Opera extensions
- [ ] uBlock Origin
- [ ] Coupert
- [ ] MAL-Sync
- [ ] MyEpitech
- [ ] 1Password extension
"@


Create-FileOnDesktop -FileName "README.md" -FileContent $fileContent

Write-ScriptMessage "You can now reboot, press any key to continue"
Read-Host
