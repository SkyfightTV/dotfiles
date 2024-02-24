param(
  [switch]$SecondStep,
  [switch]$KeepOneDrive,
  [switch]$GitConfig,
  [string]$SSHFolder,
  [string]$GPGKey,
  [switch]$FirefoxExtensions,
  [switch]$VSCode,
  [switch]$WSL
)

function Write-ScriptMessage {
  param (
      [String]$Message
  )

  Write-Host "[SETUP SCRIPT] $Message" -ForegroundColor Green
}

function Invoke-Code {
  param(
      [string]$Code
  )

  $scriptBlock = [scriptblock]::Create($Code)
  Start-Process powershell.exe -ArgumentList "-NoProfile -NoExit -Command $scriptBlock" -PassThru
}

function Launch-AndCloseWhenReady {
  param(
      [string]$ProgramName
  )

  Start-Process $ProgramName

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
      [string]$DownloadURL
  )

  Invoke-Code "
  $downloadPath = "$env:TEMP\$ApplicationName.exe"
  try {
      Invoke-WebRequest -Uri $DownloadURL -OutFile $downloadPath -ErrorAction Stop
      Write-Host "Download of $ApplicationName completed."
  } catch {
      Write-Host "An error occurred while downloading $ApplicationName : $_"
      return
  }
  try {
      Start-Process -FilePath $downloadPath -Wait -ErrorAction Stop
      Write-Host "$ApplicationName has been successfully installed."
  } catch {
      Write-Host "An error occurred while installing $ApplicationName : $_"
  }
  Remove-Item -Path $downloadPath -ErrorAction SilentlyContinue"
}

function Unpin-AllFromTaskbar {
  $shell = New-Object -ComObject "Shell.Application"
  $taskbar = $shell.Namespace("shell:::{0}" -f (0xa))
  $taskbarItems = $taskbar.Items()

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
  $taskbar = $shell.Namespace("shell:::{0}" -f (0xa))
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
        Write-Host "$ProgramName has been pinned to the taskbar."
    } else {
        Write-Host "Program '$ProgramName' not found in the Start menu."
    }
  } catch {
      Write-Host "An error occurred while pinning $ProgramName to the taskbar: $_"
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
        Write-Host "Le fichier $FileName a été créé sur le bureau avec succès."
    } catch {
        Write-Host "Une erreur s'est produite lors de la création du fichier $FileName : $_"
    }
}

Checkpoint-Computer -Description "[SETUP SCRIPT] Before installation" -RestorePointType "MODIFY_SETTINGS"

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
Get-ChildItem $env:USERPROFILE\Desktop\*.lnk | ForEach-Object {
  Write-ScriptMessage "Removing desktop shortcut $_.FullName"
  Remove-Item -Path $_.FullName
}

Write-ScriptMessage "Part 2 : Installation"
Invoke-RestMethod get.scoop.sh | Invoke-Expression

Write-ScriptMessage "Installing scoop buckets"
$scoopBuckets = @(
  "extras",
  "games",
  "nerd-fonts",
  "nonportable",
  "versions"
)
Invoke-Code "
foreach ($scoopBucket in $scoopBuckets) {
  Write-ScriptMessage "Adding scoop bucket $scoopBucket"
  scoop bucket add $scoopBucket
}"

Write-ScriptMessage "Installing scoop packages"
$scoopPackages = @(
  'main/python',
  'main/nodejs-lts',
  'main/aws',
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
  'games/epic-games-launcher',
  'games/goggalaxy',
  'games/steam',
  'games/battlenet',
  'extras/gog-galaxy-plugin-downloader',
  'nonportable/office-365-apps-np',
  'versions/ubisoftconnect',
  'nerd-fonts/CascadiaCode-NF-Mono',
)
Invoke-Code "
foreach ($scoopPackage in $scoopPackages) {
  Write-ScriptMessage "Installing scoop package $scoopPackage"
  scoop install $scoopPackage
}"
Start-Process pwsh -Verb runAs -Args "-ExecutionPolicy Bypass $($MyInvocation.Line) -SecondStep"
Start-Process DISM -Args "/online /disable-feature /featurename:WindowsMediaPlayer"

Write-ScriptMessage "Installing Vencord"
Invoke-Code "
Launch-AndCloseWhenReady "discord"
Invoke-WebRequest 'https://raw.githubusercontent.com/Vencord/Installer/main/install.ps1' -UseBasicParsing | Invoke-Expression"

Write-ScriptMessage "Installing WSL"
wsl.exe --install

Install-ApplicationFromURL "TradingView" "https://tvd-packages.tradingview.com/stable/latest/win32/TradingView.msix"
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
Copy-Item -Force .\vscode\keybindings.json $env:USERPROFILE\scoop\persist\vscode\data\user-data\User\

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
  "shardulm94.trailing-spaces",
)
Invoke-Code "
foreach ($extension in $code_extensions) {
  code --install-extension $extension
}"

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
Invoke-Code "gog-plugins-downloader.exe -p steam,battlenet"


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
