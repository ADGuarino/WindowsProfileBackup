

#Define Workstation and User
###################################################################################################################
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') 
$Computer = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the Computer Name', 'Export User Profile')

[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$Profile = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the Profile Username to Export', 'Export TEC Profile')

#Defining Remote Paths
###################################################################################################################
$Destination = # "Enter Destination Folder Here"

$HomePath = "\\$Computer\c$\Users\$Profile"
$RoamingPath = "\\$Computer\c$\Users\$Profile\AppData\Roaming\Microsoft"
$LocalPath = "\\$Computer\c$\Users\$Profile\\AppData\Local\Microsoft"
$TaskbarPath = "$RoamingPath\Internet Explorer\Quick Launch\User Pinned"
$WallpaperPath = "$RoamingPath\Windows"

#If User Profile Folder doesn't exist at $Destination, create it
###################################################################################################################
If (!(Test-Path -Path $Destination )) {
    New-Item -ItemType Directory -Path $Destination | Out-Null
    }

#Backup Mapped Drives
###################################################################################################################
$MappedDrives = Get-WmiObject win32_MappedLogicalDisk -ComputerName $Computer
Foreach ($Drive in $MappedDrives) {
            $NetUse = "Net Use"
            $Letter = $Drive.name
            $Path = $Drive.ProviderName
            $NetUse + " " + $Letter + " " + $Path | Add-Content "$Destination\Map.bat"
            }

#Backup $HomePath Items (Desktop, Downloads, Documents)
###################################################################################################################
Write-Progress -Activity "Backing up Desktop..."
Copy-Item -Path "$HomePath\Desktop" -Recurse  -ErrorAction SilentlyContinue -Destination $Destination -Force

Write-Progress -Activity "Backing up Downloads..."
Copy-Item -Path "$HomePath\Downloads" -Recurse -ErrorAction SilentlyContinue -Destination $Destination -Force

Write-Progress -Activity "Backing up Documents..."
Copy-Item -Path "$HomePath\Documents" -Recurse -ErrorAction SilentlyContinue -Destination $Destination -Force

#Backup $RoamingPath Items (Signatures, Dictionary, Templates, Sticky Notes)
###################################################################################################################
Write-Progress -Activity "Backing up Signatures..."
Copy-Item -Path "$RoamingPath\Signatures" -Recurse -ErrorAction SilentlyContinue -Destination $Destination -Force

Write-Progress -Activity "Backing up Dictionary..."
Copy-Item -Path "$RoamingPath\UProof" -Recurse -ErrorAction SilentlyContinue -Destination $Destination -Force

Write-Progress -Activity "Backing up Templates..."
Copy-Item -Path "$RoamingPath\Templates" -Recurse -ErrorAction SilentlyContinue -Destination $Destination -Force

Write-Progress -Activity "Backing up Sticky Notes..."
Copy-Item -Path "$RoamingPath\Sticky Notes" -Recurse -ErrorAction SilentlyContinue -Destination $Destination -Force

#Backup $LocalPath Items (OneNote, Forms)
###################################################################################################################
Write-Progress -Activity "Backing up OneNote Notebooks..."
Copy-Item -Path "$LocalPath\OneNote" -Recurse -ErrorAction SilentlyContinue -Destination $Destination -Force

Write-Progress -Activity "Backing up Forms..."
Copy-Item -Path "$LocalPath\Forms" -Recurse -ErrorAction SilentlyContinue -Destination $Destination -Force

#Backup Taskbar Pins
###################################################################################################################
Write-Progress -Activity "Backing up Taskbar Pins..."
Copy-Item -Path $TaskbarPath\Taskbar -Recurse -ErrorAction SilentlyContinue -Destination $Destination -Force

#Backup Wallpaper & Themes
###################################################################################################################
Write-Progress -Activity "Backing up Wallpaper..."
Copy-Item -Path $WallpaperPath\Themes -Recurse -ErrorAction SilentlyContinue -Destination $Destination -Force

#Backup Completed Dialog Box Popup
###################################################################################################################
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("$Profile has been successfully backed up",0,"Done",0x1)