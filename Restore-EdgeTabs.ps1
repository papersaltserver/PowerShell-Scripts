<#
.SYNOPSIS 
Helps to restore Microsoft Edge tabs
.DESCRIPTION
Based on https://stackoverflow.com/questions/34117660/how-to-restore-session-in-microsoft-edge answer
Step by step instruction to restore your Microsoft Edge tabs from previous session
1. Find backup of C:\Users\<YOUR_USERNAME>\AppData\Local\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\AC\MicrosoftEdge\User\Default\Recovery\Active
2. Run .\Restore-EdgeTabs.ps1 <PATH_TO_BACKUP>
Done! Script will do all other needed stuff for you
.PARAMETER backupDir
Specifies path to a directory contained {*}.dat files from a session you want to resotre
.EXAMPLE
.\Restore-EdgeTabs.ps1 c:\EdgeBackup
.NOTES
ATTENTION! All current Edge tabs will be closed, current session will be backed up to a subfolder within specified backup folder
Please try "previous versions" feature of metioned Edge folder if you have not backuped this folder manually
If you have no backup of mentioned Edge folder, it is not possible to restore Edge tabs.
Author: Dmitry Trukhanov
#>

param(
    [Parameter(Mandatory=$true,Position=1)]
    [string]$backupDir
)

$backupFiles = Get-ChildItem $backupDir -ErrorAction SilentlyContinue -File -Filter "{*}.dat"
if($backupFiles -eq $null){
    Write-Host "Can't access specified directory"
    exit 1
}
Get-Process MicrosoftEdge -ErrorAction SilentlyContinue| ForEach-Object{$_.Kill()}
Start-Sleep -Seconds 1
$edgePath = "$($env:LOCALAPPDATA)\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\AC\MicrosoftEdge\User\Default\Recovery\Active\"
$currentTime = Get-Date
$currentBackup = New-Item -ItemType Directory -Path "$backupDir\currentEdgeTabsBackup-$($currentTime.Year)-$($currentTime.Month)-$($currentTime.Day)-$($currentTime.Hour)-$($currentTime.Minute)-$($currentTime.Second)"
Copy-Item -Recurse "$edgePath\*" $currentBackup
$currentFiles = Get-ChildItem $edgePath
$currentFiles | ForEach-Object{$_.Delete()}
for($i=0;$i -lt $backupFiles.Count;$i++){
    Start-Process "microsoft-edge:about:blank"
}
Get-Process MicrosoftEdge -ErrorAction SilentlyContinue| ForEach-Object{$_.Kill()}
Start-Sleep -Seconds 1
$newNames = (Get-ChildItem -File -Filter "{*}.dat" -Path $edgePath).Name
for($i=0;$i -lt $backupFiles.Count;$i++){
    Copy-Item $backupFiles[$i].FullName "$edgePath\$($newNames[$i])" -Force
}
Start-Process "microsoft-edge:"
