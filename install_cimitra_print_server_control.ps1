# IGNORE THIS ERROR! IGNORE THIS ERROR! JUST A POWERSHELL THING THAT HAPPENS ON THE FIRST LINE OF A POWERSHELL SCRIPT 

# Cimitra Cimitra Windows Print Server Control Install Script
# Author: Tay Kratzer tay@cimitra.com
# 9/21/2021

Write-Output "IGNORE THIS ERROR! IGNORE THIS ERROR! JUST A POWERSHELL THING THAT HAPPENS ON THE FIRST LINE OF A POWERSHELL SCRIPT"

function CHECK_ADMIN_LEVEL{

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
[Security.Principal.WindowsBuiltInRole] "Administrator")) {
Write-Output ""
Write-Warning "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
Write-Output ""
exit 1
}

}
CHECK_ADMIN_LEVEL


$global:INSTALLATION_DIRECTORY = "C:\cimitra\scripts\cimitra_win_print_server_admin"
 
write-output ""
write-output "START: INSTALLING - Cimitra Windows Print Server Control Practice"
write-output "-----------------------------------------------------------------"


if ($args[0]) { 
    $global:INSTALLATION_DIRECTORY = $args[0]
}

try{
    New-Item -ItemType Directory -Force -Path $INSTALLATION_DIRECTORY 2>&1 | out-null
}catch{}

$theResult = $?

if (!($theResult)){
    Write-Output "Error: Could Not Create Installation Directory: $INSTALLATION_DIRECTORY"
    exit 1
}

try{
Set-Location -Path $INSTALLATION_DIRECTORY
}catch{
Write-Output ""
Write-Output "Error: Cannot access directory: $INSTALLATION_DIRECTORY"
Write-Output ""
exit 1
}


$CurrentPath = Get-Location
$CurrentPath= $CurrentPath.Path

$CIMITRA_SCRIPT_DOWNLOAD = "https://raw.githubusercontent.com/cimitrasoftware/windows_print_server/main/cimitra_win_print_server_admin.ps1"
$CIMITRA_MERGE_DOWNLOAD = "https://raw.githubusercontent.com/cimitrasoftware/windows_print_server/main/CreateAction.ps1"


$CIMITRA_SCRIPT_DOWNLOAD_OUT_FILE = "$INSTALLATION_DIRECTORY\cimitra_win_print_server_admin.ps1"
$CIMITRA_MERGE_DOWNLOAD_OUT_FILE = "$INSTALLATION_DIRECTORY\CreateAction.ps1"


$ThisScript = $MyInvocation.MyCommand.Name



$global:runSetup = $true


if($Verbose){
    Write-Output ""
    Write-Output "Downloading File: $CIMITRA_SCRIPT_DOWNLOAD"
}else{
    Write-Output ""
    Write-Output "Downloading Script File From GitHub"
}

try{
    $RESULTS = Invoke-WebRequest $CIMITRA_SCRIPT_DOWNLOAD -OutFile $CIMITRA_SCRIPT_DOWNLOAD_OUT_FILE -UseBasicParsing 2>&1 | out-null
}catch{}

$theResult = $?

if (!$theResult){
    Write-Output "Error: Could Not Download The File: $CIMITRA_SCRIPT_DOWNLOAD"
    exit 1
}

if($Verbose){
    Write-Output ""
    Write-Output "Downloading File: $CIMITRA_MERGE_DOWNLOAD"
}else{
    Write-Output ""
    Write-Output "Downloading Configuration Script"
}

try{
    $RESULTS = Invoke-WebRequest $CIMITRA_MERGE_DOWNLOAD -OutFile $CIMITRA_MERGE_DOWNLOAD_OUT_FILE -UseBasicParsing 2>&1 | out-null
}catch{}

$theResult = $?

if (!$theResult){
    Write-Output "Error: Could Not Download The File: $CIMITRA_MERGE_DOWNLOAD"
    exit 1
}


try{
    $SUCCESS = Remove-Item -Path $CIMITRA_DOWNLOAD_IMPORT_READ_OUT_FILE -Force -Recurse -ErrorAction SilentlyContinue 2>&1 | out-null
}catch{}

Write-Output ""
Write-Host "Configuring Windows to Allow PowerShell Scripts to Run" -ForegroundColor blue -BackgroundColor white
Write-Output ""
Write-Output ""
Write-Host "If Prompted: Use 'A' For 'Yes to All'" -ForegroundColor blue -BackgroundColor white
Write-Output ""
Unblock-File * 

try{
    powershell.exe -NonInteractive -Command Set-ExecutionPolicy Unrestricted -ErrorAction SilentlyContinue 2>&1 | out-null
}catch{
    Set-ExecutionPolicy Unrestricted -ErrorAction SilentlyContinue 2>&1 | out-null
}

try{
    powershell.exe -NonInteractive -Command Set-ExecutionPolicy Bypass -ErrorAction SilentlyContinue 2>&1 | out-null
}catch{
    Set-ExecutionPolicy Bypass -ErrorAction SilentlyContinue 2>&1 | out-null 2> $null 1> $null
}

try{
    powershell.exe -NonInteractive -Command Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process -ErrorAction SilentlyContinue 2>&1 | out-null
}catch{

    try{
        Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process -ErrorAction SilentlyContinue *> $null | out-null
    }catch{}
}

try{
    powershell.exe -NonInteractive -Command Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser -ErrorAction SilentlyContinue 2>&1 | out-null
}catch{
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser -ErrorAction SilentlyContinue 2>&1 | out-null
    }

try{
    powershell.exe -NonInteractive -Command Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine -ErrorAction SilentlyContinue 2>&1 | out-null
}catch{
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine -ErrorAction SilentlyContinue 2>&1 | out-null
}


$ConfigureScriptExists = Test-Path -Path $CIMITRA_MERGE_DOWNLOAD_OUT_FILE -PathType Leaf

if($ConfigureScriptExists){

    .\configure.ps1
}
