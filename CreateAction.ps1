﻿# Cimitra Active Directory Integration Module Import Script
# Author: Tay Kratzer tay@cimitra.com
# 9/21/2021
# Created on a Windows 2019 Server

Param(
[switch] $DisablePrinterTitleCase,
[string] $InstallationDirectory
)
 

$global:LegacyPowershell = $false

$versionMinimum = [Version]'6.0'

if ($versionMinimum -gt $PSVersionTable.PSVersion){ 
    $global:LegacyPowershell = $true
 }

if($InstallationDirectory.Length -gt 4){
    $global:INSTALLATION_DIRECTORY = $InstallationDirectory
}else{
    $global:INSTALLATION_DIRECTORY = "C:\cimitra\scripts\cimitra_win_print_server_admin"
}



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

$global:IMPORT_HOME_DIRECTORY = "$INSTALLATION_DIRECTORY\import"

try{
    New-Item -ItemType Directory -Force -Path $IMPORT_HOME_DIRECTORY 2>&1 | out-null
}catch{}

$ThisScript = $MyInvocation.MyCommand.Name


$Global:PRINTERS_CSV_FILE = "$IMPORT_HOME_DIRECTORY\Printers.csv"
$Global:PRINTERS_ACTION_JSON_FILE = "$IMPORT_HOME_DIRECTORY\Printers_Action.json"
$Global:DiscoverPrintersRan = $false


$EXTRACTED_DIRECTORY = "$INSTALLATION_DIRECTORY\cimitra_win_print_server_admin-main"

function Call-Exit($ExitMessage,$ExitCode){

Write-Output "Error: $ExitMessage"
exit $ExitCode

}



function Discover-Printers-CSV(){

# Put all Shared Printers in a CSV File

if($DiscoverPrintersRan){
return
}

$Global:DiscoverPrintersRan = $true

$Printers = [System.Collections.ArrayList]::new()

$ListOfPrinters = Get-Printer -Name  * | select-object "Name" -ErrorAction Stop

$NumberOfPrinters = $ListOfPrinters.Length

if( $ListOfPrinters.Length -lt 1){
    Call-Exit "Cannot Discover any Printers" "1"
}

# Make an array to hold all of the Printers
$SharedPrinters = [System.Collections.ArrayList]::new()


$ListOfPrinters.ForEach({ $CurrentPrinter = $_.Name

# Iterate through all printers
$PrinterObject = Get-Printer -Name "$CurrentPrinter"

        # If the printer is shared, add it to the Array $SharedPrinters
        if($PrinterObject.Shared){
 
            [void]$SharedPrinters.Add("$CurrentPrinter")
       }

})


$NumberOfPrinters = $SharedPrinters.Length

if( $NumberOfPrinters -eq 0 ){
    Call-Exit "Cannot Discover any Shared Printers" "1"
}



$TEMP_FILE_ONE = New-TemporaryFile
# Create a CSV file with the name of each printer
$SharedPrinters.ForEach({ 

    $ThePrinter = $_
    

    $ThePrinterHasQuote = $ThePrinter -contains "'"

    if($ThePrinterHasQuote){
        continue
      }

    $ThePrinterHasQuotes = $ThePrinter -contains '"'

    if($ThePrinterHasQuotes){
        continue
      }

    $ThePrinterHasCommas = $ThePrinter -contains ','


    if($ThePrinterHasCommas){
        continue
      }

    if(!($DisablePrinterTitleCase)){
        $ThePrinterTitle = $ThePrinter.ToUpper()
    }else{
        $ThePrinterTitle = $ThePrinter
    }

    Add-Content -Path $TEMP_FILE_ONE -Value "$ThePrinterTitle,$ThePrinter"
})


Move-Item -Force -Path $TEMP_FILE_ONE -Destination ${PRINTERS_CSV_FILE}

$Global:CSVImportFile = ${PRINTERS_CSV_FILE}
$Global:DiscoverPrintersRan = $true
}


function Make-Printers-Action-JSON-File(){

if(!($DiscoverPrintersRan)){
return
}

    $FirstPrinter = $true

    $CSVFileContent = Get-content -Path "$PRINTERS_CSV_FILE"
    $Counter = 0

       if($LegacyPowershell){
            $PARAMETER_JSON_FILE='{"paramsRunButtons":"topandbottom","cronEnabled":false,"type":1,"status":"active","itemtype":"App","injectParams":[{"boolStrings":{"true":"","false":""},"paramtype":6,"required":true,"multipleParamDelim":",","maxVisibleLines":0,"private":false,"allowUnmasking":false,"encapsulateChar":"\"","allowed":[{"default":true,"name":"PRINT SERVER STATUS","value":"-ServerReport"},{"default":false,"name":"REPORT ALL JOBS FOR A PRINTER","value":"-Report"},{"default":false,"name":"REMOVE A PRINTER JOB BY ID","value":""},{"default":false,"name":"REMOVE ALL JOBS FOR A PRINTER","value":"-ClearJobs"},{"default":false,"name":"START PRINT SERVER","value":"-Start"},{"default":false,"name":"STOP PRINT SERVER","value":"-Stop"},{"default":false,"name":"RESTART PRINT SERVER","value":"-Restart"}],"param":"","value":false,"label":"ACTIONS","regex":"/^[a-zA-Z0-9\\-\\+\\=\\_ ]+$/","dateType":"datetime-local","dateMin":"","dateMax":"","dateFormatNamed":"epochSec","dateFormat":"","placeholder":""},{"paramtype":4,"required":false,"multipleParamDelim":",","maxVisibleLines":0,"private":false,"allowUnmasking":false,"encapsulateChar":"\"","allowed":[_REPLACE_WITH_PRINTERS_LIST_],"param":"-PrinterName ","value":"","label":"PRINTERS","regex":"/^[a-zA-Z0-9\\-\\+\\=\\_ ]+$/"},{"boolStrings":{"true":"","false":""},"paramtype":2,"required":false,"multipleParamDelim":",","maxVisibleLines":0,"private":false,"allowUnmasking":false,"encapsulateChar":"","allowed":[],"param":"-ClearJobById","value":"","label":"JOB ID","regex":"/^[0-9\\ ]+$/","dateType":"datetime-local","dateMin":"","dateMax":"","dateFormatNamed":"epochSec","dateFormat":"","placeholder":"4"}],"comments":[],"platform":"win32","interpreter":"C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe","command":"C:\\cimitra\\scripts\\cimitra_win_print_server_admin\\cimitra_win_print_server_admin.ps1","cronZone":"Etc/UTC","name":"CONTROL PRINT JOBS","description":"<span style=\"color:#2B60DE; font-size: 19px; font-family: Arial, Helvetica, sans-serif;font-weight:900\">\nControl Print Jobs on a Print Server</span>","notes":"[PowerShell 5]","__v":2,"shares":[]}'
            
        }else{
            $PARAMETER_JSON_FILE='{"paramsRunButtons":"topandbottom","cronEnabled":false,"type":1,"status":"active","itemtype":"App","injectParams":[{"boolStrings":{"true":"","false":""},"paramtype":6,"required":true,"multipleParamDelim":",","maxVisibleLines":0,"private":false,"allowUnmasking":false,"encapsulateChar":"\"","allowed":[{"default":true,"name":"PRINT SERVER STATUS","value":"-ServerReport"},{"default":false,"name":"REPORT ALL JOBS FOR A PRINTER","value":"-Report"},{"default":false,"name":"REMOVE A PRINTER JOB BY ID","value":""},{"default":false,"name":"REMOVE ALL JOBS FOR A PRINTER","value":"-ClearJobs"},{"default":false,"name":"START PRINT SERVER","value":"-Start"},{"default":false,"name":"STOP PRINT SERVER","value":"-Stop"},{"default":false,"name":"RESTART PRINT SERVER","value":"-Restart"}],"param":"","value":false,"label":"ACTIONS","regex":"/^[a-zA-Z0-9\\-\\+\\=\\_ ]+$/","dateType":"datetime-local","dateMin":"","dateMax":"","dateFormatNamed":"epochSec","dateFormat":"","placeholder":""},{"paramtype":4,"required":false,"multipleParamDelim":",","maxVisibleLines":0,"private":false,"allowUnmasking":false,"encapsulateChar":"\"","allowed":[_REPLACE_WITH_PRINTERS_LIST_],"param":"-PrinterName ","value":"","label":"PRINTERS","regex":"/^[a-zA-Z0-9\\-\\+\\=\\_ ]+$/"},{"boolStrings":{"true":"","false":""},"paramtype":2,"required":false,"multipleParamDelim":",","maxVisibleLines":0,"private":false,"allowUnmasking":false,"encapsulateChar":"\"","allowed":[],"param":"-ClearJobById","value":"","label":"JOB ID","regex":"/^[0-9\\ ]+$/","dateType":"datetime-local","dateMin":"","dateMax":"","dateFormatNamed":"epochSec","dateFormat":"","placeholder":"4"}],"comments":[],"platform":"win32","interpreter":"C:\\Program Files\\PowerShell\\7\\pwsh.exe","command":"C:\\cimitra\\scripts\\cimitra_win_print_server_admin\\cimitra_win_print_server_admin.ps1","params":"","cronZone":"Etc/UTC","name":"CONTROL PRINT JOBS","description":"<span style=\"color:#2B60DE; font-size: 19px; font-family: Arial, Helvetica, sans-serif;font-weight:900\">\nControl Print Jobs on a Print Server</span>","notes":"[PowerShell 7]","__v":10,"shares":[]}'
        }

       
    if($LegacyPowershell){
        if($FirstPrinter){
            $PARAMETER_PRINTERS_LINE='{"default":true,"name":"_TITLE_REPLACE_","value":"''_VALUE_REPLACE_''"}'
            $FirstPrinter = $false
        }else{
            $PARAMETER_PRINTERS_LINE='{"default":false,"name":"_TITLE_REPLACE_","value":"''_VALUE_REPLACE_''"}'
        }
    }else{

        if($FirstPrinter){
            $PARAMETER_PRINTERS_LINE='{"default":true,"name":"_TITLE_REPLACE_","value":"\"_VALUE_REPLACE_\""}'
            $FirstPrinter = $false
        }else{
            $PARAMETER_PRINTERS_LINE='{"default":false,"name":"_TITLE_REPLACE_","value":"\"_VALUE_REPLACE_\""}'
        }
    }

    $THE_PARAMETER_PRINTERS_LINE = $PARAMETER_PRINTERS_LINE

    $PARAMETER_COPIED = $true

        while($Counter -lt $CSVFileContent.Count){
                $Counter++
                if(!($PARAMETER_COPIED)){
                    $THE_PARAMETER_PRINTERS_LINE = "${THE_PARAMETER_PRINTERS_LINE},${PARAMETER_PRINTERS_LINE}"
                }
                $TheLine = Get-Content $PRINTERS_CSV_FILE | select -First $Counter | select -Last 1
  
                $ThePrinterTitle = $TheLine.Split(',')[0]
                $ThePrinterValue = $TheLine.Split(',',2)[1]

                $THE_PARAMETER_PRINTERS_LINE = $PARAMETER_PRINTERS_LINE.Replace("_TITLE_REPLACE_", "$ThePrinterTitle")
                $THE_PARAMETER_PRINTERS_LINE = $THE_PARAMETER_PRINTERS_LINE.Replace("_VALUE_REPLACE_", "$ThePrinterValue")
                $PARAMETER_COPIED = $false
            
        }

$PARAMETER_JSON_FILE = $PARAMETER_JSON_FILE.Replace("_REPLACE_WITH_PRINTERS_LIST_", "$THE_PARAMETER_PRINTERS_LINE")

try{
Set-Content -Path "$PRINTERS_ACTION_JSON_FILE" -Value $PARAMETER_JSON_FILE
}catch{
Call-Exit "Cannot Write to Temporary File: $PRINTERS_ACTION_JSON_FILE" "1"
}


Write-Output ""
Write-Output "Made Cimitra Action [ Control Print Server ] Import File"
Write-Output ""
Write-Output "$PRINTERS_ACTION_JSON_FILE"
Write-Output ""

}


Discover-Printers-CSV
Make-Printers-Action-JSON-File


     










 