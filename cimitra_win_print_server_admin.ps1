# Manage a Windows Print Server
# Author: Tay Kratzer tay@cimitra.com
# Modify Date: 9/21/21
# Created and tested on a Windows 2019 Server
# -------------------------------------------------

<#
.DESCRIPTION
Manage a Windows Print Server
#>

## This is the Parameters section ##
# Parameters have to be at the very top of a PowerShell script
# The Parameters section is an Array of elements
# Parameters are called or passed to the script with this synax: <script name> -Parameter
# Example Use: ./cimitra_win_print_server_admin.ps1 -Report
Param(

[switch] $Report,
[switch] $Status,
[switch] $Start,
[switch] $Stop,
[switch] $Restart,
[switch] $ClearJobs,
[switch] $ReportJobs,
[string] $ClearJobById,
[string] $PrinterName,
[string] $ComputerName,
[switch] $QuietMode,
[switch] $ServerReport,
[string] $NonInputA,
[string] $NonInputB
)

$Global:VerboseMode = $true
if($QuietMode){
    $Global:VerboseMode = $false
}




function Call-Exit($ErrorMessageIn,$ErrorExitCode){
    Write-Output "ERROR: $ErrorMessageIn"
    exit $ErrorExitCode
}

#Confirm That Windows Print Management is installed
$PrinterManagementTest = Get-Printer -ErrorAction Stop

if(!($PrinterManagementTest)){

    Call-Exit "Printer Services Are Not Available" "1"
}

if(!($ServerReport)){

    if($PrinterName.Length -lt 2){
        Call-Exit "Specify a Printer Name With The -PrinterName Parameter" "1"
    }

}

if($ClearJobById.Length -gt 0){

    $Global:ThePrintJobId = $ClearJobById -replace '\s+', ' '
 
    if(!($ThePrintJobId -match "^\d+$")){
        Call-Exit "Use a Number Value Only" "1"
    }   

}

if($ComputerName){
$Global:TheComputerName = $ComputerName
}else{
$Global:TheComputerName = $env:COMPUTERNAME
}


# This is a Try/Catch method
# In this script, the Try/Catch method is determining if the Windows Printer Name specified exists
$PrinterNameExists = $true
try{
    $ThePrinterInput = Get-Printer -ComputerName "$TheComputerName" -Name "$PrinterName" 
}catch{
    $PrinterNameExists = $false
}

if(!($ServerReport)){

    if(!($PrinterNameExists)){
        Call-Exit "A Printer Does Not Exist With The Name: $PrinterName" "1"
    }else{
        Write-Output "Acting On Printer: $PrinterName"
        Write-Output "" 
    }

}


# This a Function
# Functions are like little programs within the script
function Stop-Print-Server(){

Write-Output "Stopped Print Server"

    try{
        net stop spooler
    }catch{
        Write-Output "Unable to Stop Print Server"
    }
}


function Start-Print-Server(){


Stop-Print-Server


Write-Output "Restart Print Server"


try{
    $SUCCESS = net start spooler 2> $null
}catch{
    Write-Output "Unable to Start Print Server"
    # Return from this function, because there is no need to go further if the Windows Service will not restart
    return
}


}

function Printer-Queue-Report(){

# Get a listing of all files in the Printers Queue
$DIR_PRINT_QUEUE_FILES = (Get-Printer -ComputerName "$TheComputerName" -Name "$PrinterName").JobCount

if($DIR_PRINT_QUEUE_FILES -eq 1){
    Write-Output ""
    Write-Output "1 Print Job Exists For Printer: $PrinterName"
}else{
    Write-Output ""
    Write-Output "$DIR_PRINT_QUEUE_FILES Print Jobs Exist For Printer: $PrinterName"
}

Write-Output ""
}

function Remove-Printer-Job(){


$JobFound = $true

try{
    $ThePrintJob = Get-PrintJob -PrinterName "$PrinterName" -ID "$ThePrintJobId" -ErrorAction Stop
}catch{
    $JobFound = $false
}


if(!($JobFound)){

    Call-Exit "Cannot Find Print Job ID: $ThePrintJobId" "1"

}

if($VerboseMode){
    Write-Output "Print Job ID: $ThePrintJobId | Details"
    $ThePrintJob | Format-List
}


$ThePrintJobCreator = $ThePrintJob.UserName
$ThePrintJobDocumentName = $ThePrintJob.DocumentName


Write-Output ""
Write-Output "Print Job ID: $ThePrintJobId"
Write-Output "Print Job Creator: $ThePrintJobCreator"
Write-Output "Print Job Document: $ThePrintJobDocumentName"
Write-Output ""


$JobDeleted = $true
try{
    $SUCCESS = Remove-PrintJob -ComputerName "$TheComputerName" -PrinterName "$PrinterName" -ID "$ThePrintJobId" -ErrorAction Stop
}catch{
    $JobDeleted = $false
}

if($JobDeleted){
    Write-Output "Successfully Deleted Job ID: $ThePrintJobId"
}else{

    Call-Exit "Did Not Delete Job ID: $ThePrintJobId" "1"
}



}

function Clear-Printer-Queue(){

Write-Output "Clear Print Queue"

$DIR_PRINT_QUEUE_FILES = (Get-Printer -ComputerName "$TheComputerName" -Name "$PrinterName").JobCount

if($DIR_PRINT_QUEUE_FILES -eq 0){
    Write-Output ""
    Write-Output "No Print Jobs Exists For Printer: $PrinterName"
    return
}


$PrintJobs = Get-PrintJob -ComputerName "$TheComputerName"  -PrinterName "$PrinterName"


$PrintJobs.ForEach({

$ThePrintJob = $_
$ThePrintJobId = $ThePrintJob.Id

if($VerboseMode){
    Write-Output "Print Job ID: $ThePrintJobId | Details"
    $ThePrintJob | Format-List
}

$ThePrintJobCreator = $ThePrintJob.UserName
$ThePrintJobDocumentName = $ThePrintJob.DocumentName


Write-Output "Removing Print Job ID: $ThePrintJobId"
Write-Output "Print Job Creator: $ThePrintJobCreator"
Write-Output "Print Job Document: $ThePrintJobDocumentName"
Write-Output ""

Remove-PrintJob -ComputerName "$TheComputerName" -PrinterName "$PrinterName" -ID  "$ThePrintJobId"


})

}


function Printer-Server-PID-Report(){

$PrintServerRunning = Get-Process spoolsv

if($PrintServerRunning){
    $PrinterPID = $PrintServerRunning.Id
    Write-Output "Printer Process ID: [ $PrinterPID ]"
}

}

function Printer-Server-Report(){

$PrintServerRunning = Get-Process spoolsv

if($PrintServerRunning){

    Write-Output "Printer Server Status: [ RUNNING ]"
    Write-Output ""
}else{
    Write-Output "==============================="
    Write-Output "Printer Status: [ NOT RUNNING ]"
    Write-Output "==============================="
    return
}


Printer-Server-PID-Report

}

function Printer-Report(){
$ThePrinter = Get-Printer -ComputerName "$TheComputerName" -Name "$PrinterName"
if($VerbosMode){
    $ThePrinter | Format-List
}


$DIR_PRINT_QUEUE_FILES = (Get-Printer -ComputerName "$TheComputerName" -Name "$PrinterName").JobCount

if($DIR_PRINT_QUEUE_FILES -eq 0){
    Write-Output "No Print Jobs Exist For Printer: $PrinterName"
    return
}


$PrintJobs = Get-PrintJob -ComputerName "$TheComputerName"  -PrinterName "$PrinterName"


Write-Output "Print Jobs Summary Report"
Write-Output "========================="

$PrintJobs.ForEach({

$ThePrintJob = $_
$ThePrintJobId = $ThePrintJob.Id
$ThePrintJobCreator = $ThePrintJob.UserName
$ThePrintJobDocumentName = $ThePrintJob.DocumentName


Write-Output ""
Write-Output "Print Job ID: $ThePrintJobId"
Write-Output "Print Job Creator: $ThePrintJobCreator"
Write-Output "Print Job Document: $ThePrintJobDocumentName"
Write-Output ""

})


Write-Output "Print Jobs Detail Report"
Write-Output "========================"

$PrintJobs.ForEach({

$ThePrintJob = $_
$ThePrintJobId = $ThePrintJob.Id
$ThePrintJobCreator = $ThePrintJob.UserName
$ThePrintJobDocumentName = $ThePrintJob.DocumentName

if($VerboseMode){
    Write-Output "Print Job ID: $ThePrintJobId | Details"
    $ThePrintJob | Format-List
}

Write-Output "Print Job ID: $ThePrintJobId"
Write-Output "Print Job Creator: $ThePrintJobCreator"
Write-Output "Print Job Document: $ThePrintJobDocumentName"

})

Write-Output ""





}


## These If(...) tests call a Function if a particular Parameter was sent to the script

# Run if the -Report Parameter was called




if($Report){

    if($PrinterName.Length -lt 3){
        Write-Output "Action - Print Server Report"
        Printer-Server-Report
     }else{
        Write-Output "Action - Printer Report"
        Printer-Queue-Report
        Printer-Report
     }

     exit 0
}

# Run if the -ServerReport Parameter was called
if($ServerReport){
    Write-Output "Action - Printer Server Report"
    Printer-Server-Report
    exit 0
}

# Run if the -Status Parameter was called
if($Status){
    Write-Output "Action - Print Server Status"
    Printer-Server-Report
    exit 0
}

# Run if the -Start Parameter was called
if($Start){
    Write-Output "Action - Start Print Server"
    Start-Print-Server
    Printer-Server-Report
    exit 0
}

# Run if the -Stop Parameter was called
if($Stop){
    Write-Output "Action - Stop Print Server"
    Stop-Print-Server
    exit 0
  
}

# Run if the -Restart Parameter was called
if($Restart){
    Write-Output "Action - Restart Print Server"
    Stop-Print-Server
    Start-Print-Server
    Printer-Server-Report
    exit 0
}


# Run if the -ClearJobs Parameter was called
if($ClearJobs){
    Write-Output "Action - Clear All Print Jobs"
    Printer-Queue-Report
    Clear-Printer-Queue
    Printer-Queue-Report
    exit 0

}

# Run if the -ReportJobs Parameter was called
if($ReportJobs){
    Write-Output "Action - Report Print Jobs"
    Printer-Queue-Report
    exit 0
}

if($ClearJobById.Length -gt 0){
    Remove-Printer-Job
}


