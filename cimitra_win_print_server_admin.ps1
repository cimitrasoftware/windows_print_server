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
[switch] $TestPrintJob,
[switch] $ReportJobs,
[string] $ClearJobById,
[string] $PrinterName,
[string] $ComputerName,
[switch] $QuietMode,
[switch] $ServerReport,
[switch] $ServerQueueReport,
[switch] $ServerQueueReportDisabled,
[string] $NonInputA,
[string] $NonInputB,
[string] $SharedPrinterList
)

$Global:VerboseMode = $true
if($QuietMode){
    $Global:VerboseMode = $false
}


$TEMP_FILE_ONE = New-TemporaryFile
$Global:ServerQueuedJobCounter = 0

function Call-Exit($ErrorMessageIn,$ErrorExitCode){
    Write-Output "ERROR: $ErrorMessageIn"
    exit $ErrorExitCode
}

#Confirm That Windows Print Management is installed
$PrinterManagementTest = Get-Printer -ErrorAction Stop

if(!($PrinterManagementTest)){

    Call-Exit "Printer Services Are Not Available" "1"
}

if(!($ServerReport -or $ServerQueueReport)){

    if($PrinterName.Length -lt 2){
        Call-Exit "Specify a Printer Name With The -PrinterName Parameter" "1"
    }

}



if($ClearJobById.Length -gt 0){

    $Global:ThePrintJobId = $ClearJobById.Trim()
 
 }



if($ComputerName){
$Global:TheComputerName = $ComputerName
}else{
$Global:TheComputerName = $env:COMPUTERNAME
}





if(!($ServerReport -or $ServerQueueReport)){
# This is a Try/Catch method
# In this script, the Try/Catch method is determining if the Windows Printer Name specified exists



$Global:PrinterNameExists = $true
try{
    $ThePrinterInput = Get-Printer -ComputerName "$TheComputerName" -Name "$PrinterName" 
}catch{
    $Global:PrinterNameExists = $false
}
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

# Use this function to add Print Job ID passed into the script into an Array
Function Correlate-Print-Job-IDs{
# Turn list of GUIDS passed into script into an array
    param(
        [Parameter(Mandatory=$true)]
        [string]$PrinterIdList,
        [array]$add
    )

        $PrinterIdsToProcess = $PrinterIdList.split(' ')
        try{
        $PrinterIdsToProcess += $add.split(' ')
        }catch{}


    return $PrinterIdsToProcess
}

$Global:RemoveMultipleJobs = $false

if($ClearJobById.Length -gt 1){


$ArrayOfPrintJobIDs = Correlate-Print-Job-IDs "$ClearJobById" 

$Global:PrintJobIDs = $ArrayOfPrintJobIDs | select -Unique

if($PrintJobIDs.Length -gt 1){
    $Global:RemoveMultipleJobs = $true
}

}

function Remove-Printer-Job(){


$JobFound = $true

if($ThePrintJobId.Length -lt 1){
    return
}

try{
    $ThePrintJob = Get-PrintJob -PrinterName "$PrinterName" -ID "$ThePrintJobId" -ErrorAction Stop
}catch{
    $JobFound = $false
}



if(!($JobFound)){

    Write-Output "Cannot Find Print Job ID: $ThePrintJobId"
    return

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
    Write-Output ""
}else{
    Write-Output "Error: Did Not Delete Job ID: $ThePrintJobId" 
    Write-Output ""
}


}

function Remove-Printer-Jobs(){

$PrintJobIDs.ForEach({
    $Global:ThePrintJobId = $_
    Remove-Printer-Job
})


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
    Write-Output "Print Server Process ID: [ $PrinterPID ]"
}

}

function Printer-Server-Report(){

$PrintServerRunning = Get-Process spoolsv

if($PrintServerRunning){
    Write-Output ""
    Write-Output "Printer Server Status: [ RUNNING ]"
    Write-Output ""
}else{
    Write-Output ""
    Write-Output "==============================="
    Write-Output "Printer Status: [ NOT RUNNING ]"
    Write-Output "==============================="
    return
}


Printer-Server-PID-Report

}

function Printer-Report(){
$ThePrinter = Get-Printer -ComputerName "$TheComputerName" -Name "$PrinterName"

Write-Output "Printer Queue Status Report"
Write-Output "==========================="

if($ThePrinter.PrinterStatus -eq 0){
    Write-Output "READY"
}else{
    Write-Output "OFFLINE"
}

Write-Output "==========================="

if($VerboseMode){
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

$NumberOfPrintJobs = $PrintJobs.Length

if($NumberOfPrintJobs -eq 0){
    Write-Output "There are Currently [ 0 ] Print Jobs"
    return 0
}

if($NumberOfPrintJobs -eq 1){
    Write-Output "There is Currently [ 1 ] Print Job"
}

if($NumberOfPrintJobs -gt 1){
    Write-Output "There are Currently [ $NumberOfPrintJobs ] Print Jobs"
}

Write-Output "Print Jobs IDs Only Report"
Write-Output "=========================="

$PrintJobs.ForEach({

$ThePrintJob = $_
$ThePrintJobId = $ThePrintJob.Id

Write-Host -NoNewline " $ThePrintJobId "

})
Write-Output ""
Write-Output "=========================="



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

Write-Output "Print Jobs IDs Only Report"
Write-Output "=========================="

$PrintJobs.ForEach({

$ThePrintJob = $_
$ThePrintJobId = $ThePrintJob.Id

Write-Host -NoNewline " $ThePrintJobId "

})

Write-Output ""
Write-Output "=========================="

if($VerboseMode){
Write-Output ""
Write-Output ""
Write-Output "Print Jobs Detail Report"
Write-Output "========================"

$PrintJobs.ForEach({

$ThePrintJob = $_
$ThePrintJobId = $ThePrintJob.Id
$ThePrintJobCreator = $ThePrintJob.UserName
$ThePrintJobDocumentName = $ThePrintJob.DocumentName

    Write-Output "Print Job ID: $ThePrintJobId | Details"
    $ThePrintJob | Format-List

Write-Output "Print Job ID: $ThePrintJobId"
Write-Output "Print Job Creator: $ThePrintJobCreator"
Write-Output "Print Job Document: $ThePrintJobDocumentName"

})

}



Write-Output ""

if($VerboseMode){

    Write-Output "Print Jobs IDs Only Report"
    Write-Output "=========================="

    $PrintJobs.ForEach({

    $ThePrintJob = $_
    $ThePrintJobId = $ThePrintJob.Id

    Write-Host -NoNewline " $ThePrintJobId "

    })
    Write-Output ""
    Write-Output "=========================="
}





}


function Print-Test-Page()
{
    $result = Get-CimInstance Win32_Printer -Filter "name LIKE '$PrinterName'" | Invoke-CimMethod -MethodName printtestpage 

if ($result.ReturnValue -eq 0)
{
    Write-Output ""
    write-output "Test page printed for $PrinterName"
}
else
{
    Write-Output ""
    write-output "Unable to print test page on $PrinterName"
    write-output "Error code $($result.ReturnValue)"
}
}

function Print-Test-Pages(){

Printer-Queue-Report

$Global:TheNumberOfPrintJobs = 1

if($ClearJobById.Length -gt 0){


$NumberOfPrintJobs = $ClearJobById.Trim()
$NumberOfPrintJobs = ($NumberOfPrintJobs.Split(" "))[0]
$NumberOfPrintJobs = $NumberOfPrintJobs.Trim()




if($NumberOfPrintJobs -match "^\d+$"){
    $Global:TheNumberOfPrintJobs = $NumberOfPrintJobs
}else{
    $Global:TheNumberOfPrintJobs = 1
}

if($TheNumberOfPrintJobs -gt 5){
    $Global:TheNumberOfPrintJobs = 5
        
}

$Counter = 0

Write-Output "Print Jobs Requested: $NumberOfPrintJobs"

while($TheNumberOfPrintJobs -gt $Counter){

     $Counter++

        Write-Output "Creating Print Job #${Counter} of $TheNumberOfPrintJobs"

        if($Counter -gt 6){
            exit 1
        }

  

        Print-Test-Page

     }
}else{

        Print-Test-Page
}

Printer-Queue-Report

}






## These If(...) tests call a Function if a particular Parameter was sent to the script

# Run if the -Report Parameter was called

function Report-Printer-Overview($ThePrinterIn,$ThePrinterLabelIn){

$Global:PrinterNameExists = $true
try{
    $ThePrinterInput = Get-Printer -ComputerName "$TheComputerName" -Name "$ThePrinterIn" -ErrorAction Stop
}catch{
    $Global:PrinterNameExists = $false
}

    if(!($PrinterNameExists)){
        return
     }

$ThePrinter = Get-Printer -ComputerName "$TheComputerName" -Name "$ThePrinterIn"

Write-Output "" >> $TEMP_FILE_ONE
Write-Output "Printer Name: $ThePrinterLabelIn" >> $TEMP_FILE_ONE

if($ThePrinter.PrinterStatus -eq 0){
    Write-Output "Queue Status: Ready" >> $TEMP_FILE_ONE
}else{
    Write-Output "Queue Status: Offline" >> $TEMP_FILE_ONE
}



$QueuedPrintJobs = (Get-Printer -ComputerName "$TheComputerName" -Name "$ThePrinterIn").JobCount
Write-Output "Jobs In Queue: $QueuedPrintJobs" >> $TEMP_FILE_ONE

$Global:ServerQueuedJobCounter = $ServerQueuedJobCounter + $QueuedPrintJobs


}

function Report-All-Shared-Printers(){


if($SharedPrinterList.Length -gt 5){
    $Global:SharedPrinterListFile = "$SharedPrinterList" 
}else{
    $Global:SharedPrinterListFile = "${PSScriptRoot}\import\AllSharedPrinters.txt"
}

$SharedPrinterListFileExists = Test-Path -Path ${SharedPrinterListFile} -PathType Leaf -ErrorAction Stop

if(!($SharedPrinterListFileExists)){

    Write-Output "Error Shared Printer List File: ${SharedPrinterListFile} Does Not Exist"
    exit 1
}


$StreamReader = New-Object System.IO.StreamReader($SharedPrinterListFile)
$line_number = 1
while (($CurrentLine = $StreamReader.ReadLine()) -ne $null)
{

    if(($CurrentLine.ToCharArray()) -contains ",")
    {
    
        $ThePrinterLabel = ($CurrentLine.Split(","))[0]
        $ThePrinterName = ($CurrentLine.Split(","))[1]
    }else{

        $ThePrinterLabel = $CurrentLine
        $ThePrinterName = $CurrentLine

    }
   
    Report-Printer-Overview "$ThePrinterName" "$ThePrinterLabel"
    $line_number++
}


# $Global:ServerQueuedJobCounter = 0

    Write-Output ""
    Write-Output "Total Queued Jobs Across All Printers: ${ServerQueuedJobCounter}"


    Get-Content $TEMP_FILE_ONE

    Remove-Item -Path $TEMP_FILE_ONE -Force 2>&1 | out-null

}



if($Report){

    if($PrinterName.Length -lt 3){
        Write-Output "[ Action - Print Server Report ]"
        Printer-Server-Report
     }else{
        Write-Output "[ Action - Printer Report ]"
        Printer-Queue-Report
        Printer-Report
     }

     exit 0
}


# Run if the -Status Parameter was called
if($Status){
    Write-Output "[ Action - Print Server Status ]"
    Printer-Server-Report
    exit 0
}

# Run if the -Start Parameter was called
if($Start){
    Write-Output "[ Action - Start Print Server ]"
    Start-Print-Server
    Printer-Server-Report
    exit 0
}

# Run if the -Stop Parameter was called
if($Stop){
    Write-Output "[ Action - Stop Print Server ]"
    Stop-Print-Server
    exit 0
  
}

# Run if the -Restart Parameter was called
if($Restart){
    Write-Output "[ Action - Restart Print Server ]"
    Stop-Print-Server
    Start-Print-Server
    Printer-Server-Report
    exit 0
}


# Run if the -ClearJobs Parameter was called
if($ClearJobs){
    Write-Output "[ Action - Clear All Print Jobs ]"
    Printer-Queue-Report
    Clear-Printer-Queue
    Printer-Queue-Report

    exit 0

}

# Run if the -ReportJobs Parameter was called
if($ReportJobs){
    Write-Output "[ Action - Report Print Jobs ]"
    Printer-Queue-Report
    exit 0
}

if(!($TestPrintJob)){
if($ClearJobById.Length -gt 0){


if($RemoveMultipleJobs){
    Write-Output "[ Action - Remove Print Jobs ]"
    Printer-Queue-Report
    Remove-Printer-Jobs
    

}else{
    Write-Output "[ Action - Remove Print Job ]"
    Printer-Queue-Report
    Remove-Printer-Job

}


    Printer-Queue-Report

    
exit 0    
}

}
    


if($TestPrintJob){

    Write-Output "[ Action - Create Test Print Job ]"

    if($PrinterName.Length -lt 2){
        Write-Output "Please Choose a Printer"
    }else{

        Print-Test-Pages
        Write-Output ""
    }

    Printer-Report

    exit 0
}

# Run if the -ServerReport Parameter was called
if($ServerReport){
    
    if(!($ServerQueueReportDisabled)){
        Write-Output "[ Action - Print Server Overview Report ]"
        Report-All-Shared-Printers
        Printer-Server-Report
    }else{
        Write-Output "[ Action - Print Server Report ]"
        Printer-Server-Report
    }
    exit 0
}

if($ServerQueueReport){
    Write-Output "[ Action - Print Server Overview Report ]"
    Report-All-Shared-Printers
}


