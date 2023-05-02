Param(
	[switch]$Console = $false,         #--[ Set to true to enable local console result display. Defaults to false ]--
	[switch]$Debug = $False,           #--[ Generates extra console output for debugging.  Defaults to false ]--
        [switch]$SafeUpdate = $False       #--[ Forces a copy made with a date/Time stamp prior to editing the spreadsheet as a safety backup. ]--  
)
<#PSScriptInfo
.VERSION 1.00
.GUID 
.AUTHOR Kenneth C. Mazie (kcmjr AT kcmjr.com)
.DESCRIPTION 
Tracks switch port status over time using MS Excel.  Full instructions are within the script.
#>
<#==============================================================================
         File Name : SwitchPortActivity.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com)
                   : 
       Description : Tracks switch port status over time using MS Excel.  This is primarily 
                   : intended for cases where your techs don't pull down x-connect jumpers in
                   : IT closets after removing equipment.  The resulting spreadsheet gives you 
                   : an idea of which ports have been abandoned.  The script was designed to 
                   : access and parse Cisco switches.
                   : 
             Notes : Normal operation is with no command line options.  If pre-stored credentials 
                   : are desired use this: https://www.powershellgallery.com/packages/CredentialsWithKey/1.10
                   : Spreadsheet can be generated via a flat list of IP addresses or pulled from a master
                   : copy.  The first column contains the list of IP addresses OR the "Switches" tab from a 
                   : master inventory.  Each IP gets a dedicated worksheet labeled with the IP.  Column A 
                   : is the port ID.  Cell "A1" is a ROUGH total port count.  The top row is the date.
                   : The spreadsheet is color coded for readability:
                   : - If a port at any time registers as connected the "A" colum gets high-lited.
                   : - Connected ports are red.
                   : - Not-connected ports are green (OK to unplug).
                   : - A connection change is flagged in bold violet
                   : - Disabled ports are tagged blue
		           :
      Requirements : Plink.exe must be available in your path or the full path must be included in the 
                   : commandline(s) below.  2 versions are used in case of version issues.  These are located
                   : in the same folder and named according to version (see line 136 below). Excel must be 
                   : available on the local PC.  SSH Keys must already be stored on the local PC through the
                   : use of PuTTY or conenction will fail.  An option exists to add it below.
	               : 
   Option Switches : $Console - If Set to $true will display status during run (Defaults to $True)
                   : $Debug - If set to $true adds extra output on screen (Defaults to $false)
                   : $SafeUpdate - If set to $True backs up spreadsheet prior to updating.  Keeps 10 copies.
                   :
          Warnings : Excel is set to be visible (can be changed) so don't mess with it while the script is 
                   : running or it can crash.  Don't click in spreadsheet while running or the script will crash.
                   :   
             Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF 
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                   : That being said, feel free to ask if you have questions...
                   :
           Credits : Code snippets and/or ideas came from many sources including but 
                   : not limited to the following:
                   : 
    Last Update by : Kenneth C. Mazie                                           
   Version History : v1.00 - 04-16-23 - Original 
    Change History : v2.00 - 00-00-00 - 
				   :                  
==============================================================================#>
Clear-Host
#Requires -version 5

#--[ Variables ]---------------------------------------------------------------
$DateTime = Get-Date -Format MM-dd-yyyy_HHmmss 
$Today = Get-Date -Format MM-dd-yyyy 
$ExcelWorkingCopy = ($MyInvocation.MyCommand.Name.Split("_")[0]).Split(".")[0]+".xlsx"
$ConfigFile = ($MyInvocation.MyCommand.Name.Split("_")[0]).Split(".")[0]+".xml"
$TestFileName = "$PSScriptRoot\test.txt"
$SafeUpdate = $True
#--[ The following can be hardcoded here or loaded from the XML file ]--
#$SourcePath = < See external config file >
#$ExcelSourceFile =  < See external config file >
#$Domain = < See external config file >
#$PasswordFile = < See external config file > 
#$KeyFile = < See external config file > 

#--[ RUNTIME OPTION VARIATIONS ]-----------------------------------------------
$Console = $true
$Script:Debug = $True
#$SafeUpdate = $False
If($Script:Debug){
    $Console = $true
}

#==============================================================================
#==[ Functions ]===============================================================

Function StatusMsg ($Msg, $Color){
    If ($Null -eq $Color){
        $Color = "Magenta"
    }
    If ($Script:Debug){Write-Host "-- Script Status: $Msg" -ForegroundColor $Color}
    $Msg = ""
}

Function LoadConfig {
    #--[ Read and load configuration file ]-------------------------------------
    if (!(Test-Path "$PSScriptRoot\$ConfigFile")){                       #--[ Error out if configuration file doesn't exist ]--
        StatusMsg "MISSING CONFIG FILE.  Script aborted." " Red"
        break;break;break
    }else{
        [xml]$Configuration = Get-Content "$PSScriptRoot\$ConfigFile"  #--[ Read & Load XML ]--    
        $Script:SourcePath = $Configuration.Settings.General.SourcePath
        $Script:ExcelSourceFile = $Configuration.Settings.General.ExcelSourceFile
        $Script:PasswordFile = $Configuration.Settings.Credentials.PasswordFile
        $Script:KeyFile = $Configuration.Settings.Credentials.KeyFile
    }
}

Function CallPlink ($IP,$command){
    $ErrorActionPreference = "silentlycontinue"
    $Switch = $False
    $UN = $Env:USERNAME
    $DN = $Env:USERDOMAIN
    $UID = $DN+"\"+$UN
    If (Test-Path -Path $PasswordFile){
        $Base64String = (Get-Content $KeyFile)
        $ByteArray = [System.Convert]::FromBase64String($Base64String)
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UID, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $ByteArray)
    }Else{
        $Credential = $Script:ManualCreds
    }
    #------------------[ Decrypted Result ]-----------------------------------------
    $Password = $Credential.GetNetworkCredential().Password
    $Domain = $Credential.GetNetworkCredential().Domain
    $Username = $Domain+"\"+$Credential.GetNetworkCredential().UserName

    If (Test-Connection -ComputerName $IP -count 1 -BufferSize 16 -Quiet) {
        #--[ Detect and store SSH key in local registry if needed ]--
        # StatusMsg "Automatically storing SSH key if needed." "Magenta"
        # Write-Output "Y" | 
        # plink-v52.exe -ssh -pw $password $username@$IP #"exit" #*>&1
        # plink-v73.exe -ssh -pw $password $username@$IP #-batch #"exit" #*>&1
        # Start-Sleep -Milliseconds 500
        #------------------------------------------------------------
        StatusMsg "Plink IP: $IP" "Magenta"
        #$test = @(plink-v73.exe -ssh -no-antispoof -pw $Password $username@$IP $command ) #*>&1)
        $test = @(plink-v73.exe -ssh -no-antispoof -batch -pw $Password $username@$IP $command *>&1)
        If ($test -like "*abandoned*"){
            StatusMsg "Switching Plink version" "Magenta"
            $Switch = $true
        }Else{
            StatusMsg 'Plink version 73 test passed' 'Magenta'
        }
        If ($Switch){
            $Msg = 'Executing Plink v52 (Command = '+$Command+')'
            StatusMsg $Msg 'blue'
            $Result = @(plink-v52.exe -ssh -no-antispoof -batch -pw $Password $username@$IP $command *>&1) 
        }Else{
            $ErrorActionPreference = "continue"
            $Msg = 'Executing Plink v73 (Command = '+$Command+')'
            StatusMsg $Msg 'magenta'
            $Result = @(plink-v73.exe -ssh -no-antispoof -batch -pw $Password $username@$IP $command *>&1)
        }
        ForEach ($Line in $Result){
            If ($Line -like "*denied*"){
                $Result = "ACCESS-DENIED"
                Break
            } 
        }
        StatusMsg "Data collected..." "Magenta"
        Return $Result
    }Else{
        StatusMsg "Pre-Plink PING check FAILED" "Red"
    }
} 

Function CellColor ($WorkSheet,$Item,$Row,$Col){
    $Status = ""
    if ($Col -ne 1){
    Switch -Wildcard ([String]$Item){
        "*notconnect*" {
            $Status = "notconnect"
            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 10 
            $Worksheet.Cells($Row, $Col).Font.Bold = $False
            Break
        }
        "*connected*" {
            $Status = "connected"
            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 3 
            $Worksheet.Cells($Row, $Col).Font.Bold = $False
            #--[ If at any time a row cell goes red the port ID in column 1 is changed to a pale ]--
            #--[ yellow background to indicate that the port has had activity at some point during ]--
            #--[ the scanning period.  Just a quick glance indicator. ]--
            $Worksheet.Cells($Row, 1).Interior.ColorIndex = 36 
            Break
        } 
        "*err-disabled*" {
            $Status = "err-disabled"                                    
            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 6
            $Worksheet.Cells($Row, $Col).Interior.ColorIndex = 3
            $Worksheet.Cells($Row, $Col).Font.Bold = $true
            Break
        } 
        "*disabled*" {
            $Status = "disabled"
            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 5
            $Worksheet.Cells($Row, $Col).Font.Bold = $False
            Break
        } 
        Default {$Status = "Unknown"}
    }
    CellBorder $WorkSheet $Row $Col
}
    Return $Status
}

Function ProcessTarget ($WorkBook,$WorkSheet,$SheetCounter){
    #==[ Begin Processing of IP from Sheet ]=============================================
    $TotalSheets = $WorkBook.Sheets.Count - 1
    $command = 'sh version'
    [string]$Port = "23"
    $ErrorActionPreference = "stop"
    $Result = ""
    $WorkSheet.activate()
    $IP = $WorkSheet.Name #.Split(";")[0]).Trim()
    $Today = Get-Date -Format MM-dd-yyyy 
    #--[ Determine the next empty column ]--
    $Col = 1
    Do {
        $Col++  #--[ 1st empty column ]--
    }
    Until ( $Null -eq $workSheet.Cells.Item(1, $Col).Value2) 

    If ($Console){
        Write-host "`n--[ Current Device: "  -ForegroundColor yellow -NoNewline
        Write-Host $IP  -ForegroundColor cyan -NoNewline
        Write-Host " ($SheetCounter of $TotalSheets) ]---------------------------------------------------" -ForegroundColor yellow 
    }

    #--[ Test and connect to target IP ]----------------------------------------------------------
    if ([string]$WorkSheet.cells.Item(1,($Col-1)).text -as [DateTime]) {  #--[ Cell IS a date ]--
        If ((Get-Date ($WorkSheet.cells.Item(1,($Col-1)).text) -Format MM-dd-yyyy) -eq $today){  #--[ Cell is today's date ]--
            StatusMsg "Today has already been processed.  Moving to next IP..." "red"   
            return  
        }
    }
    $WorkSheet.cells.Item(1,$Col) = $Today  #--[ Either way datestamp next empty cell ]--
    $Worksheet.Cells(1,$Col).Font.Bold = $true
    CellBorder $WorkSheet "1" $Col
    If (Test-Connection -ComputerName $IP -count 1 -BufferSize 16 -Quiet){
        $Worksheet.Cells(1, $Col).Font.ColorIndex = 0
        If (Test-Path -Path 'C:\Program Files\PuTTY\'){ 
            StatusMsg "Calling PLINK" "Magenta"
            $command = 'sh int status'
            $Result = CallPlink $IP $command
        }Else{
            StatusMsg "Cannot find PLINK.EXE.   Aborting..." "Red"
            break;break
        }

        #==[ Parse Main Result Variable ]=======================================================
        StatusMsg "Parsing collected data..." "Magenta"
        If ($Result -eq "ACCESS-DENIED"){
            StatusMsg "ACCESS DENIED" "Red"
            $Worksheet.Cells.Item(2, $Col) = "No Access"
            $Worksheet.Cells(2, $Col).Font.ColorIndex = 3   
            CellBorder $WorkSheet "2" $Col              
        }Else{
            $Row = 2  
            StatusMsg  "Writing to column $Col" "Magenta"
            ForEach ($Item in $Result){  #--[ Parse results ]--
                If ($Console){
                    Write-host "." -NoNewline
                }
                $ErrorActionPreference = "silentlycontinue"
                Start-Sleep -Milliseconds 50  #--[ Slow things down a bit to avoid data skewing ]--    
                If (($Item -is [string]) -and ($Item.length -gt 0)){
                    If (([string]$Item.Substring(0,2) -Like "Gi*" ) -Or ([string]$Item.Substring(0,2) -Like "Te*" )){
                        $Port = [string]$Item.Split(" ")[0]
                        $Status = CellColor $WorkSheet $Item $Row $Col  #--[ Assign cell color by status ]--
                        #--[ Write Port ID ]--
                        $WorkSheet.cells.Item($Row,1) = $Port
                        $Worksheet.Cells($Row, 1).Font.Bold = $true
                        CellBorder $WorkSheet $Row 1
                        #--[ Write Port ID ]--
                        $WorkSheet.cells.Item($Row,$Col) = $Status  
                        #--[ if a change is detected color the cell bold magenta ]--
                        If (($WorkSheet.cells.Item($Row,($Col-1)).Text -notlike "*/*") -And (($WorkSheet.cells.Item($Row,$Col).Text.Trim()) -ne ($WorkSheet.cells.Item($Row,($Col-1)).Text.Trim()))){
                            $Worksheet.Cells($Row, $Col).Font.Bold = $true
                            $Worksheet.Cells($Row, $Col).Font.ColorIndex = 7  #--[ Violet to denote a change ]--
                        }
                        #--[ The below line changes the previous days cell color removing the Magenta. ]--
                        #--[ This may make seeing changes harder. Change at your discression. ]--
                        #$Status = CellColor $WorkSheet $Item $Row ($Col-1)  
                        $Row++
                    }
                }
            }
            If ($Console){Write-Host ""}
        }  #--[ End of result ]--
        If ($Null -eq $WorkSheet.Cells.Item(1,1).Value2){
            $Row = 2
            Do {
                $Row++  #--[ 1st empty row ]--
            }
            Until ( $Null -eq $workSheet.Cells.Item($row,1).Value2) 
            $Row=$Row-2
            $Value = "$Row Ports"
            $workSheet.Cells.Item(1,1) = $Value
            $Worksheet.Cells(1,1).Font.Bold = $true
            CellBorder $WorkSheet 1 1
        }
    }Else{
        StatusMsg "--- No Connection ---" "Red"
        $Worksheet.Cells.Item(2,$Col) = "No Ping"
        $Worksheet.Cells(2,$Col).Font.ColorIndex = 3   # --[ End of connection ]--
        CellBorder $WorkSheet 2 $Col
    }
    $Item = ""
    $Resize = $WorkSheet.UsedRange
    [Void]$Resize.EntireColumn.AutoFit()
}

Function Open-Excel ($Excel,$ExcelWorkingCopy,$SheetName,$Console) {
    If (Test-Path -Path "$PSScriptRoot\$ExcelWorkingCopy" -PathType Leaf){
        If ($SafeUpdate){
            If ($Console){Write-Host "-- Script Status: Safe-Update Enabled. Creating a backup copy of the spreadsheet..." -ForegroundColor Green}
            $Backup = $DateTime+"_"+$ExcelWorkingCopy
            Copy-Item -Path "$PSScriptRoot\$ExcelWorkingCopy"  -Destination "$PSScriptRoot\$Backup"
        }
        $Script:SpreadSheet = "Existing"
        $WorkBook = $Excel.Workbooks.Open("$PSScriptRoot\$ExcelWorkingCopy")
        $WorkSheet = $Workbook.WorkSheets.Item($SheetName)
        $WorkSheet.activate()
    }Else{
        $Script:SpreadSheet = "New"
        $Workbook = $Excel.Workbooks.Add()
        $Worksheet = $Workbook.Sheets.Item(1)
        $Worksheet.Activate()
        $WorkSheet.Name = $SheetName
        [int]$Col = 1
        CellBorder $WorkSheet 1 $Col
        $WorkSheet.cells.Item(1,$Col++) = "Port #"   
        $Range = $WorkSheet.Range(("A1"),("AZ1")) 
        $Range.font.bold = $True
        $Range.HorizontalAlignment = -4108  #Alignment Middle
        $Range.Font.ColorIndex = 1
        $Range.font.bold = $True
        $Resize = $WorkSheet.UsedRange
        [Void]$Resize.EntireColumn.AutoFit()
    }
    Return $WorkBook
}

Function CellBorder ($WorkSheet, $Row, $Col){
    $WorkSheet.Cells($Row,$Col).Borders.ColorIndex = 1
    $WorkSheet.Cells($Row,$Col).Borders.weight = 2  
}
#==[ End of Functions ]===================================================

#==[ Begin ]==============================================================
LoadConfig

If (!(Test-Path -Path $PasswordFile)){
    $Script:ManualCreds = Get-Credential -Message 'Enter an appropriate Domain\User and Password to continue.'
}

StatusMsg "Processing Cisco Switches" "Yellow"

#--[ Close copies of Excel that PowerShell has open ]--
$ProcID = Get-CimInstance Win32_Process | where {$_.name -like "*excel*"}
ForEach ($ID in $ProcID){  #--[ Kill any open instances to avoid issues ]--
    Foreach ($Proc in (get-process -id $id.ProcessId)){
        if (($ID.CommandLine -like "*/automation -Embedding") -Or ($proc.MainWindowTitle -like "$ExcelWorkingCopy*")){
            Stop-Process -ID $ID.ProcessId -Force
            write-host "-- Killing any existing open PowerShell instance of Excel..." -ForegroundColor Red
            Start-Sleep -Milliseconds 100
        }
    }
}

#--[ Create new Excel COM object ]--
$Excel = New-Object -ComObject Excel.Application -ErrorAction Stop

#--[ Make a backup of the working copy, keep only the last 10 ]--
If (($SafeUpdate)-And (Test-Path -Path "$PSScriptRoot\$ExcelWorkingCopy")){
    StatusMsg "Safe-Update Enabled. Creating a backup copy of the working spreadsheet..." "Green"
    $Backup = $DateTime+"_"+$ExcelWorkingCopy+".bak"
    Copy-Item -Path "$PSScriptRoot\$ExcelWorkingCopy"  -Destination "$PSScriptRoot\$Backup"
    #--[ Only keep 10 of the last backups ]-- 
    Get-ChildItem -Path $PSScriptRoot | Where-Object {(-not $_.PsIsContainer) -and ($_.Name -like "*.bak")} | Sort-Object -Descending -Property LastTimeWrite | Select-Object -Skip 10 | Remove-Item
}

#--[ If this file exists the IP list will be pulled from it ]--
If (Test-Path -Path $TestFileName){   
    $ListFileName = $TestFileName   #--[ Select an alternate short IP text file to use ]--
}Else{ 
    $ListFileName = "$PSScriptRoot\IPlist.txt"   #--[ Select the normal IP text file to use ]--
}

#--[ Identify IP address list source and process. ]--
If (Test-Path -Path $ListFileName){  #--[ If text file exists pull from there. ]--
    $IP = Get-Content $ListFileName          
    StatusMsg "IP text list was found, loading IP list from it... " "green" 
    If (Test-Path -Path "$PSScriptRoot\$ExcelWorkingCopy"){
        StatusMsg ">>>     WARNING: Working copy already exists.     <<<" "Yellow"
        StatusMsg ">>>  New copy will be created and NOT over-write. <<<" "Yellow"
        StatusMsg ">>> Remember to delete IP file prior to next run. <<<" "Yellow"
        Start-Sleep -Seconds 5
    }
    StatusMsg "Creating new Spreadsheet..." "green"
    $Workbook = $Excel.Workbooks.Add()
    $Excel.Visible = $True
    $Worksheet = $Workbook.Sheets.Item(1)   
    $WorkSheet.Name = "Switches"
    $Worksheet.Activate()
    $Worksheet.Cells(1,1).Font.Bold = $true
    $Worksheet.Cells(1,1).Font.ColorIndex = 1
    $WorkSheet.Cells.Item(1,1) = "IP Address"
    $WorkSheet.Cells(1,1).Font.Bold = $true
    $Row = 2
    ForEach ($Address in $IP){
        $WorkSheet.Cells.Item($Row,1) = $Address
        $Row++    
        StatusMsg  "Adding new workbook for $Address" "Magenta"
        $LastSheet = $WorkBook.Worksheets | Select-Object -Last 1
        $NewSheet = $WorkBook.worksheets.add($LastSheet)
        $NewSheet.Name = $Address
        $LastSheet.Move($NewSheet)
    }
    $Resize = $WorkSheet.UsedRange
    [Void]$Resize.EntireColumn.AutoFit()
}Else{  #--[ If no text file exists try to pull IPs from Excel ]--
    StatusMsg "IP text list not found, attempting to process spreadsheet... " "cyan"
    #--[ Verify if a working copy exists, otherwise copy from source ]--
    If (Test-Path -Path "$PSScriptRoot\$ExcelWorkingCopy" -PathType Leaf){
        Write-host "-- Referencing " -NoNewline -ForegroundColor Green
        Write-Host $ExcelWorkingCopy -ForegroundColor Yellow -NoNewline
        Write-Host " spreadsheet for IP list. --- " -ForegroundColor Green
        $WorkBook = $Excel.Workbooks.Open("$PSScriptRoot\$ExcelWorkingCopy")
        $Excel.Visible = $True
        $WorkSheet = $Workbook.WorkSheets.Item("Switches")
        $WorkSheet.activate()
    }Else{
        If (!(Test-Path -Path "$PSScriptRoot\$ExcelWorkingCopy" -PathType Leaf)){ 
            StatusMsg "Excel working copy is missing, copying from source..." "Magenta"
            If (Test-Path -Path "$SourcePath\$ExcelSourceFile" -PathType Leaf){
                StatusMsg "Master source file located and copied to script folder..." "Green"
                Copy-Item -Path "$SourcePath\$ExcelSourceFile"  -Destination "$PSScriptRoot\$ExcelWorkingCopy" -force
                $Excel.Visible = $true
                $Excel.displayalerts = $False
                $WorkBook = $Excel.Workbooks.Open("$PSScriptRoot\$ExcelWorkingCopy")
                StatusMsg "Removing un-needed worksheets..." "Green"
                ForEach ($Sheet in $WorkBook.Worksheets){
                    If ($Sheet.Name -ne "Switches"){
                        $Sheet.Delete()
                    }
                }
                $sheet = $workbook.Sheets.Item(1) #--[ Activate the worksheet ]--
                [void]$sheet.Cells.Item(1, 1).EntireRow.Delete() #--[ Delete the first row ]--
                [void]$sheet.Cells.Item(1, 1).EntireRow.Delete() #--[ Delete the first row again ]--
                $WorkBook.Save()  #--[ Save the new spreadsheet prior to processing ]--
            }Else{
                StatusMsg "Master source file not found, Nothing to process.  Exiting... " "Red"
                Break;Break
            }
        }
    }

    $Row = 2   
    $IPList = @() 
    $PreviousIP = ""
    $CurrentIP = ""
    If ($Console){
        StatusMsg "Reading Spreadsheet... " "Magenta"
    }
    Do {
        $CurrentIP = $WorkSheet.Cells.Item($Row,1).Text  
        $PortCount = $WorkSheet.Cells.Item($Row,29).Text   
        If ($CurrentIP -ne $PreviousIP){  #--[ Make sure IPs are added only once per switch stack ]--
            $IPList += ,@($CurrentIP) #+";"+$Facility+";"+$Address+";"+$IDF+";"+$Description+";"+$Row)
            $PreviousIP = $CurrentIP
        }
        $Row++

        If (!($WorkBook.worksheets | Where-Object {$_.name -eq $CurrentIP})){
            StatusMsg  "Adding new workbook for $CurrentIP" "Magenta"
            $LastSheet = $WorkBook.Worksheets | Select-Object -Last 1
            $NewSheet = $WorkBook.worksheets.add($LastSheet)
            $NewSheet.Name = $CurrentIP
            $NewSheet.Cells.Item(1,1) = $PortCount+" Ports"
            $NewSheet.Cells(1,1).Font.Bold = $true
            CellBorder $NewSheet 1 1
            $LastSheet.Move($NewSheet)
        }
    } Until (
        $WorkSheet.Cells.Item($row,1).Text -eq ""   #--[ Condition that stops the loop if it returns true ]--
    )
    $Excel.DisplayAlerts = $false    
    #$Excel.quit()  #--[ Close it.  Only do so if you are pulling the IP list and doing nothing else, otherwise bad things happen  ]--
}

#--[ Process each worksheet except the list source on SWITCHES sheet ]--
$SheetCounter = 1
ForEach ($Book in $WorkBook.worksheets | Where-Object {$_.Name -ne "Switches"}){
    $WorkSheet = $Workbook.WorkSheets.Item($Book.Name)
    ProcessTarget $WorkBook $WorkSheet $SheetCounter
    $SheetCounter++
}

#--[ Cleanup ]--
Write-host ""
Try{ 
    If ((Test-Path -Path $ListFileName) -And (Test-Path -Path "$PSScriptRoot\$ExcelWorkingCopy")){
        StatusMsg 'Saving as "NewSpreadsheet.xlsx" ...' "Green"
        $Workbook.SaveAs("$PSScriptRoot\NewSpreadsheet.xlsx")
    }ElseIf(!(Test-Path -Path "$PSScriptRoot\$ExcelWorkingCopy")){
        StatusMsg "Saving as a new working spreadsheet... " "Green"
        $Workbook.SaveAs("$PSScriptRoot\$ExcelWorkingCopy")
    }Else{
        StatusMsg "Saving working spreadsheet... " "Green"
        $WorkBook.Close($true) #--[ Close workbook and save changes ]--
    }
    $Excel.quit() #--[ Quit Excel ]--
    [Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) #--[ Release the COM object ]--
}Catch{
    StatusMsg "Save Failed..." "Red"
}

Write-Host `n"--- COMPLETED ---" -ForegroundColor red

<#--[ XML File Example -- File should be named same as the script ]--
<!-- Settings & configuration file -->
<Settings>
    <General>
        <SourcePath>'C:\Users\Documents'</SourcePath>
        <ExcelSourceFile>Master-Inventory.xlsx</ExcelSourceFile>
        <Domain>company.org</Domain>
    </General>
    <Credentials>
	<PasswordFile>c:\pw.txt</PasswordFile>
        <KeyFile>c:\key.txt</KeyFile>
    </Credentials>
</Settings>    


#>

   
