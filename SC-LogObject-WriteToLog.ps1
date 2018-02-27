<#
.NOTES   
    Name: SC-LogObject-WriteToLog.ps1
    Version: 1.2.1
    Author: Kevin McClean
    DateCreated: 2/27/2018
    DateUpdated: 

.SYNOPSIS   
    PowerShell Log Object.

.DESCRIPTION
    Sample Code to use Log Object.

.PARAMETERS [Param]

.Dependencies
    Log Object Ver 1.2.1

.Resources

.Version Notes
#>


#Import Modules


#Global Variables
#----------------
    [String]$gServerName = (Get-WmiObject -Class Win32_ComputerSystem).name
    [String]$gAppName = "TestLogObject"
    [String]$gLoggingDirectory = "C:\Scripts\"

    #Debugging Information
    #---------------------
    [Boolean]$gbolDebug = $true

    #Text Editing
    #------------
    $CRLF = "`r`n"  #Carriage Return and New Line
    $Tab = "'t"     #Tab


#Global Objects
#----------------
    ##Log Object Ver 1.2.1
    #--------------------------------------------
    $gobjLog = New-object -TypeName PSObject
    Add-Member -InputObject $gobjLog -MemberType ScriptMethod -Name "Initialize" -Value {
        param (	
            [Parameter(Position=0, Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [String]$AppName,

            [Parameter(Position=1, Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [String]$FilePath
        )

        #Create NoteProperties
        #----------------------
        Add-Member -InputObject $gobjLog -MemberType NoteProperty -Name "AppName" -value ""
        Add-Member -InputObject $gobjLog -MemberType NoteProperty -Name "FilePath" -value ""
        Add-Member -InputObject $gobjLog -MemberType NoteProperty -Name "Counter" -value 0
        Add-Member -InputObject $gobjLog -MemberType NoteProperty -Name "MaxCount" -value 90000
        Add-Member -InputObject $gobjLog -MemberType NoteProperty -Name "FileDate" -value (Get-Date -Format MMddyyyyHm)
        Add-Member -InputObject $gobjLog -MemberType NoteProperty -Name "FileCount" -value 90000
        Add-Member -InputObject $gobjLog -MemberType NoteProperty -Name "FullFileName" -value ""
        Add-Member -InputObject $gobjLog -MemberType NoteProperty -Name "FileName" -value ""

        #Set Defaults
        $gobjLog.AppName = $AppName
        If ($FilePath.Substring($FilePath.Length -1 ,1) -ne "\" ) {
            $FilePath = $FilePath + "\"
        }
        $gobjLog.FilePath = $FilePath

        $intCounter = 0
        $gobjLog.Counter = $intCounter
        $intMaxCount = 9999
        $gobjLog.MaxCount = $intMaxCount
        $FileDate = Get-Date -Format MMddyyyyHm
        $gobjLog.FileDate = $FileDate
        $intFileCount = 0
        $gobjLog.FileCount = $intFileCount
        
        $gobjLog.FileName = $AppName + "_" + $FileDate + ".txt"
        $gobjLog.FullFileName = $FilePath + $gobjLog.FileName
        

        ##Create Log File
        $sFunctionSection = "Create Log File"

        If ((Test-Path $FilePath) -eq $false){New-Item -ItemType directory $FilePath}
        $StreamWriter = New-Object System.IO.StreamWriter($gobjLog.FullFileName)
        Add-Member -InputObject $gobjLog -MemberType NoteProperty -Name "StreamWriter" -value $StreamWriter


        ##Add Other Methods
        Add-Member -InputObject $gobjLog -MemberType ScriptMethod -Name "Close" -Value {
            #Write Code Here
            #---------------
            #Close Stream Writer Object and Check if File is Empty
            #-----------------------------------------------------
            $gobjLog.StreamWriter.Close()
            $gobjLog.CheckFileEmpty()

            # Remove Members No Longer Needed
            #---------------------------------
            $gobjLog.PSObject.Members.Remove("StreamWriter")
            $gobjLog.PSObject.Members.Remove("AppName")
            $gobjLog.PSObject.Members.Remove("FilePath")
            $gobjLog.PSObject.Members.Remove("Counter")
            $gobjLog.PSObject.Members.Remove("MaxCount")
            $gobjLog.PSObject.Members.Remove("FileDate")
            $gobjLog.PSObject.Members.Remove("FileCount")
            $gobjLog.PSObject.Members.Remove("FullFileName")
            $gobjLog.PSObject.Members.Remove("FileName")

            $gobjLog.PSObject.Members.Remove("CheckFileEmpty")
            $gobjLog.PSObject.Members.Remove("WriteLog")
            $gobjLog.PSObject.Members.Remove("WriteErrorLog")
            $gobjLog.PSObject.Members.Remove("UpdateLogFile")
            $gobjLog.PSObject.Members.Remove("VarDump")
        }

        Add-Member -InputObject $gobjLog -MemberType ScriptMethod -Name "CheckFileEmpty" -Value {
            #Write Code Here
            #---------------
            #Check if StreamWriter is closed
            #--------------------------------
            If ($gobjLog.StreamWriter.BaseStream -eq $null){
                If ((Get-ChildItem $gobjLog.FullFileName).Length -eq 0){
                        remove-Item $gobjLog.FullFileName
                }
            }
        }

        Add-Member -InputObject $gobjLog -MemberType ScriptMethod -Name "WriteLog" -Value {
            param (	
                [Parameter(Position=0, Mandatory=$true)]
                [ValidateNotNullOrEmpty()]
                [String]$sMsg
            )

            #Write Code Here
            #---------------
            #Check Counter
            #-------------
            if($gobjLog.Counter -gt $gobjLog.MaxCount){$gobjLog.UpdateLogFile()}
            
            #Write to Log File
            $gobjLog.StreamWriter.WriteLine($sMsg)

            #Update Internal Counter
            $gobjLog.Counter ++
        }

        Add-Member -InputObject $gobjLog -MemberType ScriptMethod -Name "WriteErrorLog" -Value {
            param (	
                [Parameter(Position=0, Mandatory=$true)]
                [ValidateNotNullOrEmpty()]
                [String]$sFunctionMessage
            )

            #Check Counter
            #-------------
            if($gobjLog.Counter -gt $gobjLog.MaxCount){$gobjLog.UpdateLogFile()}

            #Write to Log File
            If($gbolDebug -eq $true){Write-Warning $sFunctionMessage}
            $gobjLog.WriteLog($sFunctionMessage)
            
            For($x=0; $x -lt $Error.Count; $x++){

                $sErrorMsg = $Error[$x]
                
                If($gbolDebug -eq $true){Write-Warning $sErrorMsg}
                $gobjLog.WriteLog($sErrorMsg)

                $gobjLog.Counter ++
            }

            $sFunctionMessage = "---------------------"
            If($gbolDebug -eq $true){Write-Warning $sFunctionMessage}
            $gobjLog.WriteLog($sFunctionMessage)

            #Reset Error Object
            $Error.Clear()
            #Update Internal Counter
            $gobjLog.Counter ++
            $gobjLog.Counter ++
        }

        Add-Member -InputObject $gobjLog -MemberType ScriptMethod -Name "UpdateLogFile" -Value {
            #Write Code Here
            #---------------
            #Close Current Stream Writer and Check if File Size = 0
            #------------------------------------------------------
            $gobjLog.StreamWriter.Close()
            $gobjLog.CheckFileEmpty()

            #Reset Internal Counter
            $gobjLog.Counter = 0

            #Increment Log Name By 1 and Update Log Name and Create New File
            $gobjLog.FileCount ++
            $gobjLog.FileName = $gobjLog.AppName + "_" + $gobjLog.FileDate + "_" + $gobjLog.FileCount + ".txt"
            $gobjLog.FullFileName =$gobjLog.FilePath + $gobjLog.FileName

            #Create New Log File
            #-------------------
            $gobjLog.StreamWriter = New-Object System.IO.StreamWriter($gobjLog.FullFileName)
        }

        Add-Member -InputObject $gobjLog -MemberType ScriptMethod -Name "VarDump" -Value {
            $gobjLog
        } -PassThru
    }
    #--------------------------------------------


#Functions
#---------
    #Function Message ver 1.0.0
    Function FunctionMessage(){
        param (	
                [Parameter(Position=0, Mandatory=$true)]
                [ValidateNotNullOrEmpty()]
                [String]$FunctionName,
                
                [Parameter(Position=1, Mandatory=$true)]
                [ValidateNotNullOrEmpty()]
                [String]$FunctionSection,
                
                [Parameter(Position=2, Mandatory=$false)]
                [String]$FUnctionMessage
        )
        
        #Function Variables for Debuging
        #-------------------------------
        return $FunctionName + ":" + $FunctionSection + " - " + $FUnctionMessage
    }


#*=============================================================================
#* SCRIPT BODY
#*=============================================================================
$ScriptDateStart = (Get-Date)
Write-Host "--- Start of Script:"($MyStartDate)"---"


    #Body Variables for Debuging
    #-------------------------------
    $sFunctionName = "ScriptBody"
    $sFunctionSection = "CurrentSection"
    $sFunctionMessage = ""


    #Write Information to Log
    #   Create Log File using variables $gAppName and $gLoggingDirectory
    #------------------------
    #Variables
    [String]$sErrorMessage 
    $sFunctionSection = "Write To Log File."

    Try{
        #Write Data to File
        Write-Host "---------------------------"
        $gobjLog.Initialize($gAppName,$gLoggingDirectory)
        $gobjLog.VarDump()
        $FullFileName = $gobjLog.FullFileName
        Write-Warning "File Created: $FullFileName"
        Write-Warning "Writing Data to Log File"
        $gobjLog.WriteLog("Hello this is a sample log message")
        
        $sErrorMessage = FunctionMessage $sFunctionName $sFunctionSection "Sample ERROR Message"
        $gobjLog.WriteErrorLog($sErrorMessage)
        $gobjLog.WriteErrorLog("Sample ERROR Message")
        Write-Host "File Closed, is file deleted?"
        Write-Host "---------------------------"
        $gobjLog.Close()
    }
    Catch{
        Write-Warning $Error[0]
    }


#End
#---------
$Timer = (Get-Date) - $ScriptDateStart
Write-Warning "Function '$sFunctionName'. Total Execution Time: $Timer"
Write-Host "--- End of Script: ---"