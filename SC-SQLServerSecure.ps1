<#
.NOTES   
    Name: _ScriptTemplate_Ver101.ps1
    Version: 1.0.2
    Author: Kevin McClean
    DateCreated: 05/01/2017
    DateUpdated: 05/8/2017

.SYNOPSIS   
    Template used to create PowerShell Scripts.

.DESCRIPTION
    Requires the Log Object as the Script Automatically creates a Log File.

.PARAMETERS [Param]


.EXAMPLE

   
.Dependencies
    Log Object Ver 1.2.1

.Resources
    Powershell
        Comparison Operators  http://ss64.com/ps/syntax-compare.html
        Data Types http://ss64.com/ps/syntax-datatypes.html
        Param Options  https://social.technet.microsoft.com/wiki/contents/articles/15994.powershell-advanced-function-parameter-attributes.aspx

.Version Notes
    1.0.2 (5/8/2017)
    Update Notes Section

    1.0.1 (5/4/2017)
    Added a section to Pass Parameters to the Template.
#>

<#
.NOTES   
    Name: _Ver100.ps1
    Version: 1.0.0
    Author: Kevin McClean
    DateCreated:
    DateUpdated:

.SYNOPSIS   
    

.DESCRIPTION
    

.PARAMETERS


.EXAMPLE

   
.Dependencies
    Log Object Ver 1.2.1

.Resources
    Powershell
        Comparison Operators  http://ss64.com/ps/syntax-compare.html
        Data Types http://ss64.com/ps/syntax-datatypes.html
        Param Options  https://social.technet.microsoft.com/wiki/contents/articles/15994.powershell-advanced-function-parameter-attributes.aspx

.Version Notes

#>

#Parameters Passed to Script
    param (	
            [Parameter(Position=0, Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [String]$gsParameter1
    )


#Import Modules


#Global Variables
#----------------
    [String]$gServerName = (Get-WmiObject -Class Win32_ComputerSystem).name
    [String]$gAppName = "MyAppName"
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

    Function MyFunction(){
<#
        param (	
                [Parameter(Position=0, Mandatory=$true)]
                [ValidateNotNullOrEmpty()]
                [Boolean]$bolDebug
        )
#>
        #Function Variables for Debuging
        #-------------------------------
        $DateStart = (Get-Date)
        $sFunctionName = "MyFunction"
        $sFunctionSection = "Start"
        $sFunctionMessage = "MessageHere"
        If($gbolDebug -eq $true){
            $sFunctionMessage = FunctionMessage $sFunctionName $sFunctionSection $sFunctionMessage
            Write-Host $sFunctionMessage
            $gobjLog.WriteLog($sFunctionMessage)
        }


        #Variables
        #---------


        #Write Code Here
        #---------------
        $sFunctionSection = "Write Code Here"
        If($gbolDebug -eq $true){
            $sFunctionMessage = FunctionMessage $sFunctionName $sFunctionSection
            Write-Host $sFunctionMessage
            $gobjLog.WriteLog($sFunctionMessage)
        }
        
        try{
        
        }
        Catch{
            #Error Found
            #-----------
            $sFunctionMessage = "Additional Details"
            $sFunctionMessage = FunctionMessage $sFunctionName $sFunctionSection $sFunctionMessage
            $gobjLog.WriteErrorLog($sFunctionMessage)
        }


        #Close Objects
        #-------------
        $sFunctionSection = "Close Objects"
        If($gbolDebug -eq $true){
            $sFunctionMessage = FunctionMessage $sFunctionName $sFunctionSection "Closing Objects."
            Write-Host $sFunctionMessage
            $gobjLog.WriteLog($sFunctionMessage)
        }


        #End
        #---------
        $Timer = (Get-Date) - $DateStart
        If($gbolDebug -eq $true){
            $sFunctionMessage = "Function '$sFunctionName'. Total Execution Time: $Timer"
            Write-Warning  $sFunctionMessage 
            $gobjLog.WriteLog($sFunctionMessage)
        }
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


    #Variables
    #---------


    #Initialize Objects
    #*-----------------
    $sFunctionSection = "Initialize Object(s)"
    #Create Log/Error Log Files
    $gobjLog.Initialize($gAppName,$gLoggingDirectory)
    If($gbolDebug -eq $true){
        $sFunctionMessage = "Creating Log File: " + $gobjLog.FullFileName
        $sFunctionMessage = FunctionMessage $sFunctionName $sFunctionSection $sFunctionMessage
        Write-Host $sFunctionMessage
        $gobjLog.WriteLog($sFunctionMessage)
    }
    #Create Other Objects Below
    #--------------------------



    ### Place Code Here
    #------------------
    $sFunctionSection = "Place Code Here"
    If($gbolDebug -eq $true){
        $sFunctionMessage = FunctionMessage $sFunctionName $sFunctionSection $sFunctionMessage
        Write-Host $sFunctionMessage
        $gobjLog.WriteLog($sFunctionMessage)
    }
    
    MyFunction


    #Close Objects
    #*============
    $sFunctionSection = "Close Objects"
    #Close Other Objects Below
    #-------------------------
    #
    #Logging Object
    #---------------
    If($gbolDebug -eq $true){
        $Timer = (Get-Date) - $ScriptDateStart
        $sFunctionMessage = FunctionMessage $sFunctionName $sFunctionSection ("Closing Log File: " + $gobjLog.FullFileName)
        $sFunctionMessage = $sFunctionMessage + $CRLF + "Function '$sFunctionName'. Total Execution Time: $Timer" + $CRLF + "--- End of Script: ---"
        Write-Host $sFunctionMessage
        $gobjLog.WriteLog($sFunctionMessage)
    }
    $gobjLog.Close()


#End
#---------
$Timer = (Get-Date) - $ScriptDateStart
Write-Warning "Function '$sFunctionName'. Total Execution Time: $Timer"
Write-Host "--- End of Script: ---"
