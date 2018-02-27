<#
.NOTES   
    Name: CustomObjects.ps1
    Version: 1.0.3
    Author: Kevin McClean
    DateCreated: 05/01/2017
    DateUpdated: 05/8/2017

.SYNOPSIS   
    Repository of Powershell Objects.

.DESCRIPTION
    Code samples .

.PARAMETERS [Param]


.EXAMPLE
    See SC-[name].ps1 for all Sample Code illustrating the use of each object
   
.Dependencies
    Log Object Ver 1.2.1

.Resources
    Powershell
        Comparison Operators  http://ss64.com/ps/syntax-compare.html
        Data Types http://ss64.com/ps/syntax-datatypes.html
        Param Options  https://social.technet.microsoft.com/wiki/contents/articles/15994.powershell-advanced-function-parameter-attributes.aspx

.Version Notes
    1.0.3 (02/27/2017)
    Adding Code to GitHub for Tracking!

    1.0.2 (5/8/2017)
    Update Notes Section

    1.0.1 (5/4/2017)
    Added a section to Pass Parameters to the Template.
#>


$ScriptBlock_LogObject = {
#Log Object Ver 1.2.1
<#  #Version Notes
    #-------------
    1.2.1
    Changed the ScriptMethod 'WriteErrorLog'
    Now It will Automatically Write a Generic Error to the Log File and the Screen.
    - Uses the following Parameters:
        NO NEED TO PASS THE FOLLOWING:
        -- $gbolDebug (Defined In Every Script!)
        -- $Error Object which is Global to the PowerShell Script
                
        Pass the Follwing String
        -- $sFunctionMessage (Message Defined using the Function 'FunctionMessage()')
                
    <# Code Snipet
    Add-Member -InputObject $gobjLog -MemberType ScriptMethod -Name "WriteErrorLog" -Value {
        param (	
            [Parameter(Position=0, Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [String]$sMsg,

            [Parameter(Position=1, Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            $ErrorMsg
        )

        #Write Code Here
        #---------------
        #Check Counter
        #-------------
        if($gobjLog.Counter -gt $gobjLog.MaxCount){$gobjLog.UpdateLogFile()}

        #Write to Log File
        $gobjLog.StreamWriter.WriteLine($sMsg)
        $gobjLog.StreamWriter.WriteLine($ErrorMsg)

        #Update Internal Counter
        $gobjLog.Counter ++
        $gobjLog.Counter ++
    }
    #>
#>

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
}


$ScriptBlock_SecurePasswordObject = {
#Secure Password Object Ver 1.0.0
<#
    ****************
    * SPECIAL NOTE 
    *
    *  This method has some limitation that it will only work for the same user on the same machine.
    *  No other user or profile on same machine or any other machine can read or decrypt this file. 
    *  Decryption can only be done via same user on same machine.
    *
    ****************

    #Resources
    https://gallery.technet.microsoft.com/scriptcenter/Secure-Password-using-c158a888

#>
    #Secure Password Object Ver 1.0.0
    #--------------------------------------------
    $objSecurePwd = New-object -TypeName PSObject
    Add-Member -InputObject $objSecurePwd -MemberType ScriptMethod -Name "Initialize" -Value {
        param (	
            [Parameter(Position=0, Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [String]$SecureFileLocation,

            [Parameter(Position=1, Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [Boolean]$bolCreateFile
        )
    
        #Add Other Methods
        #-----------------
        Add-Member -InputObject $objSecurePwd -MemberType NoteProperty -Name "SecureFileLocation" -value $SecureFileLocation
        Add-Member -InputObject $objSecurePwd -MemberType ScriptMethod -Name "GetPassword" -Value {
            get-content $objSecurePwd.SecureFileLocation | convertto-securestring
        }
        Add-Member -InputObject $objSecurePwd -MemberType ScriptMethod -Name "Close" -Value {
            ## Remove Members No Longer Needed
            $objSecurePwd.PSObject.Members.Remove("SecureFileLocation")

        }


        #Determine if we are Creating an Encrypted File?
        #-----------------------------------------------
        If($bolCreateFile -eq $true){
            #Create Encrypted File
            read-host -assecurestring | convertfrom-securestring | out-file $SecureFileLocation
        }
    }
    #--------------------------------------------
}


$ScriptBlock_SQLServerSecureObject ={
#SQL Server Secure Object Ver 1.2.3
<#
.NOTES   
    Name: _Ver123.ps1
    Version: 1.2.3
    Author: Kevin McClean
    DateCreated: 05/01/2017
    DateUpdated: 05/11/2017

.SYNOPSIS   
    

.DESCRIPTION
    

.PARAMETERS


.EXAMPLE

   
.Dependencies
    Secure Password Object Ver 1.0.0
    Log Object Ver 1.2.1

.Resources
    https://ask.sqlservercentral.com/questions/121106/using-powershell-credentials-to-connect-to-sql-ser.html
    https://msdn.microsoft.com/en-us/library/system.data.sqlclient.sqldatareader.read(v=vs.110).aspx
    http://www.systemcentercentral.com/powershell-how-to-connect-to-a-remote-sql-database-and-retrieve-a-data-set/
    (Time Out Property) http://stackoverflow.com/questions/19362793/powershell-sql-server-update-query
    https://msdn.microsoft.com/en-us/library/system.data.sqlclient.sqlcommand.executenonquery%28v=vs.110%29.aspx?f=255&MSPPError=-2147217396

.Version Notes
    1.2.3 (5/11/2017)
    - Now tells if the Object Initialized Properly or Not.

    1.2.2 (5/3/2017)
    Added CommandTimeOut Property and set Default to 3 minutes
    Updated ScriptMethod "ExecuteCommandOnly" to return # of Rows affected if it's an Update, Delete or Create Query!
        - No Longer using the ExecuteReader commmand.
    
    1.2.1 (5/1/2017)
    Changed the ScriptMethod 'ExecuteCommand'
    - The NoteProperty "SQLReader" did not function properly after Assingment.  Run the code below with the Write-Host statements
    and you'll see that after assingment, the object does not have any records but before it does.
    - This also means that the 'Close' Method also needed to be updated.
    - Added an Execute SQL Method ONLY which just runs an SQL Query but Doesn't Return anything.
    - Added an Execute SQL Method that adds a Method that Returns a SQL Data Set.
#>
    #SQL Server Secure Object Ver 1.2.3
    #--------------------------------------------
    $gobjSQL = New-object -TypeName PSObject
    Add-Member -InputObject $gobjSQL -MemberType ScriptMethod -Name "Initialize" -Value{
        param (	
            [Parameter(Position=0, Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [String]$ServerName,

            [Parameter(Position=1, Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [String]$DatabaseName,

            [Parameter(Position=2, Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [String]$UserName,

            [Parameter(Position=3, Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [String]$SecureFullFileName
        )

        #Variables
        #---------
        [Boolean]$bolInitialized = $false

        #Set Defaults
        #-------------
        Add-Member -InputObject $gobjSQL -MemberType NoteProperty -Name ServerName -value $ServerName
        Add-Member -InputObject $gobjSQL -MemberType NoteProperty -Name DatabaseName -value $DatabaseName
        Add-Member -InputObject $gobjSQL -MemberType NoteProperty -Name UserName -value $UserName
        Add-Member -InputObject $gobjSQL -MemberType NoteProperty -Name SecureFileLocation -value $SecureFullFileName

        #Test SQL Connection
        #-------------------
        #Secure Password Object Ver 1.0.0
        #################################
        $objSQLPwd = New-object -TypeName PSObject
        Add-Member -InputObject $objSQLPwd -MemberType ScriptMethod -Name "Initialize" -Value {
            param (	
                [Parameter(Position=0, Mandatory=$true)]
                [ValidateNotNullOrEmpty()]
                [String]$SecureFileLocation,

                [Parameter(Position=1, Mandatory=$true)]
                [ValidateNotNullOrEmpty()]
                [Boolean]$bolCreateFile
            )
    
            #Add Other Methods
            #-----------------
            Add-Member -InputObject $objSQLPwd -MemberType NoteProperty -Name "SecureFileLocation" -value $SecureFileLocation
            Add-Member -InputObject $objSQLPwd -MemberType ScriptMethod -Name "GetPassword" -Value {
                get-content $objSQLPwd.SecureFileLocation | convertto-securestring
            }
            Add-Member -InputObject $objSQLPwd -MemberType ScriptMethod -Name "Close" -Value {
                ## Remove Members No Longer Needed
                $objSQLPwd.PSObject.Members.Remove("SecureFileLocation")
                $objSQLPwd.PSObject.Members.Remove("GetPassword")
            }


            #Determine if we are Creating an Encrypted File?
            #-----------------------------------------------
            If($bolCreateFile -eq $true){
                #Create Encrypted File
                read-host -assecurestring | convertfrom-securestring | out-file $SecureFileLocation
            }
        }
        #################################
        Try{
            #Add Connection Object
            $Connection = New-Object System.Data.SqlClient.SqlConnection
            Add-Member -InputObject $gobjSQL -MemberType NoteProperty -Name Connection -value $Connection

            #Create SQL Credentials
            $objSQLPwd.Initialize($gobjSQL.SecureFileLocation, $false)
            $MyPassword = $objSQLPwd.GetPassword()
            $MyPassword.MakeReadOnly()
            $creds = New-Object System.Data.SqlClient.SqlCredential($gobjSQL.UserName,$MyPassword)
            
            $gobjSQL.Connection.Credential = $creds
            $gobjSQL.Connection.ConnectionString = "Server=" + $gobjSQL.ServerName + ";Database= " + $gobjSQL.DatabaseName + ";MultipleActiveResultSets=True"

            #Open SQL Connection
            $gobjSQL.Connection.open()
            $bolInitialized = $true
        }
        Catch{
            $bolInitialized = $False
            #Remove the Following Objects
            #*** It remembers the Information Passed during the 1st Init, if called again the value do not update.
            $gobjSQL.PSObject.Members.Remove("Connection")
            $gobjSQL.PSObject.Members.Remove("ServerName")
            $gobjSQL.PSObject.Members.Remove("DatabaseName")
            $gobjSQL.PSObject.Members.Remove("UserName")
            $gobjSQL.PSObject.Members.Remove("SecureFileLocation")
            
            $gobjLog.WriteErrorLog("Could Not Connect To SQL Server on $ServerName, Test Initialization Settings!")
        }

        #Set Security Related Objects to Null
        $objSQLPwd = $null
        $MyPassword = $null


        ##Add Other Methods
        If($bolInitialized -eq $true){
            Add-Member -InputObject $gobjSQL -MemberType ScriptMethod -Name "Close" -Value{
                $gobjSQL.Connection.Close()

                ## Remove Members No Longer Needed
                $gobjSQL.PSObject.Members.Remove("ServerName")
                $gobjSQL.PSObject.Members.Remove("DatabaseName")
                $gobjSQL.PSObject.Members.Remove("UserName")
                $gobjSQL.PSObject.Members.Remove("SecureFileLocation")
                $gobjSQL.PSObject.Members.Remove("CommandTimeOut")
                $gobjSQL.PSObject.Members.Remove("Connection")

                $gobjSQL.PSObject.Members.Remove("ExecuteCommandOnly")
                $gobjSQL.PSObject.Members.Remove("ExecuteCommandReturnDataSet")
                $gobjSQL.PSObject.Members.Remove("SQLDataSet")
                $gobjSQL.PSObject.Members.Remove("BulkUpdate")
                $gobjSQL.PSObject.Members.Remove("VarDump")
                $gobjSQL.PSObject.Members.Remove("Close")
            }

            Add-Member -InputObject $gobjSQL -MemberType NoteProperty -Name CommandTimeOut -value [Int]$CommandTimeOut
            $gobjSQL.CommandTimeOut = 0  # No Time Limit, Default is 30 Seconds

            Add-Member -InputObject $gobjSQL -MemberType ScriptMethod -Name "ExecuteCommandOnly" -Value{
                param (	
                    [Parameter(Mandatory=$true)]
                    [ValidateNotNullOrEmpty()]
                    [String]$SQLQuery
                )

                $SqlCMD = New-Object System.Data.SqlClient.SqlCommand($SQLQuery,$gobjSQL.Connection)
                $SqlCMD.CommandTimeout = $gobjSQL.CommandTimeOut
                $Results = $SQLCMD.ExecuteNonQuery()

                Return $Results
            }

            Add-Member -InputObject $gobjSQL -MemberType ScriptMethod -Name "ExecuteCommandReturnDataSet" -Value{
                param (	
                    [Parameter(Mandatory=$true)]
                    [ValidateNotNullOrEmpty()]
                    [String]$SQLQuery
                )

                try{
                        $SQLCMD = New-Object System.Data.SqlClient.SqlCommand($SQLQuery,$gobjSQL.Connection)
                        $SqlCMD.CommandTimeout = $gobjSQL.CommandTimeOut
                        $SqlReader = $SQLCMD.ExecuteReader()
                        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
                        $SqlAdapter.SelectCommand = $SqlCmd
                        $DataSet = New-Object System.Data.DataSet
                        $SqlReader.CLose()
                        $SqlAdapter.Fill($DataSet)

                        Add-Member -InputObject $gobjSQL -MemberType NoteProperty -Name SQLDataSet -value $DataSet
                }
                Catch{
                    #Error Found
                    #-----------
                    $sFunctionName = "gobjSQL"
                    $sFunctionSection = "ExecuteCommandReturnDataSet"
                    $sFunctionMessage = "Create Data Set"
                    $sFunctionMessage = FunctionMessage $sFunctionName $sFunctionSection $sFunctionMessage

                    Write-Warning $sFunctionMessage
                    $gobjLog.WriteLog($sFunctionMessage)

                    For($x=0; $x -lt $Error.Count; $x++){
                        $sErrorMsg = $Error[$x]
                        $gobjLog.WriteLog($sErrorMsg)
                        Write-Warning $sErrorMsg
                    }
                    #Reset Error Object
                    $Error.Clear()
                }

            }

            Add-Member -InputObject $gobjSQL -MemberType ScriptMethod -Name "BulkUpdate" -Value{
                param (	
                    [Parameter(Mandatory=$true)]
                    [ValidateNotNullOrEmpty()]
                    [String]$DestTable,
                    [String]$BatchSize,
                    [system.Data.DataTable]$BulkTable
                )

                #Create DB Connection
                $SQLBulkCopy = New-Object ("System.Data.SqlClient.SqlBulkCopy") $gobjSQL.Connection
                $SQLBulkCopy.DestinationTableName = $DestTable
                $SQLBulkCopy.BatchSize = $BatchSize
                #$SQLBulkCopy.BulkCopyTimeout = $Timeout #Previously set to 0 

                #foreach ($Column in $Source.columns.columnname){[void]$SQLBulkCopy.ColumnMappings.Add($Column, $Column)}
                $SQLBulkCopy.WriteToServer($BulkTable)
                $SQLBulkCopy.Close()
            }

            Add-Member -InputObject $gobjSQL -MemberType ScriptMethod -Name "VarDump" -Value{
                Write-Output $gobjSQL
            }

        }

        #Return if Successful
        return $bolInitialized
    } -passthru
    #--------------------------------------------
}