<#
.NOTES   
    Name: SC-ScurePassword.ps1
    Version: 1.0.0
    Author: Kevin McClean
    DateCreated: 05/01/2017
    DateUpdated: 05/8/2017

.SYNOPSIS   
    Encrypting a String of Data to a File.

.DESCRIPTION
    Encrypt data to a file but note that it is user specific!.

.PARAMETERS [Param]


   
.Dependencies
    SecurePasswordObject Ver 1.0.0

.Resources
    Powershell
        https://gallery.technet.microsoft.com/scriptcenter/Secure-Password-using-c158a888

.Version Notes

#>

#Import Modules


#Global Variables
#----------------
    [String]$gServerName = (Get-WmiObject -Class Win32_ComputerSystem).name
    [String]$gAppName = "SecurePassword"
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


    #Variables
    #---------


    ### Place Code Here
    #------------------
    $sFunctionSection = "SecurePassword Object"




    Try{
        Write-Warning "SPECIAL NOTE"
        Write-Warning "------------"
        Write-Warning "This method has some limitation that it will only work for the same user on the same machine."
        Write-Warning "No other user or profile on same machine or any other machine can read or decrypt this file."
        Write-Warning "Decryption can only be done via same user on same machine."
        
        ##Create Encrypted Password File
        $MyFile = "C:\Scripts\Test.cred"
        Write-Host "Creating Encrypted Password File C:\Scripts\Test.cred"
        Write-Host "You will be prompted to enter a Password in 2 Seconds."
        Start-Sleep -s 2
        $objSecurePwd.Initialize($MyFile, $true)

        ##Retrieve Encrypted Password
        Write-Host "Look an Encrypted Password: " $objSecurePwd.GetPassword()

        ##Convert to Plain Text (Not recommended)
        Write-Warning "THIS IS NOT RECOMENDED"
        Function ConvertFrom-SecureToPlain{ 
            param( [Parameter(Mandatory=$true)][System.Security.SecureString] $SecurePassword) 
 
            # Create a "password pointer" 
            $PasswordPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword) 
 
            # Get the plain text version of the password 
            $PlainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto($PasswordPointer) 
 
            # Free the pointer 
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($PasswordPointer) 
 
            # Return the plain text password 
            return $PlainTextPassword 
        } 
        $MyPassword = ConvertFrom-SecureToPlain $objSecurePwd.GetPassword()
        Write-Warning "Here is the Password as Plain Text: $MyPassword"
    }
    Catch{
        Write-Warning $Error[0]
    }
    

#End
#---------
$Timer = (Get-Date) - $ScriptDateStart
Write-Warning "Function '$sFunctionName'. Total Execution Time: $Timer"
Write-Host "--- End of Script: ---"
