#----------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Variables
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Stops script if error occurs
$ErrorActionPreference = "Stop"

# Clear the $Error variable
$Error.Clear()

# Create file paths
$importFile = "$PSScriptRoot\<chris21-export>.xml"
$logsDir = "$PSScriptRoot\Logs"
$supplementaryLogDir = "$PSScriptRoot\Logs\Supplementary Logs"
$archiveDir = "$PSScriptRoot\Archives"
$statusLogFile = "$PSScriptRoot\Logs\Log.txt"
$errorLogFile = "$PSScriptRoot\Logs\ErrorLog.txt"

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Functions
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Formats timestamp
Function Get-TimeStamp 
{    
    Return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)   
}

# Checks solution structure integrity. Corrects if needed.
Function Check-ScriptStructure
{
    If(!(Test-Path -Path $logsDir))
    {
        $null = New-Item -ItemType Directory -Path $logsDir -Force
    }

    If(!(Test-Path -Path $statusLogFile))
    {
        $null = New-Item -ItemType File -Path $statusLogFile -Force
    }

    If(!(Test-Path -Path $errorLogFile))
    {
        $null = New-Item -ItemType File -Path $errorLogFile -Force
    }

    If(!(Test-Path -Path $supplementaryLogDir))
    {
        $null = New-Item -ItemType Directory -Path $supplementaryLogDir -Force
    }

    If(!(Test-Path -Path $archiveDir))
    {
        $null = New-Item -ItemType Directory -Path $archiveDir -Force
    }
}

# Constructs XML body to pass onto API. NOTE: Element order is important and must be in the order below!
Function Build-XMLBody
{
    $bdy = "
    <UserImport>
    <Username>$userName</Username>
    <Email>$email</Email>
    <FirstName>$firstName</FirstName>
    <LastName>$lastName</LastName>
    <Title>$title</Title>
    <CustomField1>$customField1</CustomField1>
    <CustomField2>$customField2</CustomField2>
    <CustomField3>$customField3</CustomField3>
    <CustomField4>$customField4</CustomField4>
    <CustomField5>$customField5</CustomField5>
    <CustomField6>$customField6</CustomField6>
    <CustomField7>$customField7</CustomField7>
    <CustomField8>$customField8</CustomField8>
    <Active>$active</Active>
    <Manager>$manager</Manager>
    </UserImport>"
   
    Return $bdy
}

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Main
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------

Try
{
    Check-ScriptStructure

    #Cast file text to XML object
    [xml]$xmlAttr = Get-Content -Path $importFile

    # Loop through XML object data set and create manager lookup hashtable
    $lookupMgr = @{}
    $xmlAttr.root.Data | ForEach-Object{
        $lookupMgr.Add($_."detnumber", $_."detemailad")
    }

    # Loop through XML object data set and set up XML export elements
    $xmlAttr.root.Data | ForEach-Object{
        # Output result object, set output elements to match destination fields in Litmos
        $obj = [pscustomobject]@{
            "UserName" = $_."detemailad"
            "Email" = $_."detemailad"
            "FirstName" = $_."detg1name1"
    	    "LastName" = $_."detsurname"
            "Title" = $_."posnumber.trn"
            "ManagerName" = $_."detcurman.trn"
            "CustomField1" = $_."detnumber"
            "CustomField2" = $_."possalgrp.trn"
            "CustomField3" = $_."posl2cd.trn"
    	    "CustomField4" = $_."posl3cd.trn"
            "CustomField5" = $_."pydlocncd.trn"
	        "CustomField6" = $_."posstatus.trn"
    	    "CustomField7" = $_."detdatejnd"
    	    "CustomField8" = $_."terdate"
            "ManagerId" = $_."detcurman"
        }  

        # Set up XML body variables
        $userName = $obj.UserName
        $email = $obj.Email
        $firstName = $obj.FirstName
        $lastName = $obj.LastName
        $title = $obj.Title
        $customField1 = $obj.CustomField1
        $customField2 = $obj.CustomField2
        $customField3 = $obj.CustomField3
        $customField4 = $obj.CustomField4
        $customField5 = $obj.CustomField5
        $customField6 = $obj.CustomField6
        $customField7 = $obj.CustomField7
        $active = "true"
        $managerId = $obj.ManagerId
        $managerName = $obj.ManagerName

        # Set manager email via hashtable lookup
        $manager = $null
        ForEach($key in $lookupMgr.Keys)
        {
            If($key -eq $obj.ManagerId) 
            {
                $manager = $lookupMgr[$key]
            }
        }

        # Replace empty XML element with null
        $customField8 = $obj.CustomField8
        If($customField8.IsEmpty)
        {
            $customField8 = $null
        }

        # Build XML body to be passed to API
        $currentAccount = Build-XMLBody($userName, $email, $firstName, $lastName, $title, $customField1, $customField2, $customField3, $customField4, $customField5, $customField6, $customField7, $customField8, $active, $manager)
        $bodyString = $bodyString + $currentAccount
    }

    # Bulk POST to API: Sendmessage=false, skipfirstlogin=true, Relplace '&' as this breaks the API. Explicitly allowing TLS, TLS 1.1 and TLS 1.2.
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

    $apiKey = "<api-key>"
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("apiKey", $apiKey)
    $headers.Add("Content-Type", "application/xml")
    $body = "<UserImports>$bodyString`n</UserImports>"
    $cleanBody = $body -replace "&", "and"
    $response = Invoke-RestMethod 'https://api.litmos.com.au/v1.svc/bulkimports?source=Chris21_Integration&sendmessage=false&skipfirstlogin=true' -Method 'POST' -Headers $headers -Body $cleanBody
    $cleanBody | Out-File -FilePath $supplementaryLogDir\LitmosXMLBody_$(Get-Date -f yyyy-MM-dd_HHmmss).txt
   
    # Archive a copy of the export
    Move-Item -Path $importFile -Destination $archiveDir\$importFile_$(Get-Date -f yyyy-MM-dd_HHmmss).xml

    # Delete archives older than 60 days
    Get-ChildItem $archiveDir -Recurse -File | Where CreationTime -lt (Get-Date).AddDays(-60) | Remove-Item -Force
}
Catch 
{
    # Log error details
    "$(Get-TimeStamp) `t Error! See ""ErrorLog.log"" for details." | Out-File $statusLogFile -Append
}
Finally
{
    # Log script runtime and status
    If(!($Error))
    {
        "$(Get-TimeStamp) `t Success." | Out-File $statusLogFile -Append
    }
    Else
    {    
        "$(Get-TimeStamp) $Error `n" | Out-File $errorLogFile -Append
    }
}
