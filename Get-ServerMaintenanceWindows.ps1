﻿
<#
.Synopsis
   Server and Maintenance Windows 
.DESCRIPTION
   The script generates a html-document 
.EXAMPLE
   Get-ServerMaintenanceWindows -Email No -PDF Yes
.EXAMPLE
   Another example of how to use this workflow
.INPUTS
   Inputs to this workflow (if any)
.OUTPUTS
   Output from this workflow (if any)
.NOTES
   Version 1.0
.FUNCTIONALITY
   The functionality that best describes this workflow
#>


# Read Settings.xml
[XML]$Settings = Get-Content "\\kvv.se\app\CMPackages\Scripts\ServersAndMaintenanceWindows\Settings.xml"

# Fuctions
function Get-CMModule
{
	[CmdletBinding()]
	param ()
	Try
	{
		Write-Verbose "Attempting to import SCCM Module"
		Import-Module (Join-Path $(Split-Path $ENV:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1) -Verbose:$false
		Write-Verbose "Successfully imported the SCCM Module"
	}
	Catch
	{
		Throw "Failure to import SCCM Cmdlets."
	}
}

$logpathSetting = $Settings.Configuration.LogPath

$detailLog = "$logpathSetting" +"ServerMaintPlan_" +(Get-Date -Format yyyyMMdd) +".log" 

# Function for append events to logfile located c:\windows\logs
Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]
    $Level = "INFO",

    [Parameter(Mandatory=$True)]
    [string]
    $Message,

    [Parameter(Mandatory=$False)]
    [string]
    $logfile
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"
    If($logfile) {
        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}

        Write-Log -Level INFO -Message "Executing script to generate a Server Maintenance Plan for the current month" -logfile $detailLog

# Get today date
$Today = get-date
$DateTime = Get-DAte -Format yyMMdd_hhmm
 Write-Log -Level INFO -Message "Start job - Create summary Maintenance Windows for Servers" -logfile $detailLog

  


# Create csv-filename
$filenameCSV = "MaintancePlan_Servers_$DateTime.csv"
Write-Log -Level INFO -Message "csv-filename: $filenameCSV" -logfile $detailLog


# Numbers of days back in time to check Maintenance Windows
$DaysMin = $Settings.Configuration.DaysMin
$YearMonthMin = $Today.AddDays($DaysMin)

# Numbers of days in future to check Maintenance Windows
$DaysMax = $Settings.Configuration.DaysMax
$YearMonthMax = $Today.AddDays($DaysMax)

# Arrays - create 
$collectionmembers = @()
$result = @()

# Read configfile for deployments
$settingsXML = $Settings.Configuration.Data
$Deployments = Import-Csv -Path $settingsXML -Delimiter ";" -Encoding UTF7

# Get the powershell module for MEMCM
if (-not(Get-Module -name ConfigurationManager)) {
    Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')
}

$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-Location -Path "$($SiteCode.Name):\"


foreach ($Deployment in $Deployments)


{ 
   $collectionmembers =  Get-CMCollection -CollectionId $Deployment.collectionid | Get-CMCollectionMember 

   $mw = Get-CMMaintenanceWindow -CollectionId $Deployment.CollectionID | Where-Object {($_.starttime -ge $YearMonthMin) -and ($_.starttime -le $YearMonthMax)}

   $collection = Get-CMCollection -CollectionId $Deployment.CollectionID

   Set-Location -Path c:\
   $Deploy = $Deployment.LongDescription
   Write-Log -Level INFO -Message "Working with $deploy" -logfile $detailLog
   $SiteCode = Get-PSDrive -PSProvider CMSITE
    Set-Location -Path "$($SiteCode.Name):\"
   
    foreach ($member in $collectionmembers)
            
                    {
                      
                         Set-Location -Path c:\
                         $server = $member.Name
                         Write-Log -Level INFO -Message "Working with $server" -logfile $detailLog
                         $SiteCode = Get-PSDrive -PSProvider CMSITE
                         Set-Location -Path "$($SiteCode.Name):\"

                                            
                    
                    $myobj = New-Object -TypeName PSObject
                    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Server Name" -Value $member.name
                    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Collection Name" -Value $collection.name
                    #Add-Member -InputObject $myobj -MemberType NoteProperty -Name "CollectionID" -Value $collection.CollectionID
                    add-member -InputObject $myobj -MemberType NoteProperty -Name "Maintenance Windows" -Value $mw.Name
                    add-member -InputObject $myobj -MemberType NoteProperty -Name "StartTime" -Value $mw.starttime
                    add-member -InputObject $myobj -MemberType NoteProperty -Name "Duration" -Value $mw.Duration
                    Add-Member -InputObject $myobj -MemberType NoteProperty -Name "Description" -Value $Deployment.LongDescription
                    $result += $myobj
                    }
}


Set-Location -Path c:\

    Write-Log -Level INFO -Message "Creating html-format message to send by mail" -logfile $detailLog

$htmlFragment = @()
$newHtmlFragment = @()
$htmlFragment = $result| Sort-Object name | ConvertTo-Html -Fragment
    #$newHtmlFragment += $htmlFragment
    $newHtmlFragment += $htmlFragment.Replace('<th>',"<th class='tableheader'>")

    $BackgroundColor = $Settings.Configuration.BackgroundColor
    $TableHeaderColor = $Settings.Configuration.TableHeaderBGColor
    $tableHeaderTextColor = $Settings.Configuration.TableHeaderTextColor
    $Emailheadline = $settings.Configuration.Emailheadline
    $Generated = $Settings.Configuration.Generated

    $html = @"
<html lang='en'>
    <head>
        <meta charset='UTF-8'>
        <meta http-equiv='X-UA-Compatible' content='IE=edge'>
        <meta name='viewport' content='width=device-width, initial-scale=1.0'>
        <title>Servers - Maintenance Windows, Starttime and duration</title>
        <style>
            body {
                font-family: Calibri, sans-serif, 'Gill Sans', 'Gill Sans MT', 'Trebuchet MS';
                background-color: $BackgroundColor;
            }
            .mainhead {
                margin: auto;
                width: 100%;
                text-align: Left;
                font-size: X-large;
                font-weight: bolder;
            }
            table {
                margin: 10px auto;
                width: 70%;
            }
            .tableheader {
                background-color: $TableHeaderColor;
                color: $tableHeaderTextColor;
                padding: 10px;
                text-align: left;
                /* font-size: large; */
                border-bottom-style: solid;
                border-bottom-color: darkgray;
            }
            td {
                background-color: white;
                border-bottom: 1px;
                border-bottom-style: solid;
                border-bottom-color: #404040;
                Font-size: 12px
            }

            span {
                color: black;
            }

            .td1 {
                background-color: #F0F0F0;
            }
        </style>
    </head>
    <body>
        <div class='mainhead'>
            <img style='vertical-align: middle;' src='data:image/jpeg;base64,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

' alt='Disk Usage' width='50' height='80'/>
            $Emailheadline
        </div>
        <br>
        <div><i><b>$Generated </b> $((Get-Date).DateTime)</i></div>
        $newHtmlFragment
    </body>
</html>
"@


Write-Log -Level INFO -Message "Creating CSV-file to attach in mail" -logfile $detailLog

$result | Export-Csv -Path "c:\temp\$filenameCSV" -Encoding UTF8 -NoClobber -Force

#$emailRecipients = "lars.garlin@kriminalvarden.se","anders.m.olsson@kriminalvarden.se","christian.brask@kriminalvarden.se"

$FileURI = "c:\temp\$filenameCSV"
$Recipients = $Settings.Configuration.EmailRecipients
$Noreply = $Settings.Configuration.Noreplyaddress
$subject = $Settings.Configuration.EmailSubject

Write-Log -Level INFO -Message "Sending mail to $emailRecipients" -logfile $detailLog
Send-MailMessage -From $Noreply -To $Recipients -Subject $subject -BodyAsHtml $html -Encoding UTF8 -SmtpServer "smtp.kvv.se" -Attachments $FileURI
Write-Log -Level INFO -Message "Script finished!" -logfile $detailLog


 