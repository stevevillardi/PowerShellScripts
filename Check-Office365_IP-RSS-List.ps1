<#
.DESCRIPTION

###############Disclaimer#####################################################
This script is provided AS IS without warranty of any kind. In no event shall 
ePlus, its authors, or anyone else involved in the creation, production, or 
delivery of the scripts be liable for any damages whatsoever (including, 
without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use 
of or inability to use the sample scripts or documentation, even if ePlus 
has been advised of the possibility of such damages.
###############Disclaimer#####################################################

This script will poll Office365 IP Change RSS Feed and return any new changes based on current date.

=========================================
Version: 03232018

Author: 
Steve Villardi - svillardi@eplus.com
=========================================
.PARAMETER From
Email Address message comes from.
.PARAMETER To
Email Address message is sent to.
.PARAMETER SMTPServer
SMTP server that can relay the message.
.PARAMETER SMTPPort
SMTP port used to connect to SMTP server.
.PARAMETER UseSSL
Connect using SSL.
.PARAMETER UserName
Username to authenticate to the SMTP server with.
.PARAMETER Password
Password to authenticate to the SMTP server with.


#>

param(
    [string]$From,
    [string]$To,
    [string]$SMTPServer,
    [string]$SMTPPort,
    [switch]$UseSSL,
    [string]$UserName,
    [string]$Password
)

################################################################################
# Get RSS Content to XML file for parsing
################################################################################
function SendNotification{
    $Msg = New-Object Net.Mail.MailMessage
    $Smtp = New-Object Net.Mail.SmtpClient($SMTPServer,$SMTPPort)
    If($UseSSL){
    $Smtp.EnableSsl = $true
    }
    Else{
    $Smtp.EnableSsl = $false
    }
    $Smtp.Credentials = New-Object Net.NetworkCredential($username, $password)
    $Msg.From = $From
    $Msg.To.Add($To)
    $Msg.Subject = $Subject
    $Msg.Body = $Body
    $Msg.IsBodyHTML = $true
    $Smtp.Send($Msg)
    #Cleanup before sending again
    $msg.Dispose();
}

$XML_Path = $PSScriptRoot + "\Office_365_IP_List.xml"

If(Test-Path $XML_Path){
    Remove-Item $XML_Path
}

$now = (Get-Date).AddDays(-2)

Invoke-WebRequest -Uri "https://support.office.com/en-us/o365ip/rss" -OutFile $XML_Path

[xml]$RSS_Document = Get-Content -Path $XML_Path
$Feed = $RSS_Document.rss.channel
$RSS_Body = @()
$RSS_Count = 0
Foreach($msg in $Feed.Item){
    $msg_date = [datetime]$msg.pubDate
    If($msg_date -ge $now){
        $Parse_Title = $msg.title -replace "\n"," "
        $Parse_Description = $msg.description -replace "\n"," "
        $RSS_Count +=1
        #Write-Host "============================================================" -ForegroundColor Gray
        Write-Host "Found new RSS entry" -ForegroundColor Green
        $pd = "    Published Date:    $($msg.pubDate)"
        Write-Host $pd -ForegroundColor Yellow
        $title = "    Product Effected:  $Parse_Title"
        Write-Host $title -ForegroundColor Yellow
        $link = "    Link:              $($msg.link)"
        Write-Host $link -ForegroundColor Yellow
        $description = "    Description:       $Parse_Description"
        Write-Host $description -ForegroundColor Yellow
        Write-Host "============================================================" -ForegroundColor Gray
        $RSS_Body += New-Object -TypeName psobject -Property @{'PubDate' = $pd;'Description' = $description;'Title' = $title;'Link' = $link}
    }
}

If(($From -ne "") -and ($To -ne "") -and ($SMTPServer -ne "") -and ($SMTPPort -ne "")){
    $Subject = "Office365 IP RSS Check - $RSS_Count New Entries Detected"
    $Body = "<b>Office365 IP/URL Changes, Additions and Deletions in the past two days since $now</b><br>"
    $Body += "<br>"
    $Body += "Found $RSS_Count New RSS Entries: <br>"
    $Body += "=========================================<br>"
    Foreach($object in $RSS_Body){
        $Body += "  $($object.PubDate)<br>"
        $Body += "  $($object.Title)<br>"
        $Body += "  $($object.Link)<br>"
        $Body += "<br>"
        $Body += "  $($object.Description)<br>"
        $Body += "=========================================<br>"
    }
    If($RSS_Count -gt 0){
        Try{
            SendNotification
            Write-Host "Successfully sent results to $to" -ForegroundColor Green
        }
        Catch{
            Write-Host "Failed to send results to $to" -ForegroundColor Red
        }
    }
    Else{
        Write-Host "Skipping Email Notification since no new entries detected" -ForegroundColor Gray
    }
}
Else{
    If($RSS_Count -gt 0){
        Write-Host "Skipping Email Notification since require switches not present" -ForegroundColor Gray
    }
    Else{
        Write-Host "No new entries in the past 2 days detected using $now as the filter date" -ForegroundColor Gray
    }
    
}
