#Created to run all my reports monthly
#Creation date : 10-25-2016
#Creator: Alix N Hoover
# 2-5-2018 added for each to tackle multi domains
# 2-5-2018 created Email report. 
#Added multi domain support

Import-Module ActiveDirectory

#Variables to configure
$MailServer = "mailserver"
#if only one domain only put in one domain controller
$Domainservers  = "dominserver1","domainerver2"

$recip = "me@you.org"
$sender = "Powershell@you.org"
$subject = "Last 30 days Audit"

#name of the machine for the schedule task
$task= "server with monthly task to run"


$today = get-date
$ScriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition
$rundate = ([datetime]$today).tostring("MM_dd_yyyy")
$30daysago = $($today.adddays(-32)).toshortdatestring() 
$year = ([datetime]$today).tostring("yyyy")
$month = ([datetime]$today).tostring("MM")
$startdateX = ([datetime]$30daysago).tostring("yyyy_MM_dd")
$enddateX = ([datetime]$today).tostring("yyyy_MM_dd")
$day = ([datetime]$today).tostring("dd")

# Set Directory Path
$Directory = $ScriptPath + "\Reports\Audit\"+ $year + "\" + $month + "\" + $day
# Create directory if it doesn't exsist
if (!(Test-Path $Directory))
{
New-Item $directory -type directory
}

#File Names
$fileName = $Directory +"\exchange2010Report "+( get-date ).ToString('MM_dd_yyyy')+".html"
$outfile = $Directory + "\email_stats_" + $startdateX + " to " + $EndDateX + ".csv" 
$dl_stat_file = $Directory + "\DL_stats____" + $startdateX + " to " + $EndDateX + ".csv"
$htmlfilename = $Directory +"\Last30Dayreport" + $rundate + ".html" 


# HTML start
New-Item -ItemType file $htmlfilename -Force
Add-Content $htmlfilename "<html>"
Add-Content $htmlfilename "<head>"
Add-Content $htmlfilename "</head>"
Add-Content $htmlfilename "<body>"
Add-Content $htmlfilename "<table align='center' border='1'>"
Add-Content $htmlfilename "<tr bgcolor='#32CD32'><td>"
Add-Content $htmlfilename "This is your 30 Day AD Audit Report."
Add-Content $htmlfilename "</td></tr></table>"



Foreach ($Server in $Domainservers)
{

#get accounts that have been created in the last 30 days
$tempusers = Get-ADUser -server $Server -Filter * -Properties * | Where-Object {$_.whenCreated -ge ((Get-Date).AddDays(-30)).Date} | sort-object WhenCreated |Select-Object Name, SamAccountName, EmailAddress, WhenCreated 
#get Computers that have been created in the last 30 days
$tempcomputers = Get-ADcomputer -server $Server -Filter * -Properties whenCreated | Where-Object {$_.whenCreated -ge ((Get-Date).AddDays(-30)).Date} | sort-object WhenCreated |Select-Object DNSHostName, WhenCreated
#get everything that has been deleted in last 30 days
$tempdelete = get-adobject -server $Server -filter {(isdeleted -eq $true) -and (objectclass -ne "container") -and (objectclass -ne "dnsnode")} -IncludeDeletedObjects -Properties * | sort-object ObjectClass, WhenChanged | Select-Object CN, WhenCreated, WhenChanged, ObjectClass

# TABLE Users START
Add-Content $htmlfilename "<p> "
Add-Content $htmlfilename "<table width='80%' align='center' border='1'>"
Add-Content $htmlfilename "<tr bgcolor='#32CD32'>"
Add-Content $htmlfilename "<td width='20%'>Name of New User on $server</td>"
Add-Content $htmlfilename "<td width='15%'>Account Name</td>"
Add-Content $htmlfilename "<td width='15%'>EmailAddress</td>"
Add-Content $htmlfilename "<td width='15%'>When Created on $server</td>"
Add-Content $htmlfilename "</tr>"

foreach ($user in $tempusers)
{
Add-Content $htmlfilename "<tr><td>"
Add-content $htmlfilename $user.name 
Add-Content $htmlfilename "</td><td>"
Add-content $htmlfilename $user.SamAccountName
Add-Content $htmlfilename "</td><td>"
Add-content $htmlfilename $user.EmailAddress
Add-Content $htmlfilename "</td><td>"
Add-Content $htmlfilename $user.WhenCreated 
Add-Content $htmlfilename "</td></tr>"
}
Add-Content $htmlfilename "</table>"
Add-Content $htmlfilename "</p>"

# TABLE Computers START
Add-Content $htmlfilename "<p>"
Add-Content $htmlfilename "<table width='80%' align='center' border='1'>"
Add-Content $htmlfilename "<tr bgcolor='#32CD32'>"
Add-Content $htmlfilename "<td width='55%'>Name of New Computer on $server</td>"
Add-Content $htmlfilename "<td width='45%'>When Created on $server</td>"
Add-Content $htmlfilename "</tr>"

foreach ($computer in $tempcomputers)
{
Add-Content $htmlfilename "<tr><td>"
Add-Content $htmlfilename $computer.DNSHostName
Add-Content $htmlfilename "</td><td>"
Add-Content $htmlfilename $Computer.WhenCreated
Add-Content $htmlfilename "</td></tr>"
}
Add-Content $htmlfilename "</table>"
Add-Content $htmlfilename "</p>"

 #TABLE Deleted START
Add-Content $htmlfilename "<p>"
Add-Content $htmlfilename "<table width='80%' align='center' border='1' >"
Add-Content $htmlfilename "<tr bgcolor='#32CD32'>"
Add-Content $htmlfilename "<td width='55%'>Deleted Object on server $Server</td>"
Add-Content $htmlfilename "<td width='15%'>When Created </td>"
Add-Content $htmlfilename "<td width='15%'>When Changed </td>"
Add-Content $htmlfilename "<td width='15%'>ObjectClass </td>"

Add-Content $htmlfilename "</tr>"

foreach ($delete in $tempdelete)
{
Add-Content $htmlfilename "<tr><td>"
Add-Content $htmlfilename $delete.CN
Add-Content $htmlfilename "</td><td>"
Add-Content $htmlfilename $delete.WhenCreated
Add-Content $htmlfilename "</td><td>"
Add-Content $htmlfilename $delete.WhenChanged
Add-Content $htmlfilename "</td><td>"
Add-Content $htmlfilename $delete.ObjectClass
Add-Content $htmlfilename "</td></tr>"
}
Add-Content $htmlfilename "</table>"
Add-Content $htmlfilename "</p>"
Add-Content $htmlfilename "</p>"
Add-Content $htmlfilename "</p>"



}
Add-Content $htmlfilename "<table width='50%' align='center'>"
Add-Content $htmlfilename "<tr><td>This File's Location is </td><td>"
Add-Content $htmlfilename $htmlfilename
Add-Content $htmlfilename "</td></tr><td>It is scheduled to run on</td><td>"
Add-Content $htmlfilename $task
Add-Content $htmlfilename "</td></tr></table>"

$Body = (Get-Content $htmlfilename) -join "<BR>"
Send-MailMessage -From $sender -To $recip -Subject $subject -Body $Body -BodyAsHtml -SmtpServer $MailServer

