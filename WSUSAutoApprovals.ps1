# Windows Update Patch Approval Script
# Copyright (C) 2018 Steve Lunn (gilgamoth@gmail.com)
# Downloaded From: https://github.com/Gilgamoth

# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see https://www.gnu.org/licenses/

$App_Version = "2015-09-17-1145"

Clear-Host
Set-PSDebug -strict
$ErrorActionPreference = "SilentlyContinue"
[GC]::Collect()

# **************************** VARIABLE START ******************************

# Only needed if not localhost or not default port, leave blank otherwise.
    $Cfg_WSUSServer = "WSUSServer.FQDN.local" # WSUS Server Name
    $Cfg_WSUSSSL = $true # WSUS Server Using SSL
    $Cfg_WSUSPort = 8531 # WSUS Port Number

# E-Mail Report Details
	$Cfg_Email_To_Address = "recipient@domain.local"
	$Cfg_Email_From_Address = "WSUS-Report@wolftdomainech.local"
	$Cfg_Email_Subject = "WSUS: Auto Approval Report " + $env:computername
	$Cfg_Email_Server = "mail.domain.local"
    $Cfg_Email_Send = $false
    $Cfg_Email_Send = $true # Comment out if no e-mail required
	[string]$Email_Body = ""

# E-Mail Server Credentials to send report (Leave Blank if not Required)
	$Cfg_Smtp_User = ""
	$Cfg_Smtp_Password = ""

$DeadlineDays=0
$PatchAge=6

$NUCount=0 #New Updates Count
$UTACount=0 #Updates To Approve Count
$AUCount=0 #Approved Updates Count
$WSUSSvr=""

$UpdateDeadline = (get-date).AddDays(+$DeadlineDays)
	
# ****************************** CODE START ******************************

$Today=(get-date).ToString("yyyy-MM-dd")

$StartTime = Get-Date

[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null 

$All = [Microsoft.UpdateServices.Administration.UpdateApprovalAction]::All  
$Install = [Microsoft.UpdateServices.Administration.UpdateApprovalAction]::Install  
$NotApproved = [Microsoft.UpdateServices.Administration.UpdateApprovalAction]::NotApproved  
$Uninstall = [Microsoft.UpdateServices.Administration.UpdateApprovalAction]::Uninstall 

$WSUSSvr = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($Cfg_WSUSServer, $Cfg_WSUSSSL, $Cfg_WSUSPort)

$Group = $WSUSSvr.GetComputerTargetGroups() | where {$_.Name -eq "All Computers"}

# Get the Updates that have arrived in the last 24 hours and if there are any, add their details to the report e-mail
$UpdateScope = new-object Microsoft.UpdateServices.Administration.UpdateScope
$UpdateScope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::NotApproved
$UpdateScope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::NotInstalled
$UpdateScope.FromArrivalDate = (get-date).AddDays(-1)

$Updates = $WSUSSvr.GetUpdates($UpdateScope)
$NUCount = $Updates.Count
write-host "$NUCount New Updates Found"
$Email_Body += "<u>$NUCount New Updates Found</u><br>"
If ($NUCount -gt 0) {
	$Cfg_Email_Send = $True
	ForEach ($Update in $Updates) {
		$UTitle = $update.Title
		write-host $update.Title
        $Email_Body += "$UTitle<br>"
    }
}

# Get the Updates that have are older than 7 days and approve them, adding their details to the report e-mail
$UpdateScope = new-object Microsoft.UpdateServices.Administration.UpdateScope
$UpdateScope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::NotApproved
$UpdateScope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::NotInstalled
$UpdateScope.ToCreationDate = (get-date).AddDays(-$PatchAge)

$Updates = $WSUSSvr.GetUpdates($UpdateScope)
$UTACount = $Updates.Count
write-host "`n$UTACount Updates to Approve"
$Email_Body += "<br><u>$UTACount Updates to Approve</u><br>"
If ($UTACount -gt 0) {
	$Cfg_Email_Send = $True
	ForEach ($Update in $Updates) {
		$UTitle = $update.Title
        $UArrival = $Update.ArrivalDate
        $UCreation = $Update.CreationDate
        write-host $update.Title -NoNewline
		If ($DeadlineDays -gt 0) {
			$Results = $Update.Approve($Install,$Group,$UpdateDeadline)
		} Else {
			$Results = $Update.Approve($Install,$Group)
		}
		If ($Results.GoLiveTime) {
			Write-Host " Approved in WSUS" -ForegroundColor Green
			If ($DeadlineDays -gt 0) {
        	    		$Email_Body += "$UTitle <font color=`"#009933`">approved</font> in WSUS with a deadline of $UpdateDeadline<br>`n - Created:`t$UCreation`tArrived:`t$UArrival<br>`n"
			} Else {
        	    		$Email_Body += "$UTitle <font color=`"#009933`">approved</font> in WSUS<br>`n - Created:`t$UCreation`tArrived:`t$UArrival<br>`n"
			}
			$AUCount++
		} Else {
			Write-Host " Not Approved in WSUS" -ForegroundColor Red
            $Email_Body += "$UTitle <font color=`"#FF0000`">not approved in WSUS</font><br>`n - Created:`t$UCreation`tArrived:`t$UArrival<br>`n"
		}
	}
}
write-host "`n$AUCount Updates Approved"

[GC]::Collect()

$EndTime = Get-Date
Write-Host "`nStart Time: $StartTime"
Write-Host "End Time: $EndTime"
$Email_Body += "<br><bStart Time:</b> ".$StartTime."<br><b>End Time:</b> ".$EndTime."<br>"

if ($Cfg_Email_Send) {
    $smtp = New-Object System.Net.Mail.SmtpClient -argumentList $Cfg_Email_Server
    if ($Cfg_Smtp_User) {
        $smtp.Credentials = New-Object System.Net.NetworkCredential -argumentList $Cfg_Smtp_User,$Cfg_Smtp_Password
    }
    $message = New-Object System.Net.Mail.MailMessage
    $message.From = New-Object System.Net.Mail.MailAddress($Cfg_Email_From_Address)
    $message.To.Add($Cfg_Email_To_Address)
    $message.Subject = $Cfg_Email_Subject + " - $NUCount New, $UTACount Updates, $AUCount Approved - " + (get-date).ToString("dd/MM/yyyy")
    $message.isBodyHtml = $true
    $message.Body = $Email_Body
    Write-Host "`nSending Report E-Mail to" $message.To.address
    $smtp.Send($message)
}
