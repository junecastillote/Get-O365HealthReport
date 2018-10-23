<#	
	.NOTES
	===========================================================================
	 Created on:   	7-August-2018
	 Created by:   	June Castillote
					june.castillote@gmail.com
	 Filename:     	Get-o365HealthReport.ps1
	 Version:		1.0 (7-August-2018)
	===========================================================================

	.LINK
		https://www.lazyexchangeadmin.com/2018/10/shd365.html
		https://github.com/junecastillote/Get-o365HealthReport

	.SYNOPSIS
		This script utilize the Office 365 Management API v2 to retrieve the service health status
		and the Microsoft Graph API to send the report thru email using an Office 365 Mailbox.

	.DESCRIPTION
		For more details and usage instruction, please visit the link:
		https://www.lazyexchangeadmin.com/2018/10/shd365.html
		https://github.com/junecastillote/Get-o365HealthReport
		
		
		
	.EXAMPLE
		.\Get-o365HealthReport.ps1

#>

#Requires -Version 4.0
$scriptVersion = "1.0"

#get root path of the script
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
Start-Transcript -Path "$($script_root)\log.txt" -Append
$WarningPreference = "SilentlyContinue"

#Selected Status to report
$statusStringArray = @(
"Investigating",
"Service degradation",
"Service interruption",
"Restoring service",
"Extended recovery",
"Investigation suspended",
"Service restored"
)

#import config.xml
[xml]$config = Get-Content "$($script_root)\resource\config.xml"

#csv files
$testCSv = "$($script_root)\output\test_data.csv"
$oldCSV = "$($script_root)\output\old.csv"
$newCSV = "$($script_root)\output\new.csv"
$updatedCSV = "$($script_root)\output\updated.csv"
$css_string = Get-Content "$($script_root)\resource\style.css"
$css_string = $css_string -join "`n" #convert to multiline string

#if testMode=true in config.xml, this will populate the initial seed with test data
#set testMode=false to work with realtime data
if ( $config.options.testMode -eq $true){
	Write-Host "---Test Mode---"
	if (Test-Path $oldCSV){
		Remove-Item $oldCSV -Force -Confirm:$false
		Copy-Item -Path $testCSv -Destination $oldCSV
	}
}

#email settings
$sendEmail = $config.options.sendEmail
[array]$toAddress = ($config.options.toAddress).Split(",")
$fromAddress = $config.options.fromAddress

#base64 images
[string]$base64_healthy = '"'+[convert]::ToBase64String((Get-Content "$($script_root)\output\image\healthy.png" -Encoding byte))+'"'
[string]$base64_incident = '"'+[convert]::ToBase64String((Get-Content "$($script_root)\output\image\incident.png" -Encoding byte))+'"'
[string]$base64_advisory = '"'+[convert]::ToBase64String((Get-Content "$($script_root)\output\image\advisory.png" -Encoding byte))+'"'

#Assign the Client/Application ID, Client Secret and Tenant Domain
$ClientID = $config.options.clientID
$ClientSecret = $config.options.clientSecret
$tenantdomain = $config.options.tenantDomain

# Office 365 Management API starts here
try {

	$body = @{grant_type="client_credentials";resource="https://manage.office.com";client_id=$ClientID;client_secret=$ClientSecret}
	$oauth = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($tenantdomain)/oauth2/token?api-version=1.0" -Body $body
	$headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
	$messages = (Invoke-RestMethod -Uri "https://manage.office.com/api/v1.0/$($tenantdomain)/ServiceComms/Messages" -Headers $headerParams -Method Get)
	#$services = (Invoke-RestMethod -Uri "https://manage.office.com/api/v1.0/$($tenantdomain)/ServiceComms/Services" -Headers $headerParams -Method Get -Verbose)
	#$currentStatus = (Invoke-RestMethod -Uri "https://manage.office.com/api/v1.0/$($tenantdomain)/ServiceComms/CurrentStatus" -Headers $headerParams -Method Get -Verbose)
	#$historicalstatus = (Invoke-RestMethod -Uri "https://manage.office.com/api/v1.0/$($tenantdomain)/ServiceComms/HistoricalStatus" -Headers $headerParams -Method Get -Verbose)
	$incidents = $messages.Value | Where-Object {$_.MessageType -eq 'Incident'}
	#$messages
	}
catch {
	#$_.Exception.Message
	$_.Exception | Format-List
	#Stop-Transcript
	EXIT
}
# Office 365 Management API ends here

#compile new results
$newResult = @()
foreach ($message in $incidents){

	if ($statusStringArray -contains $message.Status) {
	#get the index of the latest message in the event.
	[int]$msgCount = ($message.Messages.Count)-1

	#build the NEW report object
	$temp = "" | Select-Object ID, Classification, Title, Workload, Status, StartTime, LastUpdatedTime, EndTime, ActionType, AffectedTenantCount, Message, MessageType, WorkloadDisplayName, Feature, FeatureDisplayName, PostIncidentDocumentUrl, Severity, ImpactDescription
	$temp.ID = $message.ID
	$temp.Title = $message.Title
	$temp.Classification = $message.Classification
	$temp.Workload = $message.Workload
	$temp.WorkloadDisplayName = $message.WorkloadDisplayName
	$temp.Status = $message.Status
	$temp.Feature = $message.Feature
	$temp.FeatureDisplayName = $message.FeatureDisplayName
	$temp.MessageType = $message.MessageType
	if ($message.StartTime) {$temp.StartTime = [datetime]$message.StartTime}
	if ($message.LastUpdatedTime) {$temp.LastUpdatedTime = [datetime]$message.LastUpdatedTime}
	if ($message.EndTime) {$temp.EndTime = [datetime]$message.EndTime}
	$temp.Message = $message.Messages[$msgCount].MessageText
	$temp.ActionType = $message.ActionType
	$temp.AffectedTenantCount = $message.AffectedTenantCount
	$temp.PostIncidentDocumentUrl = $message.PostIncidentDocumentUrl
	$temp.Severity = $message.Severity
	$temp.ImpactDescription = $message.ImpactDescription
	$newResult += $temp
	}
}
#process the retrieved records
$newResult | Sort-Object LastUpdatedTime -Descending | export-Csv -notypeInformation $newCSV

#this import is to makes sure that the comparison between new and old records are as accurate as possible
#if not imported, the comparison fails and all records are treated as new
$newResult = Import-Csv $newCSV
#=======

Write-Host "Retrieved Records: $($newResult.Count)"
$updatedRecord = @()

#import the old events
if (Test-Path $oldCSV){
	
	$oldResult = Import-Csv $oldCSV | Sort-Object Classification -Descending
	Remove-Item $oldCSV
	
	foreach ($newRecord in $newResult){
	$temp = "" | Select-Object EventType, ID, Classification, Title, Workload, Status, StartTime, LastUpdatedTime, EndTime, ActionType, AffectedTenantCount, Message, MessageType, WorkloadDisplayName, Feature, FeatureDisplayName, PostIncidentDocumentUrl, Severity, ImpactDescription
		
		#search ID (if exists)
		$oldRecord = $oldResult | Where-Object {$_.ID -eq $newRecord.ID}
		
		#if ID exists, compare ID
		if ($oldRecord) {			
			#check if ID is updated based on LastUpdatedTime
			if ($oldRecord.LastUpdatedTime -ne $newRecord.LastUpdatedTime) {
				$temp.EventType = "Update"
				$temp.ID = $newRecord.ID
				$temp.Title = $newRecord.Title
				$temp.Classification = $newRecord.Classification
				$temp.Workload = $newRecord.Workload
				$temp.WorkloadDisplayName = $newRecord.WorkloadDisplayName
				$temp.Status = $newRecord.Status
				$temp.Feature = $newRecord.Feature
				$temp.FeatureDisplayName = $newRecord.FeatureDisplayName
				$temp.MessageType = $newRecord.MessageType
				$temp.StartTime = $newRecord.StartTime
				$temp.LastUpdatedTime = $newRecord.LastUpdatedTime
				$temp.EndTime = $newRecord.EndTime
				$temp.Message = $newRecord.Message
				$temp.ActionType = $newRecord.ActionType
				$temp.AffectedTenantCount = $newRecord.AffectedTenantCount
				$temp.PostIncidentDocumentUrl = $newRecord.PostIncidentDocumentUrl
				$temp.Severity = $newRecord.Severity
				$temp.ImpactDescription = $newRecord.ImpactDescription
				$updatedRecord += $temp
			}
		}
		#if ID does not exist, new record
		else {
			
				$temp.EventType = "New"
				$temp.ID = $newRecord.ID
				$temp.Title = $newRecord.Title
				$temp.Classification = $newRecord.Classification
				$temp.Workload = $newRecord.Workload
				$temp.WorkloadDisplayName = $newRecord.WorkloadDisplayName
				$temp.Status = $newRecord.Status
				$temp.Feature = $newRecord.Feature
				$temp.FeatureDisplayName = $newRecord.FeatureDisplayName
				$temp.MessageType = $newRecord.MessageType
				$temp.StartTime = $newRecord.StartTime
				$temp.LastUpdatedTime = $newRecord.LastUpdatedTime
				$temp.EndTime = $newRecord.EndTime
				$temp.Message = $newRecord.Message
				$temp.ActionType = $newRecord.ActionType
				$temp.AffectedTenantCount = $newRecord.AffectedTenantCount
				$temp.PostIncidentDocumentUrl = $newRecord.PostIncidentDocumentUrl
				$temp.Severity = $newRecord.Severity
				$temp.ImpactDescription = $newRecord.ImpactDescription
				$updatedRecord += $temp
		}				
	}
	Write-Host "Updated Records: $($updatedRecord.Count)"
	$updatedRecord | Export-Csv -noTypeInformation $updatedCSV
	
	#create the report
	if ($updatedRecord) {
		$mail_Body1 = @()
		$mail_Body2 = @()
		Write-Host "Writing Report"
		$mailSubject = $config.options.mailSubject
		$mail_Body1 = "<html><head><title>$($mailSubject)</title>"
		$mail_Body1 = $mail_Body1 -join "`n" #convert to multiline string
		$mail_Body2 += "</head><body>"

		if ( $config.options.testMode -eq $true){
			$mail_Body2 += '<table id="HeadingInfo"><th>' + $mailSubject + '<br />' + $tenantdomain + '<br />' + ('{0:dd-MMM-yyyy H:mm}' -f (get-date)) + '<br />[TEST MODE]</th></table><hr>'
		}
		else {
			$mail_Body2 += '<table id="HeadingInfo"><th>' + $mailSubject + '<br />' + $tenantdomain + '<br />' + ('{0:dd-MMM-yyyy H:mm}' -f (get-date)) + '</th></table><hr>'
		}

		#$mail_Body2 += '<table id="HeadingInfo"><th>' + $mailSubject + '<br />' + $tenantdomain + '<br />' + ('{0:dd-MMM-yyyy H:mm}' -f (get-date)) + '</th></table><hr>'
		$mail_Body2 += '<table id="section"><tr><th width="2%"><img src="image/advisory.png" width="18" height="18"></th><th width="5%">Advisory</th><th width="2%"><img src="image/incident.png" width="18" height="18"></th><th width="5%">Incident</th><th width="2%"><img src="image/healthy.png" width="18" height="18"></th><th width="84%">Restored</th></tr></table>'
				
		foreach ($record in $updatedRecord)
		{	
			$mail_Body2 += "<hr>"
			#advisory + restored
			if ($record.Classification -eq 'Advisory' -and $record.Status -eq "Service Restored"){
				$mail_Body2 += '<table id="section"><tr><th width="5%"><img src="image/advisory.png" alt="Advisory"><img src="image/healthy.png" alt="Restored"></th><th width="95%">' + $record.WorkloadDisplayName +' | ' + $record.ID + ' | ' + $record.Title + '</th></tr></table>'
			}
			#advisory + ongoing
			elseif ($record.Classification -eq 'Advisory' -and $record.Status -ne "Service Restored"){
				$mail_Body2 += '<table id="section"><tr><th width="5%"><img src="image/advisory.png" alt="Advisory"></th><th width="95%">' + $record.WorkloadDisplayName +' | ' + $record.ID + ' | ' + $record.Title + '</th></tr></table>'
			}
			#incident + restored
			elseif ($record.Classification -eq 'Incident' -and $record.Status -eq "Service Restored") {
				$mail_Body2 += '<table id="section"><tr><th width="5%"><img src="image/incident.png" alt="Incident"><img src="image/healthy.png" alt="Restored"></th><th width="95%">' + $record.WorkloadDisplayName +' | ' + $record.ID + ' | ' + $record.Title + '</th></tr></table>'
			}
			#incident + ongoing
			elseif ($record.Classification -eq 'Incident' -and $record.Status -ne "Service Restored") {
				$mail_Body2 += '<table id="section"><tr><th width="5%"><img src="image/incident.png" alt="Incident"></th><th width="95%">' + $record.WorkloadDisplayName +' | ' + $record.ID + ' | ' + $record.Title + '</th></tr></table>'
			}
			$mail_Body2 += "<hr>"
						
			$mail_Body2 += '<table id="data">'
			$mail_Body2 += '<tr><th>Status</th><th>User Impact</th><th>Last Updated</th><th>Start</th><th>End</th><th>Lastest Message</th></tr>'
			if ($record.Status -eq 'Service Restored')
			{
				$mail_Body2 += '<tr><td width="10%" class="good">' + $record.Status + '</td>'
			}
			else {
				$mail_Body2 += '<tr><td width="10%" class="bad">' + $record.Status + '</td>'
			}
			$mail_Body2 += '<td width="20%">' + $record.ImpactDescription + '</td>'
			$mail_Body2 += '<td width="10%">' + ('{0:dd-MMM-yyyy H:mm}' -f $record.LastUpdatedTime) + '</td>'
			$mail_Body2 += '<td width="10%">' + ('{0:dd-MMM-yyyy H:mm}' -f $record.StartTime) + '</td>'
			$mail_Body2 += '<td width="10%">' + ('{0:dd-MMM-yyyy H:mm}' -f $record.EndTime) + '</td>'
			$mail_Body2 += '<td width="30%">' + ($record.Message).Replace("`n","<br />") + '</td></tr>'
			$mail_Body2 += '</table>'
		}
		$mail_Body2 += '<p><table id="section">'
		$mail_Body2 += '<tr><th><center>----END of REPORT----</center></th></tr></table></p>'
		$mail_Body2 += '<p><font size="2" face="Tahoma"><br />'
		$mail_Body2 += '<br />'
		$mail_Body2 += '<p><a href="https://www.lazyexchangeadmin.com/2018/10/shd365.html">Office365 Events Monitor v.'+ $scriptVersion +'</a></p>'
		$mail_body2 += '</body>'
		$mail_body2 += '</html>'
		$mail_Body2 = $mail_Body2 -join "`n"
		#combine body texts
		$mail_Body = $mail_Body1 + $css_string + $mail_Body2
		$mail_body | Out-File "$($script_root)\output\report.html"
		}
	
#send email if new or updated events are found
if ($updatedRecord.Count -gt 0 -and $sendEmail -eq $true) {

Write-Host "Sending report"
$mail_body = $mail_Body.Replace("image/advisory.png","cid:advisory")
$mail_body = $mail_Body.Replace("image/incident.png","cid:incident")
$mail_body = $mail_Body.Replace("image/healthy.png","cid:healthy")
$mail_body = $mail_Body.Replace("""","\""")
$mail_body = '"'+$mail_body+'"'

try {
#MS Graph API Starts Here
$body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$ClientID;client_secret=$ClientSecret}
$oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$tenantdomain/oauth2/v2.0/token -Body $body
$headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
 
$ToAddressJSON = $toAddress | ForEach-Object{'{"EmailAddress": {"Address": "'+$_+'"}},'}
$ToAddressJSON = ([string]$ToAddressJSON).Substring(0, ([string]$ToAddressJSON).Length - 1)
#$ToAddressJSON

$mailSubject = '"' + $mailSubject + '"'
#$newBody = '"body" : {	"contentType": "HTML",	"content": "'+$mail_body+'"	},'
$uri = "https://graph.microsoft.com/v1.0/users/$($fromAddress)/sendmail"
$mailbody = 
@"
{
"message" : {
	"subject": $mailSubject,
	"body" : {
		"contentType": "HTML",
		"content": $mail_Body
		},
  "toRecipients": [
	$ToAddressJSON
   ],
   "attachments":[
	   {
		"@odata.type":"#microsoft.graph.fileAttachment",
		"contentID":"advisory",
		"name":"advisory",
		"IsInline":true,
		"contentType":"image/png",
		"contentBytes":$base64_advisory
	   },
	   {
		"@odata.type":"#microsoft.graph.fileAttachment",
		"contentID":"incident",
		"name":"incident",
		"IsInline":true,
		"contentType":"image/png",
		"contentBytes":$base64_incident
	   },
	   {
		"@odata.type":"#microsoft.graph.fileAttachment",
		"contentID":"healthy",
		"name":"healthy",
		"IsInline":true,
		"contentType":"image/png",
		"contentBytes":$base64_healthy
	   }	   
   ]
}
}
"@
Invoke-RestMethod -Method Post -Uri $uri -Body $mailbody -Headers $headerParams -ContentType application/json
Write-Host "Report Sent!"
#MS Graph API Ends Here
}
catch {
	Write-Host "Failed to send report!"
	$_.Exception | Format-List
}
}
}
if ( $config.options.testMode -eq $false){
	Write-Host "Old Records File is not found. This is considered as first run. No report is generated or sent."
}

Rename-Item $newCSV $oldCSV
Stop-Transcript