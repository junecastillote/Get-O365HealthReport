<#	
	.NOTES
	===========================================================================
	 Created on:   	7-August-2018
	 Created by:   	June Castillote
					june.castillote@gmail.com
	 Filename:     	Get-o365HealthReport.ps1
	 Version:		1.3 (25-February-2019)
	===========================================================================

	.LINK
		https://www.lazyexchangeadmin.com/2018/10/shd365.html
		https://github.com/junecastillote/Get-O365HealthReport

	.SYNOPSIS
		This script utilize the Office 365 Management API v2 to retrieve the service health status
		and the Microsoft Graph API to send the report thru email using an Office 365 Mailbox.

	.DESCRIPTION
		For more details and usage instruction, please visit the link:
		https://www.lazyexchangeadmin.com/2018/10/shd365.html
		https://github.com/junecastillote/Get-O365HealthReport		
				
	.EXAMPLE
		.\Get-o365HealthReport.ps1
#>

<# CHANGE LOGS:

v1.0
- Initial Build

v1.1
- Added “organizationName” field in config.xml
- Removed “mailSubject” field from config.xml
- Send one email per event (alerts are no longer consolidated in one single email)

v1.2
- Modified to also check the changes in "Status" to trigger an update alert. (eg. Service Degradation to Service Restored). 
This is because I observed that some events' Last Updated Time does not change but the Status change which is not getting
captured by the previous script.

v1.3
- exclusions.csv file inside the \resource folder can not be used to exclude workloads from the report.
- the csv file lists current workloads available (eg. exchange online, Sharepoint Onine..)
- to exclude the specific workloads from the report, just change the value under the Exclude column.

Example:

		WorkLoad,Exclude <---- this are the column names, do not change
		Exchange Online,0
		Microsoft Intune,1
		...

- The above example excludes Microsoft Intune from the report, and will only report on Exchange Online events.

#>

#Requires -Version 4.0
$scriptVersion = "1.3"

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

$workLoadStringArrays = @(
	"WorkLoad,Exclude"
	"Azure Information Protection,0",
	"Dynamics 365,0",
	"Dynamics 365 for Operations,0",
	"Exchange Online,0",
	"Identity Service,0",
	"Microsoft Intune,0",
	"Microsoft StaffHub,0",
	"Microsoft Teams,0",
	"Mobile Device Management for Office 365,0",
	"Office 365 Portal,0",
	"Office Online,0",
	"Office Subscription,0",
	"OneDrive for Business,0",
	"Planner,0",
	"Power BI,0",
	"SharePoint Online,0",
	"Skype for Business,0",
	"Social Engagement,0",
	"Sway,0",
	"Yammer Enterprise,0"	
)

#import config.xml
[xml]$config = Get-Content "$($script_root)\resource\config.xml"

#import exclusions
if (Test-Path "$($script_root)\resource\exclusions.csv") {
	$exclusions = Import-Csv "$($script_root)\resource\exclusions.csv" | Where-Object {$_.Exclude -eq 1}
}
else {
	#if the exclusions.csv file does not exist, create the file with default values
	$workLoadStringArrays | ConvertFrom-Csv | Export-Csv "$($script_root)\resource\exclusions.csv" -NoTypeInformation
	$exclusions = Import-Csv "$($script_root)\resource\exclusions.csv" | Where-Object {$_.Exclude -eq 1}
}

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
	
	#Why no filter? Because at the time of writing this code, filters are not working for the v2 API.
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

if ($exclusions.Workload -notcontains $message.WorkloadDisplayName) {
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
}
#process the retrieved records
$newResult | Sort-Object LastUpdatedTime -Descending | export-Csv -notypeInformation $newCSV

#this import is to makes sure that the comparison between new and old records are as accurate as possible
#if not imported, the comparison fails and all retrieved records are treated as new
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
			#v1.2 update 
			# - Add Status changes to the comparison
			if ($oldRecord.LastUpdatedTime -ne $newRecord.LastUpdatedTime -OR $oldRecord.Status -ne $newRecord.Status) {
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
				
		foreach ($record in $updatedRecord)	{
			
			$mail_Body1 = @()
			$mail_Body2 = @()
			#Write-Host "Writing Report"
			$mailSubject = '[' + $record.Status + '] ' + $record.ID + ' | ' + $record.WorkloadDisplayName +' | ' + $record.Title
			if ( $config.options.testMode -eq $true){
				$mailSubject = "[TEST MODE] | " + $mailSubject
			}
			$mail_Body1 = "<html><head><title>$($mailSubject)</title>"
			$mail_Body1 = $mail_Body1 -join "`n" #convert to multiline string
			$mail_Body2 += "</head><body>"
			$mail_Body2 += "<hr>"			
			$mail_Body2 += '<table id="section"><tr><th width="95%">' + $record.ID + ' | ' + $record.WorkloadDisplayName +' | ' + $record.Title + '</th></tr></table>'
			$mail_Body2 += "<hr>"				
			$mail_Body2 += '<table id="data">'

			if ($record.Status -eq 'Service Restored'){
				$mail_Body2 += '<tr><th>Status</th><td class="good">'+$record.Status+'</td></tr>'
			}
			else {
				$mail_Body2 += '<tr><th>Status</th><td class="bad">'+$record.Status+'</td></tr>'
			}
			$mail_Body2 += '<tr><th>Organization</th><td>'+$config.options.organizationName+'</td></tr>'
			$mail_Body2 += '<tr><th>Classification</th><td>'+$record.Classification+'</td></tr>'
			$mail_Body2 += '<tr><th>Event Type</th><td>'+$record.EventType+'</td></tr>'
			$mail_Body2 += '<tr><th>User Impact</th><td>'+ $record.ImpactDescription+'</td></tr>'
			$mail_Body2 += '<tr><th>Last Updated</th><td>'+ $record.LastUpdatedTime +'</td></tr>'
			$mail_Body2 += '<tr><th>Start Time</th><td>'+ $record.StartTime +'</td></tr>'
			$mail_Body2 += '<tr><th>End Time</th><td>'+ $record.EndTime+'</td></tr>'
			$mail_Body2 += '<tr><th>Latest Message</th><td>'+($record.Message).Replace("`n","<br />")+'</td></tr>'
			$mail_Body2 += '</table>'

			$mail_Body2 += '<p><table id="section">'
			$mail_Body2 += '<tr><th><center>----END of REPORT----</center></th></tr></table></p>'
			$mail_Body2 += '<p><font size="2" face="Tahoma"><br />'
			$mail_Body2 += '<br />'
			$mail_Body2 += '<p><a href="https://github.com/junecastillote/Get-O365HealthReport">Get-O365HealthReport v.'+ $scriptVersion +'</a></p>'
			$mail_body2 += '</body>'
			$mail_body2 += '</html>'
			$mail_Body2 = $mail_Body2 -join "`n" #convert to multiline string
			#combine body texts
			$mail_Body = $mail_Body1 + $css_string + $mail_Body2
			$mail_body | Out-File "$($script_root)\output\$($record.ID).html"

#send email if new or updated events are found
if ($sendEmail -eq $true) {

Write-Host "Sending Alert for $($record.id)"
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
#Write-Host "Report Sent!"
#MS Graph API Ends Here
}
catch {
	Write-Host "Failed to send report!"
	$_.Exception | Format-List
}
}
}	
			
		}
		
		}
		
		else {
	if ( $config.options.testMode -eq $false){
		Write-Host "Old Records File is not found. This is considered as first run. No report is generated or sent."
	}
}


Rename-Item $newCSV $oldCSV
Stop-Transcript