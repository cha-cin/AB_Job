param
(
    [String]$JobListPath = "E:\MTAPPS\IS_Backend\AB_job\JobList.xlsx",
    [Array]$JobList = (Import-Excel -Path $JobListPath),
    [String]$jsname = "AD://ABTAGEN2",
    [String]$ABJobPath = "ABJobPath",
    #[String]$EmailRecipient = "lichiasin@micron.com,donliang@micron.com,IT_MFG_BE_OPS_MTB@micron.com",
	[String]$EmailRecipient = "lichiasin@micron.com",
    [System.MarshalByRefObject]$JobScheduler,
    [system.Data.DataTable]$InstanceTable = (New-Object system.Data.DataTable) #Main table
)



function Connect-JobScheduler
{
    param
    (
        #[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]
        [String]
        $JobSchedulerName
    )

    $JS = New-Object -ComObject "ActiveBatch.Abatjobscheduler"
    $JS.Connect($JobSchedulerName)
    Write-Host "`nConnect to $JobSchedulerName..." -ForegroundColor Cyan
    return $JS
}

#send mail
Function SendMail
{ 
    param(
        [string] $P_subject = "Please input the email subject",
        [string] $P_text = "Please input email body / test message",
        [string] $P_recipient = "",
        $file
      )
    #Additional message for this script
    #$P_text = $P_text + "<br>This job is triggered by ActiveBatch, in " + $ABJobPath
   
    #if($P_recipient -eq $null)
    #{
       # $P_recipient = ""
    #}

    #Orignal Send Mail Code
    $MailMessage = New-Object System.Net.Mail.MailMessage 
    $SMTPClient = New-Object System.Net.Mail.smtpClient 
    $SMTPClient.host = "RELAY.MICRON.COM"
    $MailMessage.IsBodyHtml=$true
    $MailMessage.Sender = "noreply@micron.com"
    $MailMessage.From = "noreply@micron.com"
    $MailMessage.Subject = $P_subject    
    $MailMessage.To.add($P_recipient)    
    #$MailMessage.CC.add($P_recipient)
    $MailMessage.Attachments.add($file)
	$MailMessage.IsBodyHTML = $true
	$MailMessage.Body = $P_text
    $SMTPClient.Send($MailMessage)    

}

#Create Instance Counter Table Head
[void]$InstanceTable.Columns.Add("ID")
[void]$InstanceTable.Columns.Add("Name")
[void]$InstanceTable.Columns.Add("FullPath")
[void]$InstanceTable.Columns.Add("LastRun")
[void]$InstanceTable.Columns.Add("NextRun")
[void]$InstanceTable.Columns.Add("Owner")
[void]$InstanceTable.Columns.Add("LastAudit")
[void]$InstanceTable.Columns.Add("SccheduledRunTime")
[void]$InstanceTable.Columns.Add("OverSccheduledRunTime")


$JobScheduler = Connect-JobScheduler -JobSchedulerName $jsname #Connecting to Job Scheduler
Write-Host("Job Scheduler Name : " + $jsname.Name + "`n")
$sentAlert = 0
#Loop On every Object

$PageContext = New-Object System.Collections.Generic.List[System.Object]
foreach ($job in $JobList)
{       
    
    $Object = $JobScheduler.GetAbatObject($job.JobPath)
	Write-Host($object.Name,$job.limitation)
    
    if ($object.enabled -eq "TRUE")
    {
        $TimeSpan = (New-TimeSpan –Start $object.LastInstanceExecutionDateTime –End $(GET-DATE)).TotalMinutes  
		Write-Host($TimeSpan)
        if ($TimeSpan -ge $job.limitation){
            $InstanceTableRow = $InstanceTable.NewRow()
            $InstanceTableRow.ID = $object.ID
            $InstanceTableRow.Name = $object.Name
            $InstanceTableRow.FullPath = $object.FullPath
            $InstanceTableRow.LastRun = $object.LastInstanceExecutionDateTime
            $InstanceTableRow.NextRun = $object.NextScheduledExecutionDateTime
            $InstanceTableRow.Owner = $object.Owner
            $InstanceTableRow.LastAudit = ($object.GetAuditsEx() | Select-Object -Property DateTime -First 1).DateTime
			$InstanceTableRow.SccheduledRunTime = $job.limitation
			$InstanceTableRow.OverSccheduledRunTime = $TimeSpan
            $InstanceTable.Rows.Add($InstanceTableRow)
            $sentAlert = 1
			
			# mail context
			$PageContext.add("<br />")
			$PageContext.add("`n")
			$PageContext.add($($object.Name))
			
			$PageContext.add($($object.FullPath))
			
			#$PageContext.ToArray()
			
        }
		
    }

}       

Write-Host ($PageContext)
Write-Host "`Exporting to CSV..." -ForegroundColor Yellow
#Export it to CSV
$InstanceTable | Export-csv -path "E:\MTAPPS\IS_Backend\AB_job\OverOnedayNotExecute.csv" -NoTypeInformation -Encoding UTF8


$PageContext2 =  $PageContext | Out-String
Write-Host ($PageContext2)
#Sending Page calls
if($sentAlert -eq 1)
{
	#Write-Host($PageContext)
    SendMail -P_subject "[ActiveBatch] Jobs have hang!" -P_text $PageContext2 -P_recipient $EmailRecipient -file "E:\MTAPPS\IS_Backend\AB_job\OverOnedayNotExecute.csv"
}



Write-Host "`Job is finished" -ForegroundColor Green