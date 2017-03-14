#Add Exchange 2010/2013 snapin if not already loaded in the PowerShell session
if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
{
	try
	{
		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
	}
	catch
	{
		#Snapin was not loaded
		Write-Warning $_.Exception.Message
		EXIT
	}
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
}
$PSEmailServer = "xxx.xxx.xxx.xx"

$Report = @()
Write-Host "Searching mailserver for audit logs." -foregroundcolor green

$mbxs = Get-Mailbox
foreach ($mbx in $mbxs) {
	$identity = $mbx.Identity
	$auditlogentries = Search-MailboxAuditLog -Identity $identity -LogonTypes Delegate, Admin -StartDate (Get-Date).AddHours(-168) -ShowDetails
	if ($($auditlogentries.Count) -gt 0)
	{
		Write-Host "Found Admin/Delegate logons in " -nonewline 
		Write-Host $mbx -foregroundcolor red -NoNewLine 
		Write-Host "!"
		$auditlog += $auditlogentries


	foreach ($entry in $auditlogentries)        	
	{
		$reportObj = New-Object -TypeName System.Object    	
	        $reportObj | Add-Member NoteProperty -Name "Mailbox" -Value $entry.MailboxResolvedOwnerName
	        $reportObj | Add-Member NoteProperty -Name "Mailbox UPN" -Value $entry.MailboxOwnerUPN
	        $reportObj | Add-Member NoteProperty -Name "Timestamp" -Value $entry.LastAccessed
	        $reportObj | Add-Member NoteProperty -Name "Accessed By" -Value $entry.LogonUserDisplayName
	        $reportObj | Add-Member NoteProperty -Name "Operation" -Value $entry.Operation
	        $reportObj | Add-Member NoteProperty -Name "Result" -Value $entry.OperationResult
	        $reportObj | Add-Member NoteProperty -Name "Folder" -Value $entry.FolderPathName
		$reportObj | Add-Member NoteProperty -Name "Client IP" -Value $entry.ClientIPAddress
       

	        $report += $reportObj
	    }



    	$htmlbody = $report | ConvertTo-Html -Fragment

	$htmlhead="<html>
				<style>
				BODY{font-family: Arial; font-size: 8pt;}
				H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
				TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
				TD{border: 1px solid #969595; padding: 5px; }
				td.pass{background: #B7EB83;}
				td.warn{background: #FFF275;}
				td.fail{background: #FF2626; color: #ffffff;}
				td.info{background: #85D4FF;}
				</style>
				<body>
                <H3>Report of mailbox audit log entries</p><BR>"
		
	$htmltail = "</body></html>"	

	$htmlreport = $htmlhead + $htmlbody + $htmltail

    Write-Host "Writing report data to htmlfile"
    $htmlreport | Out-File "c:\mbx.html" -Encoding UTF8
}
}    
$auditlog | Export-CSV "c:\mbx.csv" -NoTypeInformation -Encoding UTF8
if ($report.count -gt 0)
    {
        Write-Host "Sending email"
	Send-MailMessage -to "alexander.kitaev@avestragroup.com" -from "exchange@avestragroup.com" -subject "Non-owner mailbox access report" -Body $htmlreport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -priority High -Attachments "c:\mbx.csv"
    }
else
{
    Write-Host "Nothing to send"
    Send-MailMessage -to "alexander.kitaev@avestragroup.com" -from "exchange@avestragroup.com" -subject "No non-owner mailbox access detected" -Body $htmlreport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
}
