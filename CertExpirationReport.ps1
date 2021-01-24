<#
.SYNOPSIS

    Interogates Windows Certificate Athority for certificate expiration info.

.DESCRIPTION

    PowerShell script for retrieving properties (expiration in this case) from a Microsoft CA. Used for alerting
    and reporting of upcoming certificate expiration events.

.PARAMETER ExportWindow

    Designates window of time to look for certificates whose expiration is approaching.

.PARAMETER Templates

    Defined as an array the templates whose certs should be evaluated.

.PARAMETER TempPath

    Designates a working location for script data. Cleaned up at script closure.

.PARAMETER MailRecipient
   
    If email is sent, designates reciever of email.

.PARAMETER MailServer

    Mail server FQDN.

.NOTES

    Author: Tim Sullivan
    Version: 0.2
    Date: 08/12/2017
    Name: CertExpirationReport.ps1

    CHANGE LOG
    0.2: Initial Release

.EXAMPLE
    
    CertExpirationReport.ps1 -ExpireWindow 30 -Templates Template1,Template2 -MailRecipient foo@mail.com -MailSender error@mail.com -MailSubject Report -Email $true

#>
Param
(
    $ExpireWindow=30,
    $Templates=("ContosoWebServer","ContosoWorkstation"),
    $TempPath="C:\Working",
    $MailRecipient="tim@Contoso.com",
    $MailServer="mail.Contoso.com",
    $MailSender="media@Contoso.com",
    $MailSubject="Expiring Cert Report",
    $Email=$true,
    $EVLog=$true
)


#Get current date
$Today = Get-Date

#Construct future date (X days from now)
$Before = $Today.AddDays($ExpireWindow)
$Before = "$($Before.Month)/$($Before.Day)/$($Before.Year)"
#Write-Output "Before date: $Before"

#Construct Current Date
$After = "$($Today.Month)/$($Today.Day)/$($Today.Year)"
#Write-Output "After date: $After"

#Gather certificate info.
Foreach ($Item in $Templates)
{

    #Get OID from template name
    $OID = Get-CATemplate | Select-Object Name,Oid | Where-Object {$_.Name -like "$Item"}
    $OID = $OID.Oid
    #Write-Output "OID for template $item is"$OID

    #Create certutil restriction statement
    $Restrict = "NotAfter<=$Before,NotAfter>=$After,CertificateTemplate = $OID,Disposition=20"
    #Write-Output "Restrict string: $Restrict"

    #Certutil execution. Exports data out to temporary CSV.
    certutil -restrict $Restrict -view csv > $TempPath\$Item.TempCSV.csv #-out "RequesterName,CommonName,Certificate Expiration Date"

    #Returned Data brought back into script.
    $CertData = Import-Csv $TempPath\$Item.TempCSV.csv | Select-Object -Property 'Issued Common Name','Certificate Expiration Date','Requester Name','Certificate Template'

    Remove-Item $TempPath\$Item.TempCSV.csv

    #$CertData

}

#Report construction.
$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
<title>
Upcoming Expiring PKI Certificates
</title>
"@

$MailBody = $CertData | ConvertTo-Html -Head $Header -PreContent "PKI Certificates Set to Expire Within $ExpireWindow Days" | Out-String

$LogMsg = 
@"
Certificate expiration report data
Expiration Window: $ExpireWindow
Tempales Evaluated: $Templates

Results
$CertData

End of run
"@

#Write to local application log
If ($EVLog -eq $true){
    #Write event to local event log.
    #Create event log source, if it does not already exist.
    if ([System.Diagnostics.EventLog]::SourceExists("CertReport") -eq $false) 
    {
        [System.Diagnostics.EventLog]::CreateEventSource("CertReport","Application")
    }
    Write-EventLog -LogName "Application" -EntryType Information -EventId 530 -Source CertReport -Message $LogMsg

}

#Send out email report.            
If ($email -eq $true)
{
    try
    {
        Send-MailMessage -To $MailRecipient -Subject $MailSubject -Body $MailBody.ToString() -SmtpServer $MailServer -From $MailSender -BodyAsHtml
    }
    catch
    {
        #Write-Output "Error sending email" $error[0]
    }
}
