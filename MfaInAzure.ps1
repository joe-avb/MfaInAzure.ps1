# Get secure Automate Credentials and assign to variable
$credObject = Get-AutomationPSCredential -Name "Autobot"

# Set variables for mail delivery
$SMTPserver = "smtp.office365.com"
$SMTPport = "587"
$from = "autobot@avb.net"
$to = "it_reporting@avb.net"
$subject = "MFA Status Report"
$emailbody = "MFA Status Report"
$Username = $credObject.UserName 
$Password = $credObject.GetNetworkCredential().Password

# This function takes the returned query data, which has been assigned to a variable, and streams it from memory into an email attachment
Function ConvertTo-CSVEmailAttachment {
Param(
    [Parameter(Mandatory=$true)]
    [String]$FileName,
    [Parameter(Mandatory=$true)]
    [Object]$PSObject,
    $Delimiter
    )
    If ($Delimiter -eq $null){$Delimiter = ","}
    $MS = [System.IO.MemoryStream]::new()
    $SW = [System.IO.StreamWriter]::new($MS)
    $SW.Write([String]($PSObject | convertto-csv -NoTypeInformation -Delimiter $Delimiter | % {($_).replace('"','') + [System.Environment]::NewLine}))
    $SW.Flush()
    $MS.Seek(0,"Begin") | Out-Null
    $CT = [System.Net.Mime.ContentType]::new()
    $CT.MediaType = "text/csv"
    Return [System.Net.Mail.Attachment]::new($MS,$FileName,$CT)
}

# Connect to the MSOnline service using the Automate Credentials, then query, filter, and sort the MFA data and assign it to a variable
Connect-MsolService -Credential $credObject
$ADList = Get-MsolUser -all |
    Select-Object DisplayName,UserPrincipalName,isLicensed,@{N="MFA Status"; E={
        if($_.StrongAuthenticationRequirements.Count -ne 0){
            $_.StrongAuthenticationRequirements[0].State
        } else {
            'Disabled'}
        }
    } |
    Where-Object {$_.IsLicensed -ne $False -and `
                  $_.UserPrincipalName -notlike "*#EXT#*" -and `
                  $_."MFA Status" -ne "Enforced"} |
    Sort-Object -Property UserPrincipalName 

# Run the function to create the attachment and use .NET commands to send the email out using the Mail variables above
$EmailAttachment = ConvertTo-CSVEmailAttachment -FileName "MfaDisabledUserReport_$(Get-Date -f yyyy-MM-dd).csv" -PSObject $ADList
$mailer = new-object Net.Mail.SMTPclient($SMTPserver,$SMTPport)
$mailer.EnableSsl = $true
$mailer.Credentials = New-Object System.Net.NetworkCredential($Username,$Password)
$msg = new-object Net.Mail.MailMessage($from, $to, $subject, $emailbody)
$msg.Attachments.Add($EmailAttachment) #### This uses the attachment made using the function above. 
$msg.IsBodyHTML = $false
$mailer.send($msg)