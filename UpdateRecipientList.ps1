# Import the O365 library 
Import-Module MSOnline

#=============================
# Edit Variables as Appropriate
#==============================

$TenantUname = "user@domain.com" 																		    # Office 365 User to run the export under
$LogFileName = "C:\O365_SpamTitan\SpamTitan_RecipientListUpdate_Log.txt"									# Log File Location for Import Details
$detailLogFileName = "C:\O365_SpamTitan\SpamTitan_Detail_log.txt"											# Log File containing the details of failed imports
$EmailTempfile = "C:\O365_SpamTitan\O365Emailproxy.csv"														# Declare an intermediate file where you will store the proxy address data
$O365EmailAddresses = "C:\O365_SpamTItan\O365EmailAddresses.csv"											# Declare a file to store the emails for import to SpamTitan
$CredentialFile = "C:\O365_SpamTitan\tenantpassword.key"													# Location of Encrypted Credential File
$SpamTitanServer = "http://spamtitan.domain.com"															# URL to SpamTitan Server
$EmailFrom = "SpamTitanImport@domain.com"																		# Email From Address
$EmailTo = "user@domain.com"																					# Email To Address
$EmailSubject = "SpamTitan Office365 Email Import Failed"														# Email Subject
$EmailBody = "The SpamTitan email address import from Office 365 has failed.  Please see log files attached"    # Email Body
$SMTPServer = "smtpserver.domain.com" 																			# SMTP Server used to send email



#============================
# Do Not Edit Below This Line
#============================
# Office 365 Export Setup
#============================

#Reference where you stored the Tenant password
$TenantPass = cat $CredentialFile | ConvertTo-SecureString
$TenantCredentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $TenantUname, $TenantPass

# Open a session to Exchange on O365 (While doing testing, comment out the next 3 lines after first run to make testing faster
# since the session will still be opened.

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $TenantCredentials -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber
Connect-MsolService -Credential $TenantCredentials
 
#Declare a file where you will store the proxy address data
$fileObjectEmail = New-Item $EmailTempFile -type file -force
 
#Declare a file to store the email info
$fileObject = New-Item $O365EmailAddresses -type file -force

#Set up Logging
$logFile = New-Item $LogFileName -type file -force

#Path to Log File containing import failure details
$detailLogFile = New-Item $detailLogFileName -type file -force

#Global Failure Indicator for email notification
$totalFailures = 0

#Write Date Stamp to log file
date | Out-File $detailLogFile -encoding utf8 -Append

#-----------------------------
# Begin Export from Office 365
#-----------------------------
 
#Write the header
$evt_string = "Domain,EmailAddresses"
$evt_string | Out-file $fileObject -encoding ascii -Append
 
#Get the proxy info
Get-Recipient -Resultsize unlimited -RecipientType DynamicDistributionGroup,UserMailbox,MailUser,MailUniversalDistributionGroup,MailUniversalSecurityGroup,MailNonUniversalGroup,PublicFolder | select EmailAddresses | Export-csv -notypeinformation $fileObjectEmail -append -encoding utf8
 
#Read output into a variable
$file = Get-Content $EmailTempfile
for($i=1;$i -lt $file.count;$i++)
{
    $evt_string=""

    #Split the proxy data into individual addresses
    $csvobj = ($file[$i] -split ",")
    $EmailAddr = $csvobj[0]

    $GetEmail = $EmailAddr -replace '"', '' -split(' ')

    #write out the display name and email address (One person can have several), filter for smtp only and exclude onmicrosoft.com addresses and external contacts
    for($k=0;$k -lt $GetEmail.count;$k++)
    {
        If (($GetEmail[$k] -match "smtp:") -and ($GetEmail[$k] -notmatch "onmicrosoft.com") -and ($GetEmail[$k] -notmatch "#EXT#"))
        {
                        $domain_string = $GetEmail[$k].split("@")[1]
                        $evt_string = $GetEmail[$k].split(":")[1]
                        $evt_string = $domain_string+","+$evt_string
                        $evt_string | Out-file $fileObject -encoding utf8 -Append
        }
    }
}


#----------------------------
# Begin Import to SpamTitan
#----------------------------
#Read Office365 email addresses from file and set up fields for import
$addresses = Import-CSV $O365EmailAddresses

#Read list of email domains configured in SpamTitan
[xml]$SpamTitanDomains = Invoke-WebRequest "$SpamTitanServer/api/domain.php?method=list"

#Iterate through SpamTitan domains
foreach ($domain in $SpamTitanDomains.response.domains.domain)
{
    #Initialize Counters for current domain
    $found = 0    #Number of addresses already in ST
    $missing = 0  #Number of addresses missing from ST
    $fail = 0     #Number of additions to ST that failed
    $success = 0  #Number of additions to ST that succeeded
    $unknown = 0  #Number of addition to ST with an unknown result code

    #Read list of email addresses already configured in SpamTitan for current domain
    [xml]$STAllowedUsers = Invoke-WebRequest "$SpamTitanServer/domain/list_rv?name=$domain"
    Write-Host -ForegroundColor yellow -BackgroundColor black "Currently in $domain"
    #Write log of current domain
    $domain | Out-file $LogFileName -encoding utf8 -Append
     
    foreach ($address in $addresses) #Iterate through Office 365 addresses
    {  
        if ($address.domain -eq $domain) #Check if domain of current address matches
        {
            if ($stallowedusers.response.allowed.auth -contains $address.emailaddresses) #If current address is already in SpamTitan
            {
                Write-Host -foreground Green Found: $address.emailaddresses
                $found++
            }
            else
            {
                Write-Host -ForegroundColor black -BackgroundColor red Missing: $address.emailaddresses
                $missing++
                $addressToAdd = $address.EmailAddresses.ToLower()
                #Add missing email address to SpamTitan using API
                $response = Invoke-WebRequest "$SpamTitanServer/domain/edit?name=$domain&email=$addressToAdd"
                $responseValue = $response.response.stat.ToString()
                If ($responseValue.contains("fail"))
                    {
                        $fail = $fail + 1
                        $totalFailures = $totalFailures + 1
                        $output = "Failed: " + $Domain + " " + $addressToAdd + " " + $response.Response.error.ToString()
                        $output | Out-File $detailLogFile -encoding utf8 -Append
                        #$responseValue | Out-File $detailLogFile -encoding utf8 -Append
                    }
                ElseIf ($responseValue.contains("ok"))
                    {
                        $success = $success + 1
                        $output = "Success: " + $domain + " " + $addressToAdd
                        $output | Out-File $detailLogFile -encoding utf8 -Append
                        #$responseValue | Out-File $detailLogFile -encoding utf8 -Append
                    }
                Else
                    {
                        $unknown = $unknown + 1
                        $totalFailures = $totalFailures + 1
                        $output = "Unknown: " + $domain + " " + $addressToAdd + " " + $response.Response.error.ToString()
                        $output | Out-File $detailLogFile -encoding utf8 -Append
                        #$responseValue | Out-File $detailLogFile -encoding utf8 -Append
                    }
            }
        }
    }

    $foundResults = "$found email addresses already in SpamTitan"
    $missingResults = "$missing email addresses missing in SpamTitan"
    $foundResults | Out-File $LogFileName -encoding utf8 -Append
    $missingResults | Out-File $LogFileName -encoding utf8 -Append
    $successes = $success.ToString() + " Add Successes"
    $failures = $fail.ToString() + " Add Failures (see detail log)"
    $unknowns = $unknown.ToString() + " Add Unknowns (see detail log)"

    $successes | Out-file $LogFileName -encoding utf8 -Append
    $failures | Out-file $LogFileName -encoding utf8 -Append
    $unknowns | Out-file $LogFileName -encoding utf8 -Append 
    " " | Out-File $LogFileName -encoding utf8 -Append
}

$waitMail = 60
If ($totalFailures -gt 0)
{
    Write-Host Sending Failure Email
    try {   
        Send-MailMessage -From $EmailFrom -Subject $EmailSubject -To $EmailTo -Attachments $logFileName, $detailLogFileName -Body $EmailBody -SMTPServer $SMTPServer -EA Stop
        Exit
    }
    catch { 
        Write-Host $sendErr
        Sleep $waitMail
        Send-MailMessage @MessageParameters
        Exit
    }
    finaly {  
        Write-Host "Error: Unable to send email."
        Exit
    }
}