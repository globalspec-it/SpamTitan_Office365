We switched to using SpamTitan for our email gateway to allow us to gain much more control over our SPAM filtering and reporting metrics.
Office 365 is a great platform for email, but their anti-SPAM leaves much to be desired, with almost no configuration.  Reporting is 
another very weak point to Office 365, and to some extent, on-prem Exchane as well.  SpamTitan solves both issues.

In SpamTitan, we want to only allow email into the system if there is a valid email address in Office 365.  Unfortunately, Office 365
doesn't allow dynamic recipient verification.  The solution to this is PowerShell.  We pull a report of all email addresses and aliases 
from Office 365 and export to CSV.  Then we will read the file and using the SpamTitan API, import the mail addresses to the appropriate
domains in SpamTitan.

This script accounts for multiple domains in both locations and saves a log file with import metrics and a detailed log of any failures. 
It will then email the log to your destination of choice.

Configuration:

The top section of the script contains a bunch of variables used through.  This section needs to be customized for the location of your 
exports and logs, tenant user name, SpamTitan server info and email log information.

Pre-Requisites:
So that your account information isn't in plain text, you need to setup a credential file that will be used by the script.
You need to do this on initial configuration and every time your password changes.

You'll need an account with admin privileges to Office 365 to be able to read all email addresses. Run the following in PowerShell

Read-Host -Prompt "Mypassword" -AsSecureString | ConvertFrom-SecureString | Out-File c:\O365_SpamTitan\tenantpassword.key

Make sure you set the $CredentialFile variable in the script tothe location of this tenantpassword.key file you just created.

If you run the script from the command line, it will log to the console as it goes.
