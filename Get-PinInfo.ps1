<#

.SYNOPSIS
This is a script to send mail to the user who Lync dialin PIN will expire

.DESCRIPTION
Original Script (v1.0) created by Petre Calinoiu (PCA)
Changelog:
v1.0.0 - 13.04.2016 - Original Script (PCA)
v1.1.0 - 16.04.2016 - Reduced console output
                      created summary
                      added progress bars
                      added event log switch and entries

.EXAMPLE
.\Get-PinInfo.ps1 -CSGroup Company_Skype_for_Business_Users
Using CSGroup parameter to provide the active directory group with the lync users

.EXAMPLE
.\Get-PinInfo.ps1 -CSGroup Company_Skype_for_Business_Users -ToEvents
Using the ToEvents switch, the output summary will be written in Event Logs

.PARAMETER CSGroup
The name of an Active Directory group which includes all Lync users

.PARAMETER ToEvents
Switch activates the output to the events log

.NOTES
You need to run this script as a member of the CSAdministrators group; doing so is the only way to ensure you have permission to query data.
The script must run with elevated privilege and can only be run in one Lync 2013 Front-End Server

#>

Param(
   [Parameter(Mandatory=$true, HelpMessage="Enter an AD Group:", ValueFromPipeline = $true)]
   [string]$CSGroup,
   [parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)]
   [switch]$ToEvents
)

Import-Module ActiveDirectory
Import-Module Lync


#global variables need to be set
#mail settings
$from = ""
$smtp = ""
$mailsubject = "Skype for Business Pin Notification"
#get dialin URL from topology
$SimpleURLEntries = Get-CsSimpleUrlConfiguration -Identity Global | select SimpleUrl
$dialinUrl = $SimpleURLEntries.SimpleUrl[0].ActiveUrl
#counter variable for progress bar
$counter = 0
$countMails = 0
$countCompliantPin = 0
$countNoPin = 0
$countExpiredPin = 0
#object for saving users that have received mails
[System.Collections.ArrayList]$mailedUsers = @()


#Create a interminate progress bar
$i = 0
#Asking AD for group members in a background job
$job =  Start-Job -ScriptBlock { Get-ADGroup $args[0] -Properties members | select -ExpandProperty members | Get-ADUser -Properties samaccountname, givenname, mail | select samaccountname, givenname, mail } -ArgumentList $CSGroup
Write-Progress -Id 1001 -Activity “Searching Lync Users...” -status “Please wait” -PercentComplete 0
while(Get-Job -State "Running")
{
    if($i -eq 100)
    {$i=0}
    Write-Progress -Id 1001 -Activity “Searching Lync Users...” -status “Please wait” -PercentComplete $i
    Start-Sleep(1)
    $i = $i + 5
}

$Members = Receive-Job -Job $job -AutoRemoveJob -Wait

#read the total number of users available in AD Group
$membersCount = $Members.Count
Write-Progress -Id 1001 -Activity “Searching Lync Users...” -status “$membersCount users found. Process completed” -PercentComplete 100
Start-Sleep(1)
Write-Progress -Id 1001 -Activity “Searching Lync Users...” -Completed

#Processing every user found
ForEach ($user in $Members)
{
    $samaccountname = $user.samaccountname
    $mailaddress = $user.mail
    $username = $user.givenname
    
    #ingreasing the progress counter
    $counter++

    $enabled = Get-CsUser -filter {SamAccountName -eq $SamAccountName}
 
    # Check if user is enabled for Lync 2013 pool
    if ($enabled.RegistrarPool -ne $null)
    {
        #ask users pin expiration date
        $usersPin = Get-CsClientPinInfo -Identity $samaccountname | Select-Object PinExpirationTime, IsPinSet
        $userExpirationTime = $usersPin.PinExpirationTime
        
        #ask current date
        $currentdate = Get-Date
        
        #calculate how many days are available till pin expires
        $diff = $userExpirationTime - $currentdate
        $daysDiff = $diff.Days

        #check if the user has a Pin already set
        if($usersPin.IsPinSet)
        {
            #when are 14 or 7 days more till pin expire, send a mail to the user
            if(($daysDiff -eq 14) -or ($daysDiff -eq 7))
	            {
                    #send a mail to the user to announce him about pin expiration period
                    $body="Hello $username,<br><br>"
                    $body+="Your Skype for Business PIN will expire in $daysDiff days<br>"
                    $body+="Please change it at $dialinUrl<br><br>"
                    $body+="Please be aware, if you do not change your PIN, you are not allowed to join the meeting via phone."
                    Send-MailMessage -From $from -SmtpServer $smtp -Subject $mailsubject -To $mailaddress -Port 25 -BodyAsHtml -Body $body
                    $countMails++
                    #save user samaccountname and mail address for summary
                    $properties = @{User=$samaccountname; Mail = $mailaddress}
                    $objectTemplate = New-Object -TypeName PSObject -Property $properties
                    $mailedUsers.Add($objectTemplate)
                }
            else
            {
                #check if Pin has expired today
                if(($daysDiff -eq 0) -and ($diff.TotalMilliseconds -lt 0))
	                {
                        #send a mail to the user to announce him about pin expiration
                        $body="Hello $username,<br><br>"
                        $body+="Your Skype for Business PIN is expired<br>"
                        $body+="Please assign a new one at $dialinUrl<br><br>"
                        $body+="Please be aware, if you do not change your PIN, you are not allowed to join the meeting via phone."
                        Send-MailMessage -From $from -SmtpServer $smtp -Subject $mailsubject -To $mailaddress -Port 25 -BodyAsHtml -Body $body
                        Write-Host "mail sent to $mailaddress"
                        $countMails++
                        #save user samaccountname and mail address for summary
                        $properties = @{User=$samaccountname; Mail = $mailaddress}
                        $objectTemplate = New-Object -TypeName PSObject -Property $properties
                        $mailedUsers.Add($objectTemplate)
                    }
                else
                    {
                        #check if pin is expired since more than 1 day
                        if($daysDiff -lt 0)
                            {
                                #count for summary
                                $countExpiredPin++
                            }
                        else                            
                            {
                                #in this case nothing can be done because the pin is compliant
                                $countCompliantPin++
                            }
                    }
            }
        }
        else
        {
            #print a message to the console if the pin is not set for this user
            $countNoPin++
        }
    }
    Write-Progress -Id 1000 -Activity “Processing Lync Users...” -status “Lync Users already completed: $counter from $membersCount” -percentComplete ($counter / $Members.Count*100)
}
Write-Progress -Id 1000 -Activity “Processing Lync Users...” -Completed

#creating summary
$logtext =  @"
---------------------------------------
Total Users                        $counter
Mails sent                         $countmails
Compliant PIN Users                $countCompliantPin
Expired PIN Users                  $countExpiredPin
Users without a PIN set            $countNoPin
---------------------------------------
`n
"@
if ($mailedUsers.Count -ne 0)
{
    $result = "Following Users received a mail notification`n"
    $result += $mailedUsers | ft -AutoSize | Out-String
    $result
}
else
{
    $result = "No notification was sent"
}
$logtext += $result

#check if output to event logs is requested and if the answer is positive write an entry otherwise put it on the screen
if($ToEvents)
{
    #check if Event Log Lync Scripts and PIN Info source exist
    if ((Get-EventLog -list | Where-Object {$_.logdisplayname -eq "Lync scripts"}) -and ([System.Diagnostics.EventLog]::SourceExists("PIN Info")))
    {
        #Write event entry
        Write-EventLog -LogName 'Lync Scripts' -Source 'PIN Info' -EntryType Information -EventId 1000 -Message $logtext
    }
    else
    {
        #if log or source not exist, create a new one and write event entry
        New-EventLog -LogName "Lync Scripts" -Source 'PIN Info'
        Write-EventLog -LogName 'Lync Scripts' -Source 'PIN Info' -EntryType Information -EventId 1000 -Message $logtext
    }
}
else
{
    Write-Host $logtext
}
