<#

.SYNOPSIS
This is a script to send mail to the user who Lync dialin PIN will expire

.DESCRIPTION
Original Script (v1.0) created by Petre Calinoiu (PCA)
Changelog:
v1.0 - Original Script (PCA) - 13.04.2016

.EXAMPLE
.\Get-PinInfo.ps1 Company_Skype_for_Business_Users
Using CSGroup parameter without his name

.EXAMPLE
.\Get-PinInfo.ps1 -CSGroup Company_Skype_for_Business_Users
Using CSGroup parameter with his name

.PARAMETER CSGroup
The name of an Active Directory group which includes all Lync users

.NOTES
You need to run this script as a member of the CSAdministrators group; doing so is the only way to ensure you have permission to query data.
The script must run with elevated privilege and can only be run in one Lync 2013 Front-End Server

#>

Param(
   [Parameter(Mandatory=$true,HelpMessage="Enter an AD Group:",ValueFromPipeline=$True)]
   [string]$CSGroup
)

Import-Module ActiveDirectory
Import-Module Lync


#global variables need to be set
$from = ""
$smtp = ""
$mailsubject = "Skype for Business Pin Notification"
#get dialin URL from topology
$SimpleURLEntries = Get-CsSimpleUrlConfiguration -Identity Global | select SimpleUrl
$dialinUrl = $SimpleURLEntries.SimpleUrl[0].ActiveUrl
#counter variable for progress bar
$counter = 0


#Create a progress bar
Write-Host "Searching for Lync Users. Please wait!" -ForegroundColor Green
Write-Progress -Id 1001 -Activity “Processing Lync Users...” -status “Lync Users already completed: 20%” -percentComplete (20)
#Import the group member list
$Members = Get-ADGroup $CSGroup -Properties members | select -ExpandProperty members | Get-ADUser -Properties samaccountname, givenname, mail | select samaccountname, givenname, mail
#close progress bar
Write-Progress -Id 1001 -Activity “Processing Lync Users...” -status “Lync Users already completed: 100%” -percentComplete (100)
Start-Sleep(1)
Write-Progress -Id 1001 -Activity “Processing Lync Users...” -Completed

#read the total number of users available in AD Group
$membersCount = $Members.Count

ForEach ($user in $Members)
{
    $samaccountname = $user.samaccountname
    $mailaddress = $user.mail
    $username = $user.givenname
    
    #ingreasing the progress counter
    $counter++
    
    write-host "Processing:" $samaccountname -ForegroundColor Yellow

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

        #check if the user has a Pin set
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
                    Write-Host "mail sent to $mailaddress" -ForegroundColor Yellow
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
                        Write-Host "mail sent to $mailaddress" -ForegroundColor Yellow
                    }
                else
                    {
                        #check if pin is expired since more than 1 day
                        if($daysDiff -lt 0)
                            {
                                #convert datetime to short date and print the message on console
                                $userExpirationTime = $userExpirationTime.ToShortDateString()
                                Write-Host "The PIN is expired since $userExpirationTime" -ForegroundColor Red
                            }
                        else
                            #in this case nothing can be done because the pin is compliant
                            {Write-Host "The PIN don't need to be changed" -ForegroundColor Red}
                    }
            }
        }
        else
        {
            #print a message to the console if the pin is not set for this user
            Write-Host "You have no PIN set. Please set a PIN to use dial in capabilities" -ForegroundColor Red
        }
    }
    Write-Progress -Id 1000 -Activity “Processing Lync Users...” -status “Lync Users already completed: $counter from $membersCount” -percentComplete ($counter / $Members.Count*100)
}
Write-Progress -Id 1000 -Activity “Processing Lync Users...” -Completed
