# ##############################################################################
# Configuration
# ##############################################################################

# This setting controls whether the script is in testing mode or not (sends emails or not)
$global:testing = $true

# HTML email to send to users
$global:warnFile = 'Reminder.html'
$global:htmlEmail = $true

# PowerShell SMTP configuration
$global:smtp = "mail.yourdomain.com"
$global:smtpPort = 25
$global:from = "noreply@yourdomain.com"
$global:enableSSL = $false
$global:useDefaultCredentials = $false

# Configure how many days before expiration will the user receive emails (array of numbers)
$global:warningDays = @(14, 3)
# If the password expires in this or less than (or equal to) this amount of days, the email will be sent with high priority
$global:highPriorityDays = 3

# Root OU in Active Directory for searching for users
$global:ADSearchRoot = 'CN=Users,DC=yourdomain,DC=com'
# ##############################################################################


# ##############################################################################
# Logging
# ##############################################################################
$global:LogFolderName = $null
$global:LogFileName = $null
$global:Log = $null

$global:LogFolderName = "logs\"
$global:LogFileName = $LogFolderName
$global:LogFileName += Get-Date -Format o | foreach {$_ -replace ":", "."}
$global:LogFileName += ".log"

function SetUpLogs {
    $global:Log = $null

    if (Test-Path $LogFolderName) {
        Write-Host "Log folder exists"
    } else {
        mkdir $LogFolderName
    }
}

function WriteLog {
    Param ([string]$_logstring)

    $global:Log += $_logstring
    $global:Log += "`r`n"
}

function WriteLogToFile {
    Param(
        [string]$_logFile,
        [string]$_logString
    )

    Add-content $_logFile -value $_logString
}
# ##############################################################################

function ChangeTextInString {
    param (
        [string]$variable,
        [string]$textToChange,
        [string]$replacement
    )
    $result = $variable -replace $textToChange, $replacement

    return $result
}

function SendWarningMail {
    param
    (
        $user,
        [string] $from,
        [string] $message,
        [string] $priority
    )
    $date = Get-Date -Format g
    $days = GetPwdDays $user.PasswordExpires
    $to = $user.Mail
    $username = $user.SamAccountName
    $subject = 'Your Password Will Expire In ' + $days + ' Days!'

    if($priority -eq "Normal" -or $priority -eq "Low" -or $priority -eq "High") {
        # valid priority
    } else {
        # invalid priorty, overwrite to Normal
        $priority = "Normal"
    }

    if (-not $to) {
        WriteLog "ERROR: No email address for '$username'. User info: $user"
        return $false
    }

    $SmtpClient = New-Object System.Net.Mail.SmtpClient($global:smtp, $global:smtpPort)

    # If useDefaultCredentials is set to true, change the property on the SMTP client object.
    # It is false by default.
    $SmtpClient.UseDefaultCredentials = $global:useDefaultCredentials

    $SmtpClient.enableSSL = $global:enableSSL

    # Send the mail
    # $SmtpClient.Send($config.from, $to, $subject, $message)
	$mailMessage = New-Object System.Net.Mail.MailMessage
	$mailMessage.From = $from
	$mailMessage.To.Add($to)
    $mailMessage.Priority = $priority
    $mailMessage.Subject = $subject
    $mailMessage.Body = ($message)
  	$mailMessage.IsBodyHTML = $global:htmlEmail

    if ($global:testing -eq $false) {
        # Send email
        $SmtpClient.Send($mailMessage)
    }

    if ( $? ) {
        WriteLog "${date}: Sent $days day warning mail to $to (Username: $username)"
    }
    else {
        WriteLog "${date}: Error sending $days day warning mail to $to (Username: $username): " + $error[0].ToString()
    }

    # "Clear" the object/variable. Seems necessary.
    $SmtpClient = $null

    return $true
}


function CustomizeEmailTemplate {
    param (
        $private:ADUser,
        $private:emailTemplate
    )
    
    # ###########
    # Change placeholders in the HTML email
    $private:namePlaceholder = "{{full-name}}"
    $private:usernamePlaceholder = "{{username}}"
    $private:dayCountPlaceHolder = "{{remaining-day-number}}"
    # ###########

    $private:fullName = $private:ADUser.GivenName + " " + $private:ADUser.Sn
    $private:userName = $private:ADUser.SamAccountName
    $private:dayCount = GetPwdDays $private:ADUser.PasswordExpires

    $filledEmail = $private:emailTemplate
    $filledEmail = ChangeTextInString $filledEmail $private:namePlaceholder $private:fullName
    $filledEmail = ChangeTextInString $filledEmail $private:usernamePlaceholder $private:userName
    $filledEmail = ChangeTextInString $filledEmail $private:dayCountPlaceHolder $private:dayCount

    return $filledEmail
}

function GetPwdDays {
    param(
        [datetime] $private:pwdExpires
    )

    ($private:pwdExpires - $(Get-Date)).Days
}


function GetAllAdUsers {
    param (
        [string]$private:searchRoot
    )
    $private:users

    $private:users = Get-ADUser -SearchBase $private:searchRoot -ResultSetSize $Null -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties SamAccountName, GivenName, Sn, Mail, "msDS-UserPasswordExpiryTimeComputed" | Select-Object -Property SamAccountName, GivenName, Sn, Mail, @{Name="PasswordExpires"; Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}

    return $private:users
}

function GetUsersWithExpirationWarning {
    param (
        $private:users,
        $private:warningDayAmount
    )

    $private:warnUsers = $private:users | Where-Object { $_.PasswordExpires -and ((GetPwdDays $_.PasswordExpires) -eq $private:warningDayAmount) }
    return $private:warnUsers
}

function RunForSingleUser {
    param (
        $private:user
    )
    $private:templateEmail = Get-Content $global:warnFile
    $private:filledEmail = CustomizeEmailTemplate $private:user $private:templateEmail
    $private:priority = "Normal"
    $private:days = GetPwdDays $user.PasswordExpires
    if ($private:days -le $global:highPriorityDays) {
        $private:priority = "High"
    }

    SendWarningMail $private:user $global:from $private:filledEmail $private:priority
}

function RunForAllUsers {
    param (
    )
    # get every AD user in the specified path
    $private:allUsers = GetAllAdUsers $global:ADSearchRoot

    # Get users with each warning days
    for ($i = 0; $i -lt $global:warningDays.Count; $i++) {
        $warnDays = $global:warningDays[$i]
        $private:warnUsers = GetUsersWithExpirationWarning $private:allUsers $warnDays

        WriteLog "users with $warnDays days left:"

        foreach ($user in $private:warnUsers) {
            RunForSingleUser $user
        }

    }

}


function Main {
    SetUpLogs

    WriteLog "--------------------------Config---------------------------------"
    WriteLog "Email template: $global:warnFile"
    WriteLog "SMTP: $global:smtp at port $global:smtpPort"
    WriteLog "Sending address: $global:from"
    WriteLog "Emails will be sent with number of days to expiration: $global:warningDays"
    WriteLog "AD search root: $global:ADSearchRoot"
    WriteLog "Emails will be sent with high priority if expiration day is less or equal than $global:highPriorityDays days"
    WriteLog "-----------------------------------------------------------------"

    RunForAllUsers


    WriteLogToFile $global:LogFileName $global:Log
}


Main
