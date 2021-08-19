# Active Directory Password Expiration Reminder

Have you ever needed to send password expiration reminder emails for your (on-prem) Active Directory users? Microsoft for some reason does not provide an easy way to send email reminders to users. While there is a [Group Policy to show reminders in Windows on logon](https://docs.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/interactive-logon-prompt-user-to-change-password-before-expiration), email reminders are still nice to send to users. This script helps you achieving this!

You can configure SMTP server, the HTML email file location, the number of days when the script sends an email before expiration (can be multiple times as well) and some more.

## Requirements

The script (PowerShell 5 was tested) needs to be ran on a domain joined Windows machine (Server or Desktop), because of the Active Directory integration. It also needs to have the [Active Directory module](https://docs.microsoft.com/en-us/powershell/module/activedirectory) installed (on non-server Windows versions, this is achieved by the Remove Server Administration Tools (RSAT)).

## Usage

Set up the configuration in the beginning of `Password-Expiration-Notifier.ps1`, provide an email template file, and you're basically good to go. Make sure to have the `testing` value set to `false` if you want to send emails to your users (you're ready to use the script for real).

There is an example `Reminder.html` email that is used by the script currently, but this needs to be converted into an inline CSS HTML file in order for it to work properly in emails (emails in general can only work with inline CSS). An example and easy to use tool for this is <https://htmlemail.io/inline/>, but there are others out there as well (just paste the source code and you'll recieve and inlined CSS source code, which then can be sent over email).

_The script sends an HTML email by default, but can be configured to send plain text as well._

### Dynamic values in the email

The script will replace the following content in the provided email file to dynamic values:

Template string            | Purpose
---------------------------|-----------------------------------------
`{{full-name}}`            | Full name of the AD User
`{{username}}`             | SamAccountName (username) of the AD user
`{{remaining-day-number}}` | Days until password expiration

### Scheduled run

To constantly keep your users informed about their password expiration status, run the script once a day as a Windows Scheduled Task. This way they will be informed when there password expiration is nearing and they need to take action.
