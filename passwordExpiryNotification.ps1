import-module activedirectory
$currDate = get-date -Format yyyy-MM-dd
$log = "\\mdm-serv\c$\scripts\logs\passwordExpiry\$currDate.log"


if (Get-ChildItem -Path $log -ErrorAction SilentlyContinue) {

    #write-host "true"

    } else {

    #write-host "false"
    new-item $log -Type file

    }

add-content -path $log -value "TimeStamp || DisplayName || Email || DaysLeft"

$users = get-aduser -filter {(PasswordNeverExpires -ne $true) -and (emailaddress -like "*@hellermanntyton.co.uk") -and (enabled -eq $true) -and (PasswordExpired -ne $true)} -properties PasswordExpired, PasswordLastSet, PasswordNeverExpires, Emailaddress
#$users = get-aduser -filter {(samaccountname -eq "snowe")-and (enabled -eq $true)} -properties PasswordExpired, PasswordLastSet, PasswordNeverExpires, Emailaddress

foreach ($user in $users) {

    if (($user -notlike "*OU=Generic*") -and ($user -notlike "*CN=Users,DC=hellermanntytongroup,DC=com") -and ($user -notlike "*#Migrated Objects*")) {

        $passwordExpiry = $user.PasswordLastSet + ((Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.TotalDays)
        $DaysLeft = (New-TimeSpan -Start (Get-Date) -End $passwordExpiry).Days
        
        if (($DaysLeft -le 7) -and ($DaysLeft -gt 0)) {
        #if (!(($DaysLeft -le 7) -and ($DaysLeft -gt 0))) {

            $name = $user.givenname
            $email = $user.emailaddress
            $displayName = $user.Name
            
            #write-host "$name,$email,$daysleft"

            if ($daysLeft -gt 1) {

                $days = "$daysLeft days"

                } else {

                $days = "$daysLeft day"

                }

            $backGroundColour = "#FFFFFF"
            $timestamp = get-date -format dd/MM/yyyy" "HH:mm


            $head = "      <style>"
            $head = $head + ""
            $head = $head + "        p {"
            $head = $head + "	  font-family: Calibri, sans-serif;"
            $head = $head + "     font-size: 16px;"
            $head = $head + "        }"
            $head = $head + ""
            $head = $head + "	h1 {"
            $head = $head + "	  text-decoration: underline;"
            $head = $head + "     font-family: Calibri, sans-serif;"
            $head = $head + "	}"
            $head = $head + ""
            $head = $head + "   h3 {"
            $head = $head + "     font-family: Calibri, sans-serif;"
            $head = $head + "   }"
            $head = $head + ""
            $head = $head + "   p.sig {"
            $head = $head + "     font-family: Arial, sans-serif;"
            $head = $head + "     font-size: 13px;"
            $head = $head + "   }"
            $head = $head + ""
            $head = $head + "      </style>"
            $html = "<p>Hi $name,<br><br>"
            $html = $html + "Your password is due to expire in $days. Please follow the instructions below, in order to change it manually.<br>"
            $html = $html + "Thanks,</p><br>"
            $html = $html + "<p class=sig><b>IT Helpdesk</b><br>"
            $html = $html + "<font color=`"#000F7C`"><b>Hellermann</font><font color=`"#FE000B`">Tyton</font></b><br>"
            $html = $html + "Tel Number:    +44 (0) 161 947 2298<br>"
            $html = $html + "Internal Ext:  800<br>"
            $html = $html + "E-mail:         helpdesk@hellermanntyton.co.uk</p><br><br>"
            $html = $html + ""
            $html = $html + "    <h1>How to change your Windows Password Manually</h1><br>"
            $html = $html + ""
            $html = $html + "    <h3>Before trying the below process, make sure you are either in the office - <b>OR </b>connected to the VPN <b>first</b>!!</h3><br>"
            $html = $html + ""
            $html = $html + "    <p><b>Note: </b>This will <b><u>NOT </u></b>change your Aurora password.</p>"
            $html = $html + "    <p>1. Press <b>Ctrl</b> + <b>Alt</b> + <b>Del</b> at the same time and you'll get to the following screen: <br><br></p>"
            $html = $html + "    <img src='https://howto.hellermanntyton.co.uk/imgs/changePass/01.png'> <br><br><br><br>"
            $html = $html + "    <p>2. Click on <b>Change a password...</b> to get to the password change screen:<br><br></p>"
            $html = $html + "    <img src='https://howto.hellermanntyton.co.uk/imgs/changePass/02.png'> <br><br>"
            $html = $html + "    <p>3. Type in your old password, followed by a new password in the boxes provided. Click on the <img src='https://howto.hellermanntyton.co.uk/imgs/changePass/arrow.png'> button which should take you to the confirmation screen: <br><br></p>"
            $html = $html + "    <img src='https://howto.hellermanntyton.co.uk/imgs/changePass/03.png'> <br><br><br><br>"



            $smtpServer = "10.2.30.100"
            $smtpFrom = "helpdesk@hellermanntyton.co.uk"
            $smtpTo = $email
            $messageSubject = "Windows Password Expiry Notification"

            $message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto
            $message.Subject = $messageSubject
            $message.IsBodyHTML = $true

            $message.Body =  ConvertTo-Html -Body $html -Head $head

            $smtp = New-Object Net.Mail.SmtpClient($smtpServer)
            $smtp.Send($message)

            #write-host "$timestamp,$displayName,$email"
            add-content -path $log -value "$timestamp || $displayName || $email || $DaysLeft"

        }

    }

}

