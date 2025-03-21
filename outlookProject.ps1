###############################################################################################################################################################################
# Author: Elliot Abou-Antoun
# Date: March 13, 2025
# Revision: 1.1 (Revised March 21, 2025 )
# Purpose: Automate the process of sending initial, reminder and final notice emails for evergreening project
###############################################################################################################################################################################
###############################################################################################################################################################################
# Revision Notes.
# 
# Issue where sent emails would appear in personal inbox sent items and be labeled as personal user. (Resolved in code line 285, 286, 305.)
# 
# Issue where it applies personal inbox signatures to replies (Resolution is to turn off signature replies. STEPS TO DO SO IN THE README!!
#
# It currently looks through the last 500 emails from the top. This number can be lowered or raised at line 381, $FindMail.
#
#
# Fixed issue where $EndFlag, line 380, would not reset back to 0 after every successful or failed attempt. Now counter resets to make sure we start from 0 up to 500 per user in array.
#
#
#
#
#
#
#
#
#
#
#
#
#
###############################################################################################################################################################################

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Software Query Tool'
$form.Width = 980
$form.Height = 1050
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.AutoScroll = $true  # Enable scrolling if the controls exceed the form size
$form.AutoScrollMinSize = New-Object System.Drawing.Size(980, 1050)

# PictureBox (Logo Image)
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
$imagePath = "\\ircctech\Tech\Uploads\ElliotAA\Script\IRCC.png"
$pictureBox.Image = [System.Drawing.Image]::FromFile($imagePath)
$pictureBox.Location = New-Object System.Drawing.Point(0, 0)
$pictureBox.Size = New-Object System.Drawing.Size(1000, 90)
$pictureBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($pictureBox)

$deviceComboBox = New-Object System.Windows.Forms.ComboBox
$emailComboBox = New-Object System.Windows.Forms.ComboBox

# Create an Outlook application object
$outlook = New-Object -ComObject Outlook.Application

#$word = New-Object -ComObject Word.Application

# Get the namespace (MAPI)
$namespace = $outlook.GetNamespace("MAPI")

# Get the accounts in Outlook
$accounts = $namespace.Accounts

# Display the available accounts
 $deviceComboBox.Items.Add("IRCC.ITTechnologiesDelivery-LivraisondeTechnologies.IRCC@cic.gc.ca") 

$deviceComboBox.Location = New-Object System.Drawing.Point(20, 125)
$deviceComboBox.Text = "Select email address"
$deviceComboBox.Width = 350
$deviceComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($deviceComboBox)

$emailComboBox.Location = New-Object System.Drawing.Point(20, 175)
$emailComboBox.Text = "Select email type"
$emailComboBox.Items.Add("Initial Email")
$emailComboBox.Items.Add("Reminder Email")
$emailComboBox.Items.Add("Final Notice Email")
$emailComboBox.Width = 150
$emailComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($emailComboBox)

$namelabel = New-Object System.Windows.Forms.Label
$namelabel.Text = 'Username'
$namelabel.AutoSize = $true
$namelabel.Location = New-Object System.Drawing.Point(20, 225)
$namelabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($namelabel)


$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(20, 250)
$textBox.Width = 150
$textbox.Height = 20
$form.Controls.Add($textBox)

$assetlabel = New-Object System.Windows.Forms.Label
$assetlabel.Text = 'Asset tag'
$assetlabel.AutoSize = $true
$assetlabel.Location = New-Object System.Drawing.Point(250, 225)
$assetlabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($assetlabel)

$datelabel = New-Object System.Windows.Forms.Label
$datelabel.Text = 'Date'
$datelabel.AutoSize = $true
$datelabel.Location = New-Object System.Drawing.Point(150, 380)
$datelabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($datelabel)


$assettextBox = New-Object System.Windows.Forms.TextBox
$assettextBox.Location = New-Object System.Drawing.Point(250, 250)
$assettextBox.ReadOnly = $true
$assettextBox.Width = 150
$assettextBox.Height = 20
$form.Controls.Add($assettextBox)

$softwarelabel = New-Object System.Windows.Forms.Label
$softwarelabel.Text = 'Software'
$softwarelabel.AutoSize = $true
$softwarelabel.Location = New-Object System.Drawing.Point(600, 225)
$softwarelabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::left
$form.Controls.Add($softwarelabel)

$softwaretextBox = New-Object System.Windows.Forms.TextBox
$softwaretextBox.Location = New-Object System.Drawing.Point(475, 250)
$softwaretextBox.Multiline = $true
$softwaretextBox.ReadOnly = $true
$softwaretextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$softwaretextBox.Width = 300
$softwaretextBox.Height = 50
$form.Controls.Add($softwaretextBox)


$dateBox = New-Object System.Windows.Forms.TextBox
$dateBox.Location = New-Object System.Drawing.Point(150, 400)
$dateBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$dateBox.ReadOnly = $true
$dateBox.Width = 150
$dateBox.Height = 20
$form.Controls.Add($dateBox)

$addButton = New-Object System.Windows.Forms.Button
$addButton.Text = 'Add'
$addButton.Location = New-Object System.Drawing.Point(20, 325)
$addButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($addbutton)

$Generatebutton = New-Object System.Windows.Forms.Button
$Generatebutton.Text = 'Generate'
$Generatebutton.Enabled = $false
$Generatebutton.Location = New-Object System.Drawing.Point(20, 400)
$Generatebutton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($Generatebutton)

$Resetbutton = New-Object System.Windows.Forms.Button
$Resetbutton.Text = 'Reset'
$Resetbutton.Location = New-Object System.Drawing.Point(125, 325)
$Resetbutton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($Resetbutton)

$Resetbutton.Add_Click({
    $dataList.Clear()
    $textBox.Clear()
    $assettextBox.Clear()
    $softwaretextBox.Clear()
    $addBox.Clear()
    $Generatebutton.Enabled = $false

})


$addBox = New-Object System.Windows.Forms.RichTextBox
$addBox.Location = New-Object System.Drawing.Point(20, 500)
$addBox.Size = New-Object System.Drawing.Size(400, 350)
$addBox.ReadOnly = $true
$addBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($addBox)

$dataList = New-Object System.Collections.ArrayList
$emailComboBox.Add_SelectedIndexChanged({
    if ($emailComboBox.SelectedItem -eq "Final Notice Email") {
        $dateBox.ReadOnly = $false
        $assettextBox.ReadOnly = $true
        $softwaretextBox.ReadOnly = $true
    }
    else { $dateBox.ReadOnly = $true}

    if($emailComboBox.SelectedItem -eq "Initial Email") {
        $assettextBox.ReadOnly = $false 
        $softwaretextBox.ReadOnly = $false
        } else {
            $assettextBox.ReadOnly = $true
            $softwaretextBox.ReadOnly = $true
        }
})
$AddButton.Add_Click({
    $username = $textBox.Text
    $asset = $assettextBox.Text
    $software = $softwaretextBox.Text
    $date = $dateBox.Text
    $GenerateButton.Enabled = $true
    $addBox.AppendText("`n$username | $asset | $software | $date`n")


    $usernameSpliced = $username.Split(".")
    $global:formattedName = $usernameSpliced[1] + ", " + $usernameSpliced[0]

    #Write-Host $formattedName


    $dataList.Add([PSCustomObject]@{
        user = $username
        asset = $asset
        software = $software
        formattedName = $formattedName
        date = $date
        })
    #Write-Host $dataList
    $textBox.Clear()
    $assettextBox.Clear()
    $softwaretextBox.Clear()
    
        
})




$Generatebutton.Add_Click({

function Append-TextWithBackgroundColor {
    param (
        [Parameter(Mandatory=$true)]
        [string]$text,      # Text to append
        
        [Parameter(Mandatory=$true)]
        [string]$backColorName  # Background color name (e.g., "Yellow", "Green")
    )

    # Get the background color from the color name
    $backColor = [System.Drawing.Color]::FromName($backColorName)
    
    if ($backColor.IsEmpty) {
        Write-Host "Invalid background color name: $backColorName"
        return
    }

    # Set the background color of the selected text
    $addBox.SelectionBackColor = $backColor
    
    # Set the selection color (text color) to black (or any other color you prefer)
    $addBox.SelectionColor = [System.Drawing.Color]::Black
    
    # Append the text to the RichTextBox
    $addBox.AppendText("$text`n")
}
    $emailSelection = $emailComboBox.SelectedItem
    $selectedAccountName = $deviceComboBox.SelectedItem

    if($dataList -eq $null) {
        $addBox.AppendText("Please input user information.")
    }

    # Find the account object by matching the DisplayName
    $accountToUse = $null
    foreach ($account in $accounts) {
        if ($account.DisplayName -eq $selectedAccountName) {
            $accountToUse = $account
            break
        }
    }

    if ($accountToUse -eq $null) {
        Write-Host "Account not found!"
        return
    }
    Write-Host "Selected Account: $($accountToUse.DisplayName)"
    $sharedMailbox = $namespace.Folders.Item($accountToUse.SmtpAddress)
    $sentItemsFolder = $sharedMailbox.Folders.Item("Sent Items")
    if($emailSelection -eq "Initial Email") {
    foreach ($items in $dataList){
            $recipientEmail = "$($items.user)@cic.gc.ca"
            $username = $items.user
            $asset = $items.asset
            $software = $items.software

            # Create a new email item
            $mail = $outlook.CreateItem(0)
            # Set email parameters
            $mail.Subject = "ACTION REQUIRED: Replace your aging work device / Remplacez votre appareil de travail vieillissant"
            $htmlBody = Get-Content '\\ircctech\tech\uploads\Deployment Team\Evergreening Email Tool\initialemail.html' -Raw -Encoding UTF8
            $htmlBody = $htmlBody -replace '\$asset', $asset
            $htmlBody = $htmlBody -replace '\$software', $software
            $mail.SentOnBehalfOfName = $accountToUse.SmtpAddress
            $mail.HTMLBody = $htmlBody
            # Set To address
            $mail.To = $recipientEmail
            $mail.SaveSentMessageFolder = $sentItemsFolder


            $mail.Importance = 2
            $mail.ReadReceiptRequested = $true

            # Display the email (for review)

            $mail.Display()
            #$mail.Send()

          

    }
    $addBox.AppendText("`nInitial Email sent successfully from account: $($accountToUse.DisplayName)")
     $Generatebutton.Enabled = $false
    if($Generatebutton.Enabled -eq $false) {
        $dataList.Clear()
    }

}

if ($emailSelection -eq "Reminder Email") {
    # Get the default Sent Items folder from the selected account
    $account = $deviceComboBox.SelectedItem
    $accountToUse = $null
    # Get the account object by matching the DisplayName
    foreach ($item in $accounts) {
        if ($item.DisplayName -eq $account) {
            $accountToUse = $item
            break
        }
    }

    if ($accountToUse -eq $null) {
        Write-Host "Selected account not found"
        return
    }
    
    Write-Host "Account found: $($accountToUse.DisplayName)"

    # Try to get the Sent Items folder from the namespace for the selected account
    try {
        $sentItemsFolder = $namespace.Folders.Item($accountToUse.DisplayName).Folders.Item("Sent Items")
    } catch {
        Write-Host "Error accessing the Sent Items folder."
        return
    }

    # Check if the Sent Items folder is empty
    if ($sentItemsFolder.Items.Count -eq 0) {
        Write-Host "No items in Sent Items folder."
        return
    }

    Write-Host "Found $($sentItemsFolder.Items.Count) items in Sent Items folder."
    $matchFound = $false
    $nameFound = $true
    $endflag = 0
    foreach($user in $dataList){
            $username = $user.user
            $recipientEmail = "$username@cic.gc.ca"
            $formattedName = $user.formattedName
            # Define search criteria
            $searchSubject = "ACTION REQUIRED: Replace your aging work device / Remplacez votre appareil de travail vieillissant"  # Replace with your specific subject
            $searchRecipientEmail = "$formattedName (IRCC/IRCC)"  # Replace with your recipient email address

            # Get all sent items and sort them by the SentOn property in descending order
            $sentItems = $sentItemsFolder.Items
            $sentItems.Sort("[SentOn]", $true)  # Sort by sent date, ascending

            # Loop through the sent items and check the subject and recipient email
            foreach ($item in $sentItems) {
                # Debug: print item details to check for correct properties
               # Write-Host "Checking item: Subject: $($item.Subject), To: $($item.To)"
                    $endflag += 1
                    $FindMail = 500
                if($endflag -eq $FindMail){
                    Append-TextWithBackgroundColor "`nCould not find $searchSubject with $searchRecipientEmail in sent items folder within $endflag searches.`n" -backColorName "Yellow"
                    $endFlag = 0
                    break
                }

                if ($item.Subject -like "*$searchSubject*" -and $item.To -like "*$searchRecipientEmail*") {
                    # Create a reply to the email
                    $addBox.AppendText("Found a matching email, replying...")
                    $addBox.AppendText($searchRecipientEmail)
            
                    $replyMail = $item.Reply()


                    $replyMail.Subject = "REMINDER:RE:ACTION REQUIRED: Replace your aging work device / Remplacez votre appareil de travail vieillissant"
                    $htmlBody = Get-Content '\\ircctech\tech\uploads\Deployment Team\Evergreening Email Tool\reminderemail.html' -Raw -Encoding UTF8
                    $replyMail.HTMLBody = $replyMail.HTMLBody.Insert(0, $htmlbody)
                    $replyMail.SentOnBehalfOfName = $accountToUse.SmtpAddress
                    $replyMail.to = $recipientEmail
                    # You can add custom HTML or text for the reply body

                    # Set the "SentOnBehalfOfName" property to ensure the reply is sent from the correct account
                    $replyMail.SaveSentMessageFolder = $sentItemsFolder

           

                    # Send the reply email
                    $replyMail.Display()

                    # Log success
                    $addBox.AppendText("`nReminder Email sent successfully to: $searchRecipientEmail`n")
                    $addBox.AppendText("Reminder email sent.`n")
                    $endFlag = 0
                    $matchFound = $true
                    break
                    
                }
            }

            if(-not $matchFound) {
                Append-TextWithBackgroundColor "`nInvalid - No matching email found for $searchRecipientEmail.`n" -backColorName "Yellow"
                $endflag = 0
                $Generatebutton.Enabled = $false
                
            }
        }
    $Generatebutton.Enabled = $false
    if($Generatebutton.Enabled -eq $false) {
        $dataList.Clear()
    }
}







if($emailSelection -eq "Final Notice Email") {
    foreach ($items in $dataList){
            $date = $items.date
            $recipientEmail = "$($items.user)@cic.gc.ca"
            $username = $items.user
            $asset = $items.asset
            $software = $items.software

            # Create a new email item
            $mail = $outlook.CreateItem(0)
            # Set email parameters
            $mail.Subject = "FINAL NOTICE: Replace your aging work device / Remplacez votre appareil de travail vieillissant"
            $htmlBody = Get-Content '\\ircctech\tech\uploads\Deployment Team\Evergreening Email Tool\finalnotice.html' -Raw -Encoding UTF8
            $htmlBody = $htmlBody -replace '\$date', $date

            $mail.HTMLBody = $htmlBody
            $mail.SentOnBehalfOfName = $accountToUse.SmtpAddress
            $mail.To = $recipientEmail
            $mail.SaveSentMessageFolder = $sentItemsFolder
           

            $mail.Importance = 2
            $mail.ReadReceiptRequested = $true

            # Display the email (for review)
            $mail.Display()
            #$mail.Send()
    }
    $addBox.AppendText("`nFinal Notice Emails sent successfully from account: $($accountToUse.DisplayName)")
     $Generatebutton.Enabled = $false
    if($Generatebutton.Enabled -eq $false) {
        $dataList.Clear()
    }

}
})


$form.ShowDialog()
