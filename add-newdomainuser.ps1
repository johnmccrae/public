
    function Add-MyDomainADUser
{
	[CmdletBinding()]
	param
	(
        [Parameter(ValueFromPipeline)]
        [switch]$DontSendEmail
	)
	begin {
		$ErrorActionPreference = 'Stop'

        Add-Type -AssemblyName System.Drawing
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName mscorlib

        import-module activedirectory

        $mgrlist = (Get-ADuser -SearchBase "OU=Employees,OU=Domain Users,DC=mydomain,DC=com" -filter {isaPeopleManager -eq $true} -Properties mail, displayname | Sort Givenname  )

	}
    process {
        $form = New-Object System.Windows.Forms.Form
	    $form.Size = New-Object System.Drawing.Size(470, 616)
	    $form.FormBorderStyle = 'FixedDialog'
	    $form.MaximizeBox = $False
	    $form.MinimizeBox = $False
	    $form.Name = "form1"
	    $form.StartPosition = 'CenterScreen'
	    $form.Text = "New AD User"

        $OKbutton = New-Object System.Windows.Forms.Button
	    $OKbutton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	    $OKbutton.Location = New-Object System.Drawing.Size(180, 555)
	    $OKbutton.Size = New-Object System.Drawing.Size(75, 23)
	    $OKbutton.TabIndex = 10
	    $OKbutton.Text = "&OK"
        $form.AcceptButton = $OKbutton
        $form.Controls.Add($OKbutton)
        #
        # Cancel Button
        #
        $Cancelbutton = New-Object System.Windows.Forms.Button
        $Cancelbutton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $Cancelbutton.Location = New-Object System.Drawing.Size(270, 555)
        $Cancelbutton.Size = New-Object System.Drawing.Size(75, 23)
        $Cancelbutton.TabIndex = 100
	    $Cancelbutton.Text = "C&ancel"
        $form.AcceptButton = $Cancelbutton
        $form.Controls.Add($Cancelbutton)
        #
	    # New Employee Label
	    #
	    $lblNewEmployee = New-Object System.Windows.Forms.Label
        $lblNewEmployee.Location = New-Object System.Drawing.Size(12, 13)
	    $lblNewEmployee.Size = New-Object System.Drawing.Size(175, 23)
	    $lblNewEmployee.Text = "Enter the new users full name"
        $form.Controls.Add($lblNewEmployee)
	    #
	    # Emplouyee textbox
	    #
        $txtNewEmployee = New-Object System.Windows.Forms.TextBox
	    $txtNewEmployee.Location = New-Object System.Drawing.Size(12, 40)
	    $txtNewEmployee.Size = New-Object System.Drawing.Size(174, 20)
	    $txtNewEmployee.TabIndex = 0
        $txtNewEmployee.TabStop = $True
        $form.Controls.Add($txtNewEmployee)
	    #
	    # Contractor label
	    #
	    $lblContract = New-Object System.Windows.Forms.Label
        $lblContract.Location = New-Object System.Drawing.Size(12, 70)
	    $lblContract.Size = New-Object System.Drawing.Size(100, 23)
	    $lblContract.Text = "New Contractor?"
        $form.Controls.Add($lblContract)
        #
	    # Contractor checkbox
	    #
        $chkContractor = New-Object System.Windows.Forms.CheckBox
	    $chkContractor.Location = New-Object System.Drawing.Size(115, 65)
	    $chkContractor.Size = New-Object System.Drawing.Size(104, 24)
	    $chkContractor.Text = "Yes!"
        $chkContractor.TabIndex = 1
        $chkContractor.TabStop = $True
        $form.Controls.Add($chkContractor)

	    #
	    # Label for Contract Date
	    #
	    $lblContractDate = New-Object System.Windows.Forms.Label
        $lblContractDate.Location = New-Object System.Drawing.Size(12, 105)
	    $lblContractDate.Size = New-Object System.Drawing.Size(150, 20)
	    $lblContractDate.Text = "Pick their end-date"
        $form.Controls.Add($lblContractDate)
	    #
	    # Calendar Control to choose End-Dates for Contractors with
	    #
        $monthcalendar1 = New-Object System.Windows.Forms.MonthCalendar
	    $monthcalendar1.Location = New-Object System.Drawing.Size(12, 125)
	    $monthcalendar1.TabIndex = 2
        $monthcalendar1.TabStop = $True
        $form.Controls.Add($monthcalendar1)
	    #
	    # Label for Job Description
	    #
        $lblDescription = New-Object System.Windows.Forms.Label
	    $lblDescription.Location = New-Object System.Drawing.Size(12, 300)
	    $lblDescription.Size = New-Object System.Drawing.Size(100, 15)
	    $lblDescription.Text = "User Role"
        $form.Controls.Add($lblDescription)

        #
	    # txtDescription - User Job Description
	    #
        $txtDescription = New-Object System.Windows.Forms.TextBox
	    $txtDescription.Location = New-Object System.Drawing.Size(12, 315)
	    $txtDescription.Size = New-Object System.Drawing.Size(225, 20)
	    $txtDescription.TabIndex = 3
        $txtDescription.TabStop = $true
        $form.Controls.Add($txtDescription)

	    #
	    # User Title Label
	    #
        $lblTitle = New-Object System.Windows.Forms.Label
	    $lblTitle.Location = New-Object System.Drawing.Size(12, 335)
	    $lblTitle.Size = New-Object System.Drawing.Size(255, 13)
	    $lblTitle.Text = "User's New Title"
        $form.Controls.Add($lblTitle)
        #
        # User Title Text Box
        #
        $txtTitle = New-Object System.Windows.Forms.TextBox
	    $txtTitle.Location = New-Object System.Drawing.Size(12, 350)
	    $txtTitle.Size = New-Object System.Drawing.Size(225, 20)
	    $txtTitle.TabIndex = 4
        $txtTitle.TabStop = $True
        $form.Controls.Add($txtTitle)
	    #
	    # Office Location Label
	    #
 
    
    
        $lblOffice = New-Object System.Windows.Forms.Label
	    $lblOffice.Location = New-Object System.Drawing.Size(12, 370)
	    $lblOffice.Size = New-Object System.Drawing.Size(300, 15)
	    $lblOffice.Text = "Select the users Office location"
        $form.Controls.Add($lblOffice)
	    #
        # Office Location Drop-Down list
        #
        [array]$UserLocaleArray = "Amsterdam","Belfast","Berlin","Canada","London","Melbourne","San Francisco","Seattle","Sydney","Remote"
    
        # This Function Returns the Selected Value and their actions then Closes the Form

        $DropDown = new-object System.Windows.Forms.ComboBox
        $DropDown.Location = new-object System.Drawing.Size(12,385)
        $DropDown.Size = new-object System.Drawing.Size(225,20)

        ForEach ($Item in $UserLocaleArray) {
	        $DropDown.Items.Add($Item) | Out-Null
        }

        $Form.Controls.Add($DropDown)


	    #
	    # Group Label
	    #
        $lblgrp = New-Object System.Windows.Forms.Label
	    $lblgrp.Location = New-Object System.Drawing.Size(12, 405)
	    $lblgrp.Size = New-Object System.Drawing.Size(255, 13)
	    $lblgrp.Text = "Choose a Team"
        $form.Controls.Add($lblgrp)
        #
        # Group Option Text Box
        #

        [array]$UserGroupArray = "Sales", "Marketing", "CSM", "CSE", "SA", "EE", "Finance", "Legal", "The Herd", "Core Engineering", "Restricted User"
    
        # This Function Returns the Selected Value and their actions then Closes the Form

        $DropDown2 = new-object System.Windows.Forms.ComboBox
        $DropDown2.Location = new-object System.Drawing.Size(12,420)
        $DropDown2.Size = new-object System.Drawing.Size(225,20)

        ForEach ($Team in $UserGroupArray) {
	        $DropDown2.Items.Add($Team) | Out-Null
        }

        $Form.Controls.Add($DropDown2)

	    #
	    # User Department Label
	    #
        $lblDept = New-Object System.Windows.Forms.Label
	    $lblDept.Location = New-Object System.Drawing.Size(12, 440)
	    $lblDept.Size = New-Object System.Drawing.Size(255, 13)
	    $lblDept.Text = "Enter a Department Name"
        $form.Controls.Add($lblDept)
        #
        # User Department Text Box
        #
        $txtDept = New-Object System.Windows.Forms.TextBox
	    $txtDept.Location = New-Object System.Drawing.Size(12, 455)
	    $txtDept.Size = New-Object System.Drawing.Size(225, 20)
	    $txtDept.TabIndex = 7	
        $txtDept.TabStop = $true
        $form.Controls.Add($txtDept)
	    #
	    # User Manager Label
	    #
        $lblMgr = New-Object System.Windows.Forms.Label
	    $lblMgr.Location = New-Object System.Drawing.Size(12, 475)
	    $lblMgr.Size = New-Object System.Drawing.Size(255, 13)
	    $lblMgr.Text = "Enter Employees' Manager Name"
        $form.Controls.Add($lblMgr)
        #
        # User Manager Combox Box
        #
        $mgrdropdown = New-Object System.Windows.Forms.ComboBox
        $mgrdropdown.Location = new-object System.Drawing.Size(12,490)
        $mgrdropdown.Size = new-object System.Drawing.Size(225,20)
        $mgrdropdown.TabIndex = 8
        $mgrdropdown.TabStop = $true

        ForEach ($mgr in $mgrlist) {
	        $mgrdropdown.Items.Add($mgr.displayname) | Out-Null

        }

        $Form.Controls.Add($mgrdropdown)
	    #
	    # Your Initials Label
	    #
        $lblInits = New-Object System.Windows.Forms.Label
	    $lblInits.Location = New-Object System.Drawing.Size(12, 510)
	    $lblInits.Size = New-Object System.Drawing.Size(255, 13)
	    $lblInits.Text = "Enter Your Initials"
        $form.Controls.Add($lblInits)
        #
        # Your Initials Text Box
        #
        $txtInits = New-Object System.Windows.Forms.TextBox
	    $txtInits.Location = New-Object System.Drawing.Size(12, 525)
	    $txtInits.Size = New-Object System.Drawing.Size(143, 20)
	    $txtInits.TabIndex = 9
        $txtInits.TabStop = $True
        $form.Controls.Add($txtInits)

        $groupBox = New-Object System.Windows.Forms.GroupBox
        $groupBox.Location = New-Object System.Drawing.Size(250,20) 
        $groupBox.size = New-Object System.Drawing.Size(200,500) 
        $groupBox.text = "Choose" 
        $Form.Controls.Add($groupBox) 

        $Checkbox1 = New-Object System.Windows.Forms.CheckBox 
        $Checkbox1.Location = new-object System.Drawing.Point(15,15) 
        $Checkbox1.size = New-Object System.Drawing.Size(110,20) 
        $Checkbox1.Text = "Developer?" 
        $groupBox.Controls.Add($Checkbox1) 

        $Radiobutton1 = New-Object System.Windows.Forms.RadioButton
        $Radiobutton1.Location = New-Object System.Drawing.Point(15, 45)
        $Radiobutton1.Size = New-Object System.Drawing.Size(130, 20)
        $Radiobutton1.Checked = $true
        $Radiobutton1.Text = "U.S. Employee?"
        $groupBox.Controls.Add($Radiobutton1)

        $Radiobutton2 = New-Object System.Windows.Forms.RadioButton
        $Radiobutton2.Location = New-Object System.Drawing.Point(15, 75)
        $Radiobutton2.Size = New-Object System.Drawing.Size(150, 20)
        $Radiobutton2.Text = "EMEA/APAC Employee?"
        $groupBox.Controls.Add($Radiobutton2)


        $Checkbox2 = New-Object System.Windows.Forms.CheckBox 
        $Checkbox2.Location = new-object System.Drawing.Point(15,105) 
        $Checkbox2.size = New-Object System.Drawing.Size(80,20) 
        $Checkbox2.Text = "Office 365" 
        $groupBox.Controls.Add($Checkbox2) 

        $isaManagerCheckBox = New-Object System.Windows.Forms.CheckBox
        $isaManagerCheckBox.Location =  new-object System.Drawing.Point(15,135)
        $isaManagerCheckBox.Size = New-Object System.Drawing.Size(130,20)
        $isaManagerCheckBox.Text = "People Manager?"
        $groupBox.Controls.Add($isaManagerCheckBox) 

        $timeSheetsUserCheckbox = New-Object System.Windows.Forms.CheckBox
        $timeSheetsUserCheckbox.Location =  new-object System.Drawing.Point(15,165)
        $timeSheetsUserCheckbox.size = New-Object System.Drawing.Size(85,40)
        $timeSheetsUserCheckbox.Text = "TimeSheets User?"
        $groupBox.Controls.Add($timeSheetsUserCheckbox)

	    $zendeskCheckbox = New-Object System.Windows.Forms.CheckBox 
        $zendeskCheckbox.Location = new-object System.Drawing.Point(15,195) 
        $zendeskCheckbox.size = New-Object System.Drawing.Size(100,40) 
        $zendeskCheckbox.Text = "Zendesk user?" 
        $groupBox.Controls.Add($zendeskCheckbox)

        $result = $form.ShowDialog()

        Write-Host $mgrdropdown.SelectedItem.ToString()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK)
        {
		    
            Get-ChefControlAccounts

            #The default password reflects the current change to use a 12-character password of random complexity set out by NIST. 
            $defaultpass = "somepassword"
		    $DC = $env:COMPUTERNAME
       		
            $Office = $DropDown.SelectedItem.ToString()
            $tmpSecGroups = $DropDown2.SelectedItem.ToString()
		
		    $newemployee = $txtNewEmployee.Text
 
            # Logic for Security Groups in AD and Email Groups in Google
            [string[]] $googlegroupsarray = "Allyalls@mydomain.com", "backchannel@mydomain.com"

            [string[]] $SecGroups = "Entire Company"

            #"Marketing", "TAM", "CSE", "SA", "EE", "Finance", "The Herd"
            switch ($tmpSecGroups)
            {
                "Sales" {[string[]]$googlegroupsarray += "sales@mydomain.com"}
                "Marketing" {[string[]] $SecGroups += "Marketing"}
                "Finance" {[string[]] $SecGroups += "Finance"; [string[]] $googlegroupsarray += "finance@mydomain.com"}
                "Legal" {[string[]] $SecGroups += "Legal Team"}
                "EE" {[string[]] $SecGroups += "Employee Experience"}
                "Core Engineering" {[string[]]$googlegroupsarray += "core-eng-team@mydomain.io"}
                "External Contractors" {[string[]] $SecGroups = "External Contractors"}
            }

            # Capture the Country locaton of the employee for Tax and Expense Reimbursement purposes
            $Locale = ""
            switch ($office)
            {
                "Amsterdam"     {$Locale = "The Netherlands"}
                "Berlin"        {$Locale = "Germany"}
                "Canada"        {$Locale = "Canada"}
                "London"        {$Locale = "UK"}
                "Melbourne"     {$Locale = "Australia"}
                "San Francisco" {$Locale = "US"}
                "Seattle"       {$Locale = "US"}
                "Sydney"        {$Locale = "Australia"}
                "Remote"        {$Locale = "US"}           
            }

            if ($Office -eq "Remote")
            {
                [string[]] $googlegroupsarray += "remotestaff@mydomain.com"  
            }

            if ($Office -eq "Seattle")
            {
                [string[]] $googlegroupsarray += "HQ@mydomain.com"  
            }

            if ($Radiobutton1.Checked)
            {
                [string[]] $googlegroupsarray += "Us@mydomain.com"
                [string[]] $SecGroups += "US"
            }
            else
            {
                [string[]] $googlegroupsarray += "europe_team@mydomain.com"  
                [string[]] $SecGroups += "EMEA"
            }
        
            if ($Checkbox1.Checked -eq $True )
             {
          	    [string[]] $SecGroups += "AD AWS","devteam","vpnusers", "All Developers"
                [string[]] $googlegroupsarray += "dev@mydomain.com"
             }

            if ($Checkbox2.Checked -eq $True)
            {
                [string[]] $SecGroups += "Office365 Users"
            }


            if ($chkContractor.Checked)
            {
                $Org = "OU=Contractors,OU=Domain Users,DC=MYDOMAIN,DC=COM"
                $date = $monthcalendar1.SelectionStart
			    if (($date -as [DateTime]) -ne $null)
        	    {
                $timetemp = Get-Date -Date $date
                $contractdate = $timetemp.AddDays(1)
        	    } 
                [string[]] $googlegroupsarray = "contingent@mydomain.com"
            }
            else
            {
                $Org = "OU=Employees,OU=Domain Users,DC=MYDOMAIN,DC=COM"
            }
            if ($timeSheetsUserCheckbox.Checked)
            {
                [string[]] $googlegroupsarray += "timesheets@mydomain.com"
            }

		    # What department do they work in
		    $department = $txtDept.Text
            $title = $txtTitle.Text
		
		    # find the mgr email and sAMAccountName from the mgrlist
            $mgremail = ""
            $mgraccount = ""
            $manager = ""
            $mgrtmp = $mgrdropdown.SelectedItem
            ForEach ($mgr in $mgrlist) {
	            if($mgrtmp -eq $($mgr.displayname)){
                Write-host "This is the correct manager: $($mgr.SamAccountName)"
                $mgraccount = $mgr.SamAccountName
                $mgremail = $mgr.mail
                $manager = $mgr.displayname
                }

            }

		    # extract the email and upn
		    $pos = $newemployee.IndexOf(" ")
            $firstname = $newemployee.Substring(0, $pos)
            $lastname = $newemployee.Substring($pos+1)
            $username = $firstname[0] + $lastname
            $email = $username + "@mydomain.io" 
		
		    # Get the initials of the person who created this
		    $your_initials = $txtInits.Text
		
		    # Confirming the basics of the user account
            Write-Host "The new employee: $($txtNewEmployee.Text)"
		    Write-Host "User Description: $($txtDescription.Text)" 	      
            Write-Host "Working in this office: $($Office)"
            Write-Host "User is a member of these security groups: $($SecGroups)"
            Write-host "User is a member of these Google Groups: $($googlegroupsarray)"
            Write-Host "The Manager selected is: $($mgrtmp)"


            #actually create the user accounts now.
            Write-Host "Creating AD User Account (Step 1 of 9)"
		    New-ADUser -AccountPassword (ConvertTo-SecureString -AsPlainText $defaultpass -Force) -ChangePasswordAtLogon $False -Company "MyCompany" -Department $department -Description $txtDescription.Text -DisplayName $newemployee -EmailAddress $email.ToLower() -Enabled $true -GivenName $firstname -Manager $mgraccount -Name $newemployee -Office $Office -Path $Org -SamAccountName $username.ToLower() -Server $DC -Surname $lastname -Title $title -UserPrincipalName $email.ToLower()

            Start-Sleep -s 5
            # we had to set a sleep here as we hit times where creating the AD account was going slower than the code below which updates the account and that caused this script to blow up - can't update an account that doesn't exist yet


            if ($tmpSecGroups -ne "Restricted User")
            {
                foreach($secgroup in $SecGroups)
		        {
		            Add-ADGroupMember -Identity $secgroup -members $username -Server $DC
		        }
            }

		    Set-ADUser -Identity $username -Replace @{info="$($_.info) User created by " + $your_initials + " " + [System.DateTime]::Now}

		    if($chkContractor.Checked)
		    {
		        Set-ADUser -Identity $username -AccountExpirationDate $contractdate
		    } 

            if($isaManagerCheckBox.Checked)
            {

                $ismanagerHashTable = New-Object HashTable
                $ismanagerHashTable.Add("isaPeopleManager", $true)
                set-aduser -identity $username  -replace $ismanagerHashTable

            }

            # Adding the user to Google and assigning them their Google Groups based on membership choices above
            Write-Host "Creating Google User Account (Step 2 of  9)"
            C:\gam-64\gam.exe create user $($email) firstname $($firstname) lastname $($lastname) password $($defaultpass)
            Write-host ""


            Write-Host "Adding user to Google Groups (Step 3 of 9)"
            foreach ($googlegroup in $googlegroupsarray)
            {

                c:\gam-64\gam.exe update group $($googlegroup) add user $($email) 

            }
            Write-host ""

            #Give the new user a Zoom Account.
            write-host "Starting Zoom Account Creation (Step 4 of 9)"
            try 
            {
                $body = @{api_key='<key>';api_secret='<secret>';email=$email;type=2;password=$defaultpass} 
                $returncode = Invoke-RestMethod -Method Post -Uri https://api.zoom.us/v1/user/autocreate -Body $body
                if($returncode.error)
                {
                   Write-Error $returncode.error
                }
                else
                {
                    Write-Host "Zoom Account created"
                }
            }
            catch
            {
                Write-Host "We encountered an error adding the user to Zoom, noted below"
                Write-Error $_.Exception.Message
                Write-Error $_.Exception.ItemName
            }


            <#
            #Give the new User a Trello Account - IEX = Invoke-Expression (Blocked due to token issue EPB 5/2018)
            Write-Host "Starting Trello Account creation (Step 5 of 9)"
            try 
            {
                Get-TrelloConfiguration | Out-Null
                Invoke-Command -ScriptBlock {Set-TrelloUser -email $email -fullName $newemployee -type normal} 
                Write-Host "Trello Account created"
            }
            catch
            {
                Write-Host "We encountered an error adding the user to Trello, noted below"
                Write-Error $_.Exception.Message
                Write-Error $_.Exception.ItemName
            }
            #>

            #Pre-create a user account in Atlassian
            Write-Host "Starting Atlassian account creation (Step 6 of 9)"

             # $mykeys is a PowerShell object that gets loaded into memory on our local systems when we start a powershell session; it loads in the profile.ps1.
            foreach ($key in $mykeys) {
                if ($key.Account -eq "ITAutomation") {
                    $apiUserName = "itautomation1"
                    $apiPlainPassword = $key.AccountPassword
                }
            }
            
            $SecurePassword = $apiPlainPassword | ConvertTo-SecureString -AsPlainText -Force
            $Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $apiUserName, $SecurePassword           
            
            try 
            {
                New-JiraSession -Credential $Credentials
                Invoke-Command -ScriptBlock {New-JiraUser -UserName $username  -EmailAddress $email -displayname $newemployee -Notify $true} 
                Write-Host "Atlassian Account created"
            }
            catch
            {
                Write-Host "We encountered an error adding the user to Atlassian, noted below"
                Write-Error $_.Exception.Message
                Write-Error $_.Exception.ItemName
            }

            write-host "Writing notes to the event log (Step 8 of 9)"
            Write-EventLog -LogName Application -Source "Add-MyDomainUser" -EntryType Information -EventId 1000 -Category 1 -Message "New User Account created for user $newemployee"


            # Email the manager with the users new account information. 
		    # Email Finance to create the new users Expensify account for regular employees only, not contractors
		    # Eamil Support to create the new users Zendesk account for inside sales or field sales employee 
		    ############################################################################## 
    
            function mailstuff    {
		
		        param (
		        [Parameter(ValueFromPipeline)]
		        [String]$tmpSendTo,
		        [Parameter(ValueFromPipeline)]
		        [String]$tmpSubject,
		        [Parameter(ValueFromPipeline)]
		        [String]$tmpBody		
		        )
		
            $SMTPUserName = ""
            $SMTPPassword = ""
		
                foreach ($key in $mykeys) {
                    if ($key.Account -eq "itconfig") {
                        $SMTPUserName = $key.AccountEmail
                        $SMTPPassword = $key.AccountPassword
                    }
                }

		
		    $Emailfrom = "itconfig@mydomain.io"
		    $EmailCC = "it@mydomain.io"
   
            $SMTPServer = "smtp.gmail.com"
            $SMTPPort = "587"

            $mailmessage = New-Object System.Net.Mail.MailMessage
            $mailmessage.From = ($Emailfrom)
            $mailmessage.CC.Add($EmailCC)
            $mailmessage.To.Add($tmpSendTo)
            $mailmessage.Subject = $tmpSubject
            $mailmessage.Body = $tmpBody
            $mailmessage.IsBodyHtml = $true

            $SMTPClient = New-Object System.Net.Mail.SmtpClient($SMTPServer,$SMTPPort)
            $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTPUsername, $SMTPPassword)
            $SMTPClient.EnableSsl = $true
            $SMTPClient.Send($mailmessage)
		
		    }
		
		    $tmpSendTo = $mgremail
		    $tmpSubject = "New Employee Account Information"
		    $tmpBody = "Good Day, " + $manager + "</br> Your new employee " + $firstname + " " + $lastname + " is starting work soon. </br> Their email address is: " + $email + "</br> Their default password is: " + $defaultpass + "<p><p>Cheers! <p>Your Friends in IT"

            Write-Host "Finishing off by emailing everyone (Step 9 of 9)"
            if(-not ($DontSendEmail))
            {             
                mailstuff -tmpSendTo $tmpSendTo -tmpSubject $tmpSubject -tmpBody $tmpBody 
            }


            if	(!$chkContractor.Checked -and $DontSendEmail -eq $false)
	        {

		            $tmpSendTo = "finance@mydomain.io"
		            $tmpSubject = "New Employee Expensify Account Request"
		            $tmpBody = "Good Day, <p> </br> Please create a new Expensify account for: <p>" + $firstname + " " + $lastname + " <p>Email: " + $email + "<p>Manager: " + $manager + ".</br> The employee works in the following location: " + $Locale + " <p><p>Cheers! <p>Your friends in IT"
	
                    mailstuff -tmpSendTo $tmpSendTo -tmpSubject $tmpSubject -tmpBody $tmpBody

	        }
	
	        if	($zendeskCheckbox.Checked -and $DontSendEmail -eq $false)	
	        {	
		    $tmpSendTo = "support@mydomain.io"
		    $tmpSubject = "New Employee Zendesk Account Request"
		    $tmpBody = "Good Day, <p> </br> Please create a new Zendesk account for: <p>" + $firstname + " " + $lastname + " <p>Email: " + $email + "<p>Manager: " + $manager + "<p><p>Cheers! <p>Your friends in IT"
	
            mailstuff -tmpSendTo $tmpSendTo -tmpSubject $tmpSubject -tmpBody $tmpBody  
	        }	
            ##############################################################################

            }
        else
        {
            $form.close()
        }

        }
    }
