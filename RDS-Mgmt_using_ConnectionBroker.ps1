
Function SetupForm {

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

	$objForm = New-Object System.Windows.Forms.Form
	$objForm.Text = "RD Session Manager"
	$objForm.Size = New-Object System.Drawing.Size(480,480)
	$objForm.StartPosition = "CenterScreen"
	$objForm.MaximizeBox = $False
	$objForm.ShowInTaskbar = $True

	$ApplyButton = New-Object System.Windows.Forms.Button
	$ApplyButton.Location = New-Object System.Drawing.Size(45,400)
	$ApplyButton.Size = New-Object System.Drawing.Size(75,23)
	$ApplyButton.Text = "Execute"
	$objForm.Controls.Add($ApplyButton)

	#Execute ScriptActions function when Apply is clicked
	$ApplyButton.Add_Click(
		{
			foreach ($objItem in $objListbox.SelectedItems)
				{$script:x += $objItem}
				$MessageText = $objMessageTextBox.Text
				$WaitTime = $objWaitforText.Text
				ScriptActions
         })


	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(180,400)
	$CancelButton.Size = New-Object System.Drawing.Size(75,23)
	$CancelButton.Text = "Cancel"
	$CancelButton.Add_Click({$objForm.Close()})
	$objForm.Controls.Add($CancelButton)

	
	$RefreshButton = New-Object System.Windows.Forms.Button
	$RefreshButton.Location = New-Object System.Drawing.Size(315,400)
	$RefreshButton.Size = New-Object System.Drawing.Size(75,23)
	$RefreshButton.Text = "Refresh"
	$objForm.Controls.Add($RefreshButton)
	$RefreshButton.Add_Click(
		{
			UpdateList
		})

	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(10,20)
	$objLabel.Size = New-Object System.Drawing.Size(380,20)
	$MyLabelText = "{0,-15}	{1,35}	{2,20}	{3,20}" -f "Username", "Server", "SessionID", "SessionState"
	$objLabel.Text = $MyLabelText
	$objForm.Controls.Add($objLabel)

	$objListBox = New-Object System.Windows.Forms.ListBox
	$objListBox.Location = New-Object System.Drawing.Size(10,50)
	$objListBox.Size = New-Object System.Drawing.Size(450,20)
	$objListBox.Height = 220
	$objListBox.SelectionMode = "MultiExtended"

	$objLabelMessage = New-Object System.Windows.Forms.Label
	$objLabelMessage.Location = New-Object System.Drawing.Size(10,315)
	$objLabelMessage.Size = New-Object System.Drawing.Size(280,20)
	$objLabelMessage.Text = "Message to send to user(s):"
	$objForm.Controls.Add($objLabelMessage)

	$objMessageTextBox = New-Object System.Windows.Forms.TextBox
	$objMessageTextBox.Location = New-Object System.Drawing.Size(10,335)
	$objMessageTextBox.Size = New-Object System.Drawing.Size(260,20)
	$objMessageTextBox.enabled = $false
	$objForm.Controls.Add($objMessageTextBox)

	$objLabelTime = New-Object System.Windows.Forms.Label
	$objLabelTime.Location = New-Object System.Drawing.Size(10,270)
	$objLabelTime.Size = New-Object System.Drawing.Size(200,20)
	$objLabelTime.Text = "Time in seconds before execution:"
	$objForm.Controls.Add($objLabelTime)

	$objWaitforText = New-Object System.Windows.Forms.TextBox
	$objWaitforText.Location = New-Object System.Drawing.Size(10,290)
	$objWaitforText.Size = New-Object System.Drawing.Size(70,20)
	$objForm.Controls.Add($objWaitforText)

	$objCheckSendMessage = New-Object System.Windows.Forms.CheckBox
	$objCheckSendMessage.Location = New-Object System.Drawing.Size(10,350)
	$objCheckSendMessage.Size = New-Object System.Drawing.Size(100,30)
	$objCheckSendMessage.Text = "Send message"
	$objForm.Controls.Add($objCheckSendMessage)
	$objCheckSendMessage.Add_CheckStateChanged({
		$objMessageTextBox.enabled = $objCheckSendMessage.Checked
		$objCheckDisconnect.Checked = $false
		$objCheckLogOff.Checked = $false
		$objMirrorSession.Checked = $false
	})

	$objCheckLogOff = New-Object System.Windows.Forms.CheckBox
	$objCheckLogOff.Location = New-Object System.Drawing.Size(120,350)
	$objCheckLogOff.Size = New-Object System.Drawing.Size(100,30)
	$objCheckLogOff.Text = "Logoff"
	$objForm.Controls.Add($objCheckLogOff)
	$objCheckLogOff.Add_CheckStateChanged({
		$objCheckDisconnect.Checked = $false
		$objCheckSendMessage.Checked = $false
		$objCheckDisconnect.Checked = $false
	})

	$objCheckDisconnect = New-Object System.Windows.Forms.CheckBox
	$objCheckDisconnect.Location = New-Object System.Drawing.Size(230,350)
	$objCheckDisconnect.Size = New-Object System.Drawing.Size(100,30)
	$objCheckDisconnect.Text = "Disconnect"
	$objForm.Controls.Add($objCheckDisconnect)
	$objCheckDisconnect.Add_CheckStateChanged({
		$objCheckLogOff.Checked = $false
		$objCheckSendMessage.Checked = $false
		$objMirrorSession.Checked = $false
	})

	$objMirrorSession = New-Object System.Windows.Forms.CheckBox
	$objMirrorSession.Location = New-Object System.Drawing.Size(340,350)
	$objMirrorSession.Size = New-Object System.Drawing.Size(100,30)
	$objMirrorSession.Text = "Mirror"
	$objForm.Controls.Add($objMirrorSession)
	$objMirrorSession.Add_CheckStateChanged({
		$objCheckLogOff.Checked = $false
		$objCheckSendMessage.Checked = $false
		$objCheckDisconnect.Checked = $false
	})

	$SelectCollection = New-Object System.Windows.Forms.Label
	$SelectCollection.Location = New-Object System.Drawing.Size(250,270)
	$SelectCollection.Size = New-Object System.Drawing.Size(160,20)
	$SelectCollection.TextAlign = "MiddleCenter"
	$SelectCollection.BackColor = "Blue";
	$SelectCollection.ForeColor = "White";
	$SelectCollection.Text = "Choose Session Collection:"
	$objForm.Controls.Add($SelectCollection)

# Get Session Collections
	$objCombobox = New-Object System.Windows.Forms.Combobox 
	$objCombobox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
	$objCombobox.Location = New-Object System.Drawing.Size(250,290) 
	$objCombobox.Size = New-Object System.Drawing.Size(180,100) 
	$SessionCollections = Get-RDSessionCollection -ConnectionBroker "$ConnectionBroker"

	ForEach ($Collection in $SessionCollections) {
		$CollName = $Collection.CollectionName
		[void] $objCombobox.Items.Add("$CollName")
	}

	$objCombobox.Height = 70
	$objForm.Controls.Add($objCombobox)
	
	$objCombobox_SelectedIndexChanged = {
		UpdateList
	}

	$objCombobox.add_SelectedIndexChanged($objCombobox_SelectedIndexChanged)

	$objForm.Controls.Add($objListBox)
	$objForm.Topmost = $False
	$objForm.Add_Shown({$objForm.Activate()})
	[void] $objForm.ShowDialog()
}


Function SendMessage {

	ForEach ($x in $objListBox.SelectedItems) {
		$x = $x.Split("`t")
		$RDSUser = $x[0].Trim()
		$RDSHost = $x[1].Trim()
		$RDSSessionID = $x[2].Trim()
		Send-RDUserMessage -HostServer $RDSHost -UnifiedSessionID $RDSSessionID -MessageTitle "Nachricht von $env:username" -MessageBody "$MessageText"
	}
}

Function Disconnect {

	ForEach ($x in $objListBox.SelectedItems) {
		$x = $x.Split("`t")
		$RDSUser = $x[0].Trim()
		$RDSHost = $x[1].Trim()
		$RDSSessionID = $x[2].Trim()
		Disconnect-RDUser -HostServer $RDSHost -UnifiedSessionID $RDSSessionID -Force
	}
}


Function LogOff {

	ForEach ($x in $objListBox.SelectedItems) {
		$x = $x.Split("`t")
		$RDSUser = $x[0].Trim()
		$RDSHost = $x[1].Trim()
		$RDSSessionID = $x[2].Trim()
		Invoke-RDUserLogoff  -HostServer $RDSHost -UnifiedSessionID $RDSSessionID -Force
	}
}


Function MirrorSession {
	# checking mstsc version because mirroring is not available in earlier version
	$mstsc_version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("C:\Windows\System32\mstsc.exe").Fileversion
	If ($mstsc_version -NotMatch "6.3.9600.") {
		[System.Windows.Forms.MessageBox]::Show("Your Remote Desktop Client does not support mirroring!")
		return
	}

    ForEach ($x in $objListBox.SelectedItems) {
		$x = $x.Split("`t")
		$RDSUser = $x[0].Trim()
		$RDSHost = $x[1].Trim()
		$RDSSessionID = $x[2].Trim()
		$RDSSessionActive = $x[3].Trim()

		If($RDSSessionActive -eq "STATE_DISCONNECTED") {
			[System.Windows.Forms.MessageBox]::Show("There is no active RDS session of $RDSUser and thus can't be mirrored")
			return
		}
		$params = @("/v:$RDSHost /shadow:$RDSSessionID /Control")
		$RemControl = @{
			"FilePath" = "mstsc.exe"
			"ArgumentList" = $params
			"NoNewWindow" = $true
			"Wait" = $false
		}
		Start-Process @RemControl
    }
}

Function UpdateList {

	$objListBox.BeginUpdate()
		$SelectCollection.BackColor = "Red";
	    $SelectCollection.Text = "Refreshing"
		$SessionHostCollection = $objCombobox.SelectedItem
		$LoggedOnUsers = Get-RDUserSession -ConnectionBroker "$ConnectionBroker" -CollectionName "$SessionHostCollection" | sort UserName
		[void] $objListBox.Items.Clear()
		ForEach ($user in $LoggedOnUsers) {
		$strOutput1 = $user.username
		$strOutput2 = $user.HostServer
		$strOutput3 = $user.UnifiedSessionID
		$strOutput4 = $user.SessionState
		$strOutput = "{0,-15}	{1,25}	{2,3}	{3,-35}" -f $strOutput1, $strOutput2, $strOutput3, $strOutput4
		[void] $objListBox.Items.Add($strOutput)
		}

		$objForm.Controls.Add($objListBox)
		$SelectCollection.BackColor = "Blue";
	    $SelectCollection.Text = "Choose Session Collection:"
	$objListBox.EndUpdate()

}

Function ResetButtonValues {
	$objCheckLogOff.Checked = $false
	$objCheckSendMessage.Checked = $false
	$objCheckDisconnect.Checked = $false
	$objMirrorSession.Checked = $false
}

Function ScriptActions {
#This function contains the actions to take depending on the selections set by the user

	If($objCheckSendMessage.Checked -eq $true) {
		$ApplyButton.Text = "Applying"
		$ApplyButton.BackColor = "Red";
		Start-Sleep -Seconds $waittime
		SendMessage
		ResetButtonValues
		$ApplyButton.Text = "Apply"
		$ApplyButton.BackColor = "Transparent"
	}

	If($objCheckLogOff.Checked -eq $true) {
		$ApplyButton.Text = "Applying"
		$ApplyButton.BackColor = "Red";
		Start-Sleep -Seconds $waittime
		LogOff
		ResetButtonValues
		Start-Sleep -Seconds 5
		$ApplyButton.Text = "Refreshing"
		UpdateList
		$ApplyButton.Text = "Apply"
		$ApplyButton.BackColor = "Transparent"
	}

	If($objCheckDisconnect.Checked -eq $true) {
		$ApplyButton.Text = "Applying"
		$ApplyButton.BackColor = "Red";
		Start-Sleep -Seconds $waittime
		Disconnect
		ResetButtonValues
		$ApplyButton.Text = "Refreshing"
		Start-Sleep -Seconds 3
		UpdateList
		$ApplyButton.Text = "Apply"
		$ApplyButton.BackColor = "Transparent"
	}
		
	If($objMirrorSession.Checked -eq $true) {
		$ApplyButton.Text = "Applying"
		$ApplyButton.BackColor = "Red";
		Start-Sleep -Seconds $waittime
		MirrorSession
		ResetButtonValues
		$ApplyButton.Text = "Apply"
		$ApplyButton.BackColor = "Transparent"
	}
}


$script:x =@()
$ConnectionBroker = "<Your_RDSH_goes_here>"


#Imports Remote Desktop module
Import-Module RemoteDesktop

#Runs the script
SetupForm
