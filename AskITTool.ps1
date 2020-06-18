#------------------------------------------------------------------------
# This powershell script is to be used by Ask IT to support in their
# diagnosis of user problems
#
# It is currently supported by the TechOps team with ITG
#------------------------------------------------------------------------


# This tools window interface is based on the System.Windows.Forms
# classes. These are detailed at https://msdn.microsoft.com/en-us/library/System.Windows.Forms(v=vs.110).aspx

#------------------------------------------------------------------------
#                                                             constants
#------------------------------------------------------------------------

# Version
$VERSION = 1.30
$TITLE   = "Ask IT Diagnostics Tool"

$nl = [System.Environment]::NewLine
$button_width  = 100
$button_height = 50
$button_gap_x  = 10
$button_gap_y  = 10

# TextBox for all output
$TEXTBOX     = 0
$STATUSTEXT  = 0

# Addresses for diagnostics
$CAFEVIK   = "www.cafevik.fs.fujitsu.com"
$emeiafujitsu = "emeia.fujitsu.local"
$emeia_srv_1 = "r01uksp01.r01.fujitsu.local"
$emeia_srv_2 = "r01uksp02.r01.fujitsu.local"
$emeia_srv_3 = "r01uksp05.r01.fujitsu.local"
$emeia_srv_4 = "r01uksp06.r01.fujitsu.local"
$Double799 = "europesuk003.europe.fs.fujitsu.com"
$pacexacc = "pac.exacc.com"

$TEMP_DIR = "C:\users\$env:username\AppData\Local\Temp\"
$F_IMAGE  = "Fujitsu.gif"
$FUJITSU_IMAGE_PATH = "$TEMP_DIR\$F_IMAGE"
Copy-Item -ErrorAction SilentlyContinue $F_IMAGE $FUJITSU_IMAGE_PATH

# ButtonType
$Safe_Colour    = "lightgreen"
$Warning_Colour = "yellow"
$Risky_Colour   = "red"


#------------------------------------------------------------------------
#                                                        CreateButton()
#------------------------------------------------------------------------

function CreateButton
{
	param
	(
		[string]$color,
		[scriptblock]$callback,
		[string]$text,
		[string]$tooltip
	)
	
	# create a button with and outline $color. $text, $tooltip and $callback 
	# is set - and the button is returned.
	
	$form = New-Object System.Windows.Forms.Panel 
    $form.Size = New-Object System.Drawing.Size(($button_width+10), ($button_height+10)) 
	$form.BackColor = $color
	$form.Margin = 0
	$form.BorderStyle = [System.Windows.Forms.BorderStyle]::None
			
	$button           = New-Object System.Windows.Forms.Button
	$button.Location  = New-Object System.Drawing.Size(5, 5)
	$button.Size      = New-Object System.Drawing.Size($button_width, $button_height)
	$button.Text      = $text
	$button.Margin   = 0
	$button.BackColor = "white"
	
	$tt = New-Object System.Windows.Forms.ToolTip
	$tt.SetToolTip($button, $tooltip)
	
	#$Button.BackColor = "white"
	$button.Add_Click($callback)
	$form.Controls.Add($button)
	
	return $form
}


#------------------------------------------------------------------------
#                                                       CreateRowGrid()
#------------------------------------------------------------------------

function CreateRowGrid
{
	param
	(
		[int]$columns,
		[bool]$setColumnStype = $True
	)
	
	# create a TableLayoutPanel of 1 row and $columns columns. Return
	# TableLayoutPanel
	
	$table = New-Object System.Windows.Forms.TableLayoutPanel
	$table.Dock = "Fill"
	$table.AutoSize = $True
	$table.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
	$table.RowCount = 1
	$table.Margin = 0
	$table.ColumnCount = $columns
	
	if ($setColumnStype)
	{
		foreach ($col_num in 1 .. $columns)
		{
			$cs = New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)
		
			# not sure why I need to assign this - but if I don't then the return value
			# $table has 0..4 in it - Simon Parsons XXXX
			$dummy = $table.ColumnStyles.Add($cs)
		}
	}
	return $table
}


#------------------------------------------------------------------------
#                                                         Createlabel()
#------------------------------------------------------------------------
function CreateLabel
{
	param
	(
		[string]$text
	)
	
	# create a label with text $text. Return the label
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = $text
	$label.AutoSize = $True
	$label.Dock = "Fill"
	$label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
		
	return $label
}


#------------------------------------------------------------------------
#                                                        ActionHeader()
#------------------------------------------------------------------------

function ActionHeader
{
	param
	(
		[string]$header
	)
	
	# should be called before every command it executed, given a nice look and feel to the user.
	
	$date = Get-Date
	$TEXTBOX.AppendText($nl)
	$TEXTBOX.AppendText("----------------------------------------------------------------------------$nl")
	$TEXTBOX.AppendText("$date - $header $nl")
	$TEXTBOX.AppendText("----------------------------------------------------------------------------$nl")
}

#------------------------------------------------------------------------
#                                                  StatusStartCommand()
#------------------------------------------------------------------------

function StatusStartCommand
{
	param
	(
		[string]$command
	)
	
	# Update the "status" textbox, with a colour to indicate that the
	# command is running
	
	$STATUSTEXT.Text = $command
	$STATUSTEXT.BackColor = "yellow"
	$STATUSTEXT.Update()
}


#------------------------------------------------------------------------
#                                                   StatusStopCommand()
#------------------------------------------------------------------------

function StatusStopCommand
{
	param
	(
		[string]$command
	)
	
	# Update the "status" textbox, with a colour to indicate that the
	# command is complete
		
	$STATUSTEXT.Text = $command
	$STATUSTEXT.BackColor = "white"
	$STATUSTEXT.Update()
}


#------------------------------------------------------------------------
#                                                          RunCommand()
#------------------------------------------------------------------------

function RunCommand
{
	param
	(
		[string]$command,
        [string]$logfile = $null
	)
	
	# Run a command - and update the status and central textboxes
	# accordingly
	
	$s = "RUNNING: $command  ... "
	$TEXTBOX.AppendText($nl + $s)
	StatusStartCommand($s)
	
	$output = Invoke-Expression $command | out-string
	
	
	$status = "SUCCEEDED"
	if (-Not $?)
	{
		$status = "FAILED - Exit Status: $LASTEXITCODE$nl"
	}
	
	$TEXTBOX.AppendText($status + $nl)
	$TEXTBOX.AppendText($output + $nl)
    if ($logfile -ne $null){
        $output >> $logfile
    }
	StatusStopCommand($s + $status)
	
	return $?
}


#------------------------------------------------------------------------
#                                                  NetworkDiagnostics()
#------------------------------------------------------------------------

function NetworkDiagnostics
{
	# Perform general network diagnostics
	
	ActionHeader("Running Complete Network Diagnostics")
	$TEXTBOX.AppendText("Collecting information$nl")
	$TEXTBOX.AppendText("Please be patient, this can take a few minutes$nl$nl")

	RunCommand "ipconfig /all"
	RunCommand "ping -n 2 $pacexacc"
	RunCommand "tracert -w 200 -d $pacexacc"
	RunCommand "ping -n 2 $emeia_srv_1"
	RunCommand "tracert -w 200 -d $emeia_srv_1"
	RunCommand "ping -n 2 $emeia_srv_2"
	RunCommand "tracert -w 200 -d $emeia_srv_2"
	RunCommand "ping -n 2 $emeia_srv_3"
	RunCommand "tracert -w 200 -d $emeia_srv_3"
	RunCommand "ping -n 2 $emeia_srv_4"
	RunCommand "tracert -w 200 -d $emeia_srv_4"
	RunCommand "ping -n 2 $Double799"
    RunCommand "reg query `"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings`""

	$web = & New-Object System.Net.WebClient
	$web = $web.DownloadFile("http://pac.exacc.com","C:\users\$env:username\AppData\Local\Temp\pac.txt")
	RunCommand "type `"C:\users\$env:username\AppData\Local\Temp\pac.txt`""
	RunCommand "route print"
    RunCommand "type `"C:\Windows\System32\drivers\etc\hosts`""
    RunCommand "nslookup vpntest.exacc.com"

	$TEXTBOX.AppendText("$nl Collection Complete$nl$nl")
}


#------------------------------------------------------------------------
#                                                            GPUpdate()
#------------------------------------------------------------------------
function GPUpdate
{
     # For some reason - GPUpdate does not like to be run through
	 # RunCommand/Invoke-Expression - so we use a subshell
	 
	 RunCommand("start -wait powershell 'gpupdate /force'")
}


#------------------------------------------------------------------------
#                                                    EmailWithOutlook()
#------------------------------------------------------------------------

function EmailWithOutlook
{
	param
	(
		[string]$text,
		[string]$attachment
	)

	# Open an outlook window for sending mail, with the body text and
	# attachment as neccessary
	
	$outlook = New-Object -com Outlook.Application
	
	# CreateItem is described here
	# https://msdn.microsoft.com/en-us/library/office/ff869635.aspx
	$mail = $outlook.CreateItem(0)
	
	# MailItem Members are described
	# https://msdn.microsoft.com/EN-US/library/office/ff861252.aspx
	$mail.Subject = "$TITLE - User:$env:username Device:$env:COMPUTERNAME"
	$mail.BodyFormat = 1       # olFormatPlain
	$mail.Body    = $text
	
	if (-Not [string]::IsNullOrEmpty($attachment))
	{
		$mail.Attachments.Add($attachment)
	}
	
	$TEXTBOX.AppendText("$nl An Outlook mail window will appear - can you please fill the 'To' address as requested and then send the email")
	
	$mail.Display()
}


#------------------------------------------------------------------------
#                                                     EmailTextWindow()
#------------------------------------------------------------------------

function EmailTextWindow
{
	# Email the contents of the main text box
	
	EmailWithOutlook $TEXTBOX.Text
}


#------------------------------------------------------------------------
#                                                     RemoteAssistant()
#------------------------------------------------------------------------

function RemoteAssistant
{
	# Run up the remote assistant
	
	ActionHeader("Runnign msra /email")
	RunCommand("msra /email")
	$TEXTBOX.AppendText("$nlAn Outlook Window will have opened on your computer, this will allow you ask for assistance$nl")
	$STATUSTEXT.Text = "An Outlook Window will have opened on your computer"
}

#------------------------------------------------------------------------
#                                                            GPResult()
#------------------------------------------------------------------------

function GPResult
{
	# Run gpresult and then create an email with the results as an
	# attachment
	
	ActionHeader("Running gpresult")
	$TEXTBOX.AppendText("This will take some time to run, once it has finished an email window will appear.$nl")
	$report_file = "$TEMP_DIR\GPReport.html"
	$cmd = "gpresult /f /h $report_file"
    TextPopup "Please Note" "This requires a temporary change to User Account Control - please click Yes when the prompt appears."

    $proc = Start-Process powershell -ArgumentList $cmd  -verb runAs -passthru
    do {start-sleep -Milliseconds 500}
    until ($proc.HasExited)
    
	EmailWithOutlook "Find attached the output of $cmd" $report_file
}


#------------------------------------------------------------------------
#                                                            DetectLync()
#------------------------------------------------------------------------

function DetectLync
{
	# Run Lync checks, retrieve Lync-related event logs then create an email with the results as an
	# attachment
	
	ActionHeader("Detecting Lync")
    $URL = "http://pac.exacc.com"
    $Desktop = "$env:USERPROFILE\Desktop"
    $file = "$destfolder\pac.txt"
    $Username = "$env:USERNAME"
    $LyncTracingFolder = "$env:USERPROFILE\AppData\Local\Microsoft\Office\15.0\Lync\Tracing"
    $destfolder = "$desktop\Diagnostic-results_$username"
    $file = "$destfolder\pac.txt"
    $zipdest = "$Desktop\diag-results_030716.zip"
    $logsf = "C:\Windows\System32\winevt\Logs"
    $appl = "$destfolder\Application.evtx"
    $sysl = "$destfolder\System.evtx"
    $oal = "$destfolder\OAlerts.evtx"

    md -Force $destfolder
    remove-item $file
    remove-item $appl
    remove-item $sysl
    remove-item $oal

    # copy SIP traces
    Copy-Item $LyncTracingFolder"\Lync-UccApi*.*" $destfolder
    
    # copy system logs
    RunCommand("wevtutil.exe epl Application $appl /ow")
    RunCommand("wevtutil.exe epl System $sysl /ow")
    RunCommand("wevtutil.exe epl OAlerts $oal /ow")

    # location
    "Location" >> $file
    "##################################################" >> $file
    $location=$env:Location
    $location | out-file -filepath $file -append
    $Userdomain=$env:Userdomain
    $Userdomain | out-file -filepath $file -append
    "##################################################" >> $file
    
    # ipconfig /all
    "ipconfig /all" >> $file
    "##################################################" >> $file
    RunCommand "ipconfig /all" $file
    "##################################################" >> $file
    
    # ping exacc.com etc.
    "ping exacc.com" >> $file
    "##################################################" >> $file
    RunCommand "ping -n 2 pac.exacc.com" $file
    "##################################################" >> $file
    "ping default proxy" >> $file
    "##################################################" >> $file
    RunCommand "ping -n 2 default.proxy.fs.fujitsu.com" $file
    "##################################################" >> $file
    "ping r01uklyfepool01.r01.fujitsu.local" >> $file
    "##################################################" >> $file
    RunCommand "ping -n 2 r01uklyfepool01.r01.fujitsu.local" $file 
    "##################################################" >> $file
    "nslookup lync.global.fujitsu.com" >> $file
    "##################################################" >> $file
    RunCommand "nslookup lync.global.fujitsu.com" $file
    "##################################################" >> $file
    "ping r01uklyfepool02.r01.fujitsu.local" >> $file
    "##################################################" >> $file
    Runcommand "ping -n 2 r01uklyfepool02.r01.fujitsu.local" $file
    "##################################################" >> $file
    
    # IE lan settings list
    "List LAN Internet Settings" >> $file
    "##################################################" >> $file
    $internetsettings = Get-Item -Path "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\"     #proxyserver status
    $internetsettings.GetValue("AutoConfigProxy") | out-file -filepath $file -append
	if ($internetsettings.GetValue("ProxyEnable") -eq $null){
        "ProxyEnable key does not exist" >> $file
    }
    else {
    	if ($internetsettings.GetValue("ProxyEnable") -ne 0){
 	  	    "Proxy is enabled" >> $file
        }
	    else {
		    "Proxy is off" >> $file
	    }
    }
	if ($internetsettings.GetValue("ProxyServer") -eq $null){
        "ProxyServer key does not exist" >> $file
    }
    else {
        "ProxyServer:" >> $file
        $internetsettings.GetValue("ProxyServer")  | out-file -filepath $file -append
    }
	if ($internetsettings.GetValue("ProxyOverride") -eq $null){
        "ProxyOverride key does not exist" >> $file
	    "No proxy exclusions." >> $file
    }
    else {
        $internetsettings.GetValue("ProxyOverride")  | out-file -filepath $file -append
	    if ($internetsettings.GetValue("ProxyOverride") -ne 0){ 
		    "Proxy exclusions are:" >> $file
		    $internetsettings.GetValue("ProxyOverride")  | out-file -filepath $file -append
        }
	    else {
		    "No proxy exclusions." >> $file
	    }		
    }
	$pacurl = $internetsettings.GetValue("AutoConfigURL")
	If ($pacurl -eq $null) { 
		"No AutoConfig PAC URL" >> $file
	}
    else {
        "AutoConfig PAC URL:" >> $file
		$pacurl  | out-file -filepath $file -append
        $destination = "$destfolder\pacfile.txt"
        
        $WebClient = New-Object System.Net.WebClient
        $WebClient.DownloadFile($pacurl,$destination)        
    }

    $targetGroup = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections\" ## DefaultConnectionSettings"
    $target = $targetGroup.DefaultConnectionSettings
    $output = ""
    
    for ($k = 0; $k -lt $target.length; $k++){
        $output = $output + "{0:X2}" -f $target[$k] + ","
    }
    $near = ($output.substring(0,26))
	$far = ($near.substring($near.length - 1, 1))             
             
    switch ($far){
        "9" { "Automatically detect settings is enabled" >> $file }
        "3" { "Use a proxy server for your LAN is enabled" >> $file }
        "B" { "Both proxy and autodetect settings are  enabled" >> $file }
        "5" { "Use automatic configuration script is enabled" >> $file }
        "D" { "Automatically detect settings and Use automatic configuration script are enabled" >> $file }
        "7" { "Use a proxy server for your LAN and Use automatic configuration script are enabled" >> $file }
        "F" { "all the three are enabled" >> $file }
        "1" { "none of three are enabled" >> $file }
    }

    if ($pacurl -ne $null) { 
	    "##################################################" >> $file
        "PAC file content" >> $file
        "##################################################" >> $file
        cat $destination >> $file
	}

    "##################################################" >> $file
    "OS Version" >> $file
    "##################################################" >> $file
		
    $OSVersion = [environment]::OSVersion.Version
    $versionstring = [string]$OSVersion
    $versioncheck = $versionstring.substring(0,3)
  
    # convert numbers into names
    switch ($versioncheck) {
        "10." { "Windows 10" >> $file }
        "6.3" { "Windows 8.1" >> $file }
        "6.2" { "Windows 8.0" >> $file }
        "6.1" { "Windows 7" >> $file }
        "6.0" { "Windows Vista" >> $file }
        "5.2" { "Windows 2003" >> $file }
        "5.1" { "Windows XP" >> $file }
        "5.0" { "Windows 2000" >> $file }
        "4.0" { "NT 4.0" >> $file }
        default { "Higher than Windows 10" >> $file }
    }
    "##################################################" >> $file
    # hosts
    "c:\windows\system32\drivers\etc\hosts" >> $file
    "##################################################" >> $file

    cat "c:\windows\system32\drivers\etc\hosts" >> $file

    "##################################################" >> $file
           
	$TEXTBOX.AppendText("$nl A folder called 'Diagnostic-results_<username>' has been created on the desktop of your computer.$nl Ask IT might ask you to copy this to a network location for further analysis.$nl")
}

#------------------------------------------------------------------------
#                                                            WindowsUpdateLog()
#------------------------------------------------------------------------

function WindowsUpdateLog
{
	# Retrieve WindowsUpdate.log then create an email with the log as an
	# attachment
	
	ActionHeader("Retrieving WindowsUpdate.log")
   	$TEXTBOX.AppendText("$nlAn Outlook Window will have opened on your computer, this will allow you to email the log file$nl")
	$STATUSTEXT.Text = "An Outlook Window will have opened on your computer"
	$log_file = "$env:WINDIR\WindowsUpdate.log"
	EmailWithOutlook "Find attached the Windows Update log" $log_file
}

#------------------------------------------------------------------------
#                                                          DeleteTree()
#------------------------------------------------------------------------

function DeleteTree
{
	param
	(
		[string]$dir
	)
	
	# Delete a directory tree
	
	StatusStartCommand("Deleting files from $dir")
	$dump = remove-item -path "$dir\*" -recurse -ErrorAction SilentlyContinue
	StatusStopCommand("Files deleted from $dir")
	
	$TEXTBOX.AppendText($nl + "Deleted files from $dir$nl")
}

#------------------------------------------------------------------------
#                                                DeleteTemporaryFiles()
#------------------------------------------------------------------------

function DeleteTemporaryFiles
{
	# Delete common temporary files
	
	ActionHeader("Deleting Temporary Files")
	DeleteTree("c:\Windows\Temp\")
	DeleteTree("C:\Users\$env:username\AppData\Local\Temp")
}

#------------------------------------------------------------------------
#                                                      TextEntryPopup()
#------------------------------------------------------------------------

function TextEntryPopup
{
	param
	(
		[string]$title, [string]$prompt, [string]$prepop, $masked
	)

	# Create a one line text entry box, and return the data entered
	
	$okButton          = New-Object System.Windows.Forms.Button
	$okButton.Text     = "OK"
	$okButton.Size     = New-Object System.Drawing.Size(75,25)
	$okButton.Location = New-Object System.Drawing.Size(100,70)
	$okButton.Add_Click({$form.Tag = $textBox.Text; $form.Close()})
	
	$cancelButton      = New-Object System.Windows.Forms.Button
	$cancelButton.Text = "Cancel"
	$cancelButton.Size = New-Object System.Drawing.Size(75,25)
	$cancelButton.Location = New-Object System.Drawing.Size(10,70)
	$cancelButton.Add_Click({$form.Tag = $null; $form.Close()})
	
	$label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Size(10,10) 
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.AutoSize = $true
    $label.Text = $prompt
	
    if ($masked -eq $false){
	    $textbox = New-Object System.Windows.Forms.TextBox
    }
    else {
        $textbox = New-Object System.Windows.Forms.MaskedTextBox
        $textbox.PasswordChar = '*'
	}
	$textBox.Location = New-Object System.Drawing.Size(10,40) 
    $textBox.Size = New-Object System.Drawing.Size(170,20)
#	$textBox.AcceptsReturn = $True
	$textBox.Multiline     = $False
	$textBox.Text = $prepop
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = $title
	$form.Topmost = $True
	$form.AcceptButton = $okButton
	$form.CancelButton = $cancelButton
	$form.StartPosition = "CenterScreen"
	$form.ShowInTaskbar = $True
	$form.Size = New-Object System.Drawing.Size(250,200)
    $form.Tag = $null
	$form.Controls.Add($label)
	$form.Controls.Add($textBox)
	$form.Controls.Add($okButton)
	$form.Controls.Add($cancelButton)
	
	$form.Add_Shown($form.Activate())
	[void]$form.ShowDialog()
	
	return $form.Tag
}


#------------------------------------------------------------------------
#                                                      TextPopup()
#------------------------------------------------------------------------

function TextPopup
{
	param
	(
		[string]$title,[string]$prompt
	)

	# Create a one line text box
	
	$okButton          = New-Object System.Windows.Forms.Button
	$okButton.Text     = "OK"
	$okButton.Size     = New-Object System.Drawing.Size(75,25)
	$okButton.Location = New-Object System.Drawing.Size(100,70)
	$okButton.Add_Click({$form.Tag = $textBox.Text; $form.Close()})
	
	$label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Size(10,10) 
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.AutoSize = $true
    $label.Text = $prompt
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = $title
	$form.Topmost = $True
	$form.AcceptButton = $okButton
	$form.StartPosition = "CenterScreen"
	$form.ShowInTaskbar = $True
	$form.Size = New-Object System.Drawing.Size(800,150)
	$form.Tag = $null
	$form.Controls.Add($label)
	$form.Controls.Add($okButton)
	
	$form.Add_Shown($form.Activate())
	[void]$form.ShowDialog()
	
	return $form.Tag
}


#------------------------------------------------------------------------
#                                               ProblemsStepsRecorder()
#------------------------------------------------------------------------

function ProblemStepsRecorder
{
	# Invoke the problem step recorder
	
	ActionHeader("Problem Steps Recorder")
	$TEXTBOX.AppendText("@Please record your issue (using the 'Start Record' and 'Stop Record' buttons)
Once you have recorded your problem, click the triangle next to the blue '?' and select 'Send to email recipient' 
Please  email the file to Ask IT Service Desk
")

	$zip_file = "$TEMP_DIR\7799record.zip"
	RunCommand("psr.exe /output $zip_file")
}


#------------------------------------------------------------------------
#                                               ShowDeleteCredentials()
#------------------------------------------------------------------------

function ShowDeleteCredential
{
	# Allow the user the ability to delete credentials
	
	ActionHeader("Invoking Credential Manager")
	$TEXTBOX.AppendText("Please remove all entries that exist, and then close the window")

	RunCommand("rundll32.exe keymgr.dll, KRShowKeyMgr")
}


#------------------------------------------------------------------------
#                                                             ResetIE()
#------------------------------------------------------------------------

function ResetIE
{
	# Reset IE settings to the defaults.
	
	ActionHeader("Resetting IE")
	RunCommand("rundll32 inetcpl.cpl ResetIEtoDefaults")
}


#------------------------------------------------------------------------
#                                                            FlushDNS()
#------------------------------------------------------------------------

function FlushDNS
{
	# Flush out the local DNS
	
	ActionHeader("Flushing The DNS Cache")
	RunCommand("ipconfig /flushdns")
}


#------------------------------------------------------------------------
#                                                NetworkConfiguration()
#------------------------------------------------------------------------

function NetworkConfiguration
{
	# Show local network configuration
	
	ActionHeader("Displaying Network Configuration")
	RunCommand("ipconfig /all")
}


#------------------------------------------------------------------------
#                                                      ReleaseRenewIP()
#------------------------------------------------------------------------

function ReleaseRenewIP
{
	# Release and renew the DHCP lease
	
	ActionHeader("Releasing and Renewing The IP Address")
	RunCommand("ipconfig /release")
	RunCommand("ipconfig /renew")
}


#------------------------------------------------------------------------
#                                                  PingSpecificServer()
#------------------------------------------------------------------------

function PingSpecificServer
{
	# Ping / tracert to a user specified server
	
	ActionHeader("Ping Specific Server")
	
	$hostname = TextEntryPopup "Host" "Enter the name/IP address of the host" "" $false

	if (-Not [string]::IsNullOrEmpty($host))
	{
		$TEXTBOX.AppendText("Host is $hostname")
		RunCommand("ping -n 2 $hostname")
		RunCommand("tracert -w 200 $hostname")
	}
	else
	{
		$TEXTBOX.AppendText("No host specified $nl")
	}
}


#------------------------------------------------------------------------
#                                                          StoreCredentials()
#------------------------------------------------------------------------

function UserPass
{
    $result = $false
    $global:username = TextEntryPopup "Username" "Please enter your Username" $env:username $false
    If($username) {
        $global:pass = TextEntryPopup "Password" "Please enter your Password" "" $true
        If($pass) {
            $result = $true
        }
    }
    $result
}

function Add-G03Credential ($target)
{
    $status = "FAILED"
	$TEXTBOX.AppendText( "Adding your credentials for: $target$nl")
    
   	[string]$result = cmdkey /add:$target /user:G03\$username /pass:$pass
   
	If($result -match "The command line parameters are incorrect")
	{
		$TEXTBOX.AppendText( "Failed to add Windows credentials to the Windows vault.$nl")
	}
	ElseIf($result -match "CMDKEY: Credential added successfully")
	{
		$TEXTBOX.AppendText( "Credentials added successfully.$nl")
        $status = "SUCCEEDED"
        write-host $status
	}
    write-host $status
    $status
}

function Add-EUROPECredential ($target)
{
    $status = "FAILED"
	$TEXTBOX.AppendText( "Adding your credentials for: $target$nl")
    
   	[string]$result = cmdkey /add:$target /user:EUROPE\$username /pass:$pass
   
	If($result -match "The command line parameters are incorrect")
	{
		$TEXTBOX.AppendText( "Failed to add Windows credentials to the Windows vault.$nl")
	}
	ElseIf($result -match "CMDKEY: Credential added successfully")
	{
		$TEXTBOX.AppendText( "Credentials added successfully.$nl")
        $status = "SUCCEEDED"
        write-host $status
	}
    write-host $status
    $status
}

function StoreCredentials
{
	# Store G03 Credentials
	$status = $env:userdomain
	if ($status -eq "EUROPE") {
	
		ActionHeader("Storing G03 Credentials")
	
		$TEXTBOX.AppendText( "If you enter your G03 Username and Password, this tool will cache your credentials for the following sites:$nl")
		$TEXTBOX.AppendText( "- *.fujitsu.local - All sites ending with .fujitsu.local$nl")
		$TEXTBOX.AppendText( "- adfs1.global.fujitsu.com - ADFS$nl")
		$TEXTBOX.AppendText( "Note: If you wish to amend any entries, please go to Control Panel -> Credential Manager$nl")
	
		if (UserPass -eq $true) {
			$s = "RUNNING: add G03 credentials  ... "
			StatusStartCommand($s)
			$status = "SUCCEEDED"
			if ((Add-G03Credential *.fujitsu.local) -eq "FAILED"){
				$status = "FAILED"
				write-host $status
			}
			if ((Add-G03Credential adfs1.global.fujitsu.com) -eq "FAILED"){
				$status = "FAILED"
				write-host $status
			}
			write-host $status
			StatusStopCommand($s + $status)
		}
		else {
			$TEXTBOX.AppendText( "Update aborted - please click button again to retry$nl")
		}
	}
	else {
		ActionHeader("Storing EUROPE Credentials")
	
		$TEXTBOX.AppendText( "If you enter your EUROPE Username and Password, this tool will cache your credentials for the following sites:$nl")
		$TEXTBOX.AppendText( "- www.cafevik.fs.fujitsu.com$nl")
		$TEXTBOX.AppendText( "- sites.cafevik.fs.fujitsu.com$nl")
		$TEXTBOX.AppendText( "- portals.cafevik.fs.fujitsu.com$nl")
		$TEXTBOX.AppendText( "- mysites.cafevik.fs.fujitsu.com$nl")
		$TEXTBOX.AppendText( "- adfs.uk.fujitsu.com$nl")
		$TEXTBOX.AppendText( "- intranet.ts.fujitsu.com$nl")
		$TEXTBOX.AppendText( "- extranet.uk.fujitsu.com$nl")
		$TEXTBOX.AppendText( "Note: If you wish to amend any entries, please go to Control Panel -> Credential Manager$nl")
	
		if (UserPass -eq $true) {
			$s = "RUNNING: add EUROPE credentials  ... "
			StatusStartCommand($s)
			$status = "SUCCEEDED"
			if ((Add-EUROPECredential www.cafevik.fs.fujitsu.com) -eq "FAILED"){
				$status = "FAILED"
				write-host $status
			}
			if ((Add-EUROPECredential sites.cafevik.fs.fujitsu.com) -eq "FAILED"){
				$status = "FAILED"
				write-host $status
			}
			if ((Add-EUROPECredential portals.cafevik.fs.fujitsu.com) -eq "FAILED"){
				$status = "FAILED"
				write-host $status
			}
			if ((Add-EUROPECredential mysites.cafevik.fs.fujitsu.com) -eq "FAILED"){
				$status = "FAILED"
				write-host $status
			}
	        if ((Add-EUROPECredential adfs.uk.fujitsu.com) -eq "FAILED"){
	            $status = "FAILED"
	            write-host $status
	        }
	        if ((Add-EUROPECredential intranet.ts.fujitsu.com) -eq "FAILED"){
	            $status = "FAILED"
	            write-host $status
	        }
	        if ((Add-EUROPECredential extranet.uk.fujitsu.com) -eq "FAILED"){
	            $status = "FAILED"
	            write-host $status
	        }
			write-host $status
			StatusStopCommand($s + $status)
		}
		else {
			$TEXTBOX.AppendText( "Update aborted - please click button again to retry$nl")
		}
	}
}


#------------------------------------------------------------------------
#                                               OSTFileRepair()
#------------------------------------------------------------------------


function RenameOST()
{
    $random = 0
    $ostfile = 0
    $random = get-random -Minimum 1 -Maximum 99
    $ostfile = Get-ChildItem "C:\Users\$env:username\Documents\EMAIL" -Recurse | where { $_.Extension -eq ".ost"} | Rename-Item -NewName {  $_.Name -replace '\.ost',('.old' + $random) }
}

function OSTFileRepair
{
    #Information for user 
   	ActionHeader("AskIT Tool will now close Outlook and Skype for Business")
    Import-Module ActiveDirectory
    #Stop Lync and Outlook process
    $OutlookLync = Get-Process Outlook, lync -ErrorAction SilentlyContinue
    if ($OutlookLync) {
        #Grace attempt
        $OutlookLync.CloseMainWindow()
        #Kill after 5 sec
        Sleep 5
        if (!$OutlookLync.HasExited) {
            $OutlookLync | Stop-Process -Force
        }
    }

    #Rename ost file
    RenameOST
    #Information for user 
    ActionHeader("Outlook .ost file has been rebuilt. Please restart Outlook and Skype for Business")
}

#------------------------------------------------------------------------
#                                               RemoveSIPFile()
#------------------------------------------------------------------------

function RemoveSIPFile
{
    #Skype process termination
    ActionHeader("Terminating Skype for Business")
    Stop-Process -processname lync 
    Start-Sleep 2
    ActionHeader("Skype for Business Process terminated")
    #Skype Files Location
    $SkypeFile = Get-ChildItem -Path C:\Users\$env:username\AppData\Local\Microsoft\Office\ sip_* -Recurse | ?{ $_.PSIsContainer }
    ActionHeader("Removing SIP files")
    #Remove Item code
    remove-item $SkypeFile.fullname -Recurse 
    ActionHeader("SIP files removed")
    #Start Skype for Business
    ActionHeader("Starting Skype for Business")
    if (test-path "C:\Program Files (x86)\Microsoft Office\*\lync.exe"){
        invoke-item "C:\Program Files (x86)\Microsoft Office\*\lync.exe" # Different versions of Lync
    }
    else
    {
        if (test-path "C:\Program Files\Microsoft Office\*\lync.exe"){
            invoke-item "C:\Program Files\Microsoft Office\*\lync.exe" # Different versions of Lync
        }
    }
    ActionHeader("Done")
}

#------------------------------------------------------------------------
#                                               StopiPass()
#------------------------------------------------------------------------

function StopiPass
{
    ActionHeader("Stopping and restarting iPass")
 	$cmd = "C:\Program Files (x86)\AskITTool\iPassStopStart.ps1"
    TextPopup "Please Note" "This requires a temporary change to User Account Control - please click Yes when the prompt appears."
    set-executionpolicy -scope process -executionpolicy remotesigned -force
    $proc = Start-Process powershell -ArgumentList "-executionpolicy bypass -file `"$cmd`"" -verb runAs -passthru
    do {start-sleep -Milliseconds 500}
    until ($proc.HasExited)
    ActionHeader("Done")
}


#------------------------------------------------------------------------
#                                                         OpenFAQInIE()
#------------------------------------------------------------------------

function OpenFAQInIE
{
	# Open the Ask IT FAQ in IE
	ActionHeader("Opening the Ask IT FAQ In IE")
	#invoke-item "\\europevuk396.europe.fs.fujitsu.com\7799DocumentStore\links.html"
    start "https://fujitsu.service-now.com/ask_it/AskIT_self_help.do"
}


#------------------------------------------------------------------------
#                                                              let's go
#------------------------------------------------------------------------

# load assemblies to allow us to access the graphics interface.
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Net.Mail")


# Main Form
$objForm      = New-Object System.Windows.Forms.Form 
$objForm.Text = "$TITLE - Version $VERSION"
#$objForm.Size = New-Object System.Drawing.Size(720,700) 
$objForm.Size = New-Object System.Drawing.Size(800,700) 
$objForm.Opacity = 1.0
$objForm.BackColor = "white"
$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})


# The main layout is a TableLayoutPanel of 1 column and 6 rows.
#
# ROW 0 - introduction to the tool
# ROW 1 - display of user name and host machine
# ROW 2 - First line of green buttons
# ROW 3 - Second line of green buttons
# ROW 4 - First Row Of Yellow Buttons
# ROW 5 - TextBox for the output of the commands
# ROW 6 - Status Bar
# ROW 7 - Red buttons	
	
$mainGrid = New-Object System.Windows.Forms.TableLayoutPanel
$mainGrid.Dock = "Fill"
$mainGrid.AutoSize = $True
$mainGrid.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$mainGrid.RowCount    = 8
$mainGrid.ColumnCount = 1
$mainGrid.Margin      = 0
$objForm.Controls.Add($mainGrid)

$fujitsu_image = [System.Drawing.Image]::Fromfile($FUJITSU_IMAGE_PATH)

# ROW 0
# Simon Parsons - you can do better than +5 XXX
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 60)
$dummy = $mainGrid.RowStyles.Add($rs)
# ROW 1
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)
$dummy = $mainGrid.RowStyles.Add($rs)
# ROW 2
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 60)
$dummy = $mainGrid.RowStyles.Add($rs)
# ROW 3
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 60)
$dummy = $mainGrid.RowStyles.Add($rs)
# ROW 4
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 60)
$dummy = $mainGrid.RowStyles.Add($rs)
# ROW 5
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)
$dummy = $mainGrid.RowStyles.Add($rs)
# ROW 6
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)
$dummy = $mainGrid.RowStyles.Add($rs)
# ROW 7
$rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 60)
$dummy = $mainGrid.RowStyles.Add($rs)


# ROW 0 - Introduction to tool
$r0 = CreateRowGrid 2 $False
$mainGrid.Controls.Add($r0, 0, 0)
$l00 = CreateLabel ""
$pb00 = New-Object Windows.Forms.PictureBox
$pb00.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$pb00.Width = 120
$pb00.Height = 60
$pb00.Image  = $fujitsu_image
$r0.Controls.Add($pb00, 0, 0)
$l01 = CreateLabel "$TITLE - is brought to you in conjunction with ITG$nl This tool should be used in conjunction with a member of the Ask IT Service Desk"
$r0.Controls.Add($l01, 1, 0)


# ROW 1 - Display the username and device
$r1 = CreateRowGrid 2
$mainGrid.Controls.Add($r1, 0, 1)
$l01 = CreateLabel ("Username: " + $env:username)
$r1.Controls.Add($l01, 0, 0)
$l02 = CreateLabel ("Machine Name: " + $env:COMPUTERNAME)
$r1.Controls.Add($l02, 2, 0)

# ROW 2 - First line of green buttons
$r2 = CreateRowGrid 7
$mainGrid.Controls.Add($r2, 0, 2)

$b20 = CreateButton $Safe_Colour {OpenFAQInIE}  `
	"Go To Ask IT Self-Service Portal" `
    "Opens a page within Internet Explorer displaying Ask IT Self-Service Portal"
$r2.Controls.Add($b20, 0, 0)

$b21 = CreateButton $Safe_Colour {RemoteAssistant}        `
	"Remote Help Invitation"       `
	"Invoke the Microsoft Remote Assistant"
$r2.Controls.Add($b21, 1, 0)

$b22 = CreateButton $Safe_Colour {ProblemStepsRecorder}     `
	"Record your issue"        `
	"This will invoke 'psr' (Problem Steps Recorder) to allow you to record your issue, and send it to Ask IT"
$r2.Controls.Add($b22, 2, 0)

$b23 = CreateButton $Safe_Colour {NetworkConfiguration} `
	"Display Network configuration"    `
	"Display the network configuration of the local computer"
$r2.Controls.Add($b23, 3, 0)

$b24 = CreateButton $Safe_Colour {GPResult} `
	"Group Policy Result"       `
	"Gathers Group Policy information from the local PC (this takes time) and then creates an email to forward the report to an email recipient"
$r2.Controls.Add($b24, 4, 0)
	
$b25 = CreateButton $Safe_Colour {PingSpecificServer}    `
	"Ping / Tracert to a specific server" `
	"This will prompt for a servername or IP address and ping/tracert to it"
$r2.Controls.Add($b25, 5, 0)

$b26 = CreateButton $Safe_Colour {NetworkDiagnostics} `
	"Full Network Diagnostics"          `
	"This command performs various network diagnostics, tracert pings etc. and then
	 adds the output to the text box"
$r2.Controls.Add($b26, 6, 0)

	
# ROW 3 - Second line of green buttons
$r3 = CreateRowGrid 7
$mainGrid.Controls.Add($r3, 0, 3)

$b30 = CreateButton $Warning_Colour {RemoveSIPFile} `
	"Delete SIP File"             `
	"Removes the SIP files as used by Skype for Business"
$r3.Controls.Add($b30, 0, 0)


$b31 = CreateButton $Warning_Colour {StopiPass}       `
	"Stop iPass"    `
	"This will stop the iPass process"
$r3.Controls.Add($b31, 1, 0)

# B32 - Missing

# B33 - Missing

# B34 - Missing

$b35 = CreateButton $Safe_Colour {DetectLync} `
	"Detect Lync"          `
	"This command retrieves the event files and other data related to Lync"
$r3.Controls.Add($b35, 5, 0)

$b36 = CreateButton $Safe_Colour {WindowsUpdateLog} `
	"Windows Update Log"          `
	"This command retrieves the log file produced by Windows Update"
$r3.Controls.Add($b36, 6, 0)



# ROW 4 - Yellow Line
$r4 = CreateRowGrid 7
$mainGrid.Controls.Add($r4, 0, 4)

$b40 = CreateButton $Warning_Colour {OSTFileRepair} `
	"Repair OST (Outlook) File" `
	"This repairs the Outlook data file"
$r4.Controls.Add($b40, 0, 0)

$b41 = CreateButton $Warning_Colour {DeleteTemporaryFiles} `
	"Delete Temp Files" `
	"Deletes temporary files, freeing up local disk space"
$r4.Controls.Add($b41, 1, 0)

$b42 = CreateButton $Warning_Colour {ShowDeleteCredential} `
	"Show/Delete Credentials"      `
	"This runs the credential manager, and allows the user to remove cached credentials"
$r4.Controls.Add($b42, 2, 0)

$b43 = CreateButton $Warning_Colour {GPUpdate}  `
	"Group Policy Update"        `
	"Forces an update of the Group Policy settings."
$r4.Controls.Add($b43, 3, 0)

$b44 = CreateButton $Warning_Colour {FlushDNS} `
	"Flush DNS Cache"                 `
	"This flushed the DNS cache - this is low risk, as the data held there is transient"
$r4.Controls.Add($b44, 4, 0)
	
$b45 = CreateButton $Warning_Colour  {ReleaseRenewIP}  `
	"Release/Renew IP Address"          `
	"Release the IP address and renew it - this means you will lose IP connectivity for a period"
$r4.Controls.Add($b45, 5, 0)

$b46 = CreateButton $Warning_Colour  {StoreCredentials}  `
	"Store AD Credentials"          `
	"Store AD credentials to allow automatic login"
$r4.Controls.Add($b46, 6, 0)



# ROW 5
$tb5 = New-Object System.Windows.Forms.TextBox
$tb5.ReadOnly = $True
$tb5.Dock = "Fill"
$tb5.AutoSize = $True
$tb5.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor `
              [System.Windows.Forms.AnchorStyles]::Top    -bor `
			  [System.Windows.Forms.AnchorStyles]::Left   -bor `
              [System.Windows.Forms.AnchorStyles]::Right
$tb5.Size = New-Object System.Drawing.Size(660, 330) 
$tb5.Text       = "Welcome to the $TITLE$nl"
$tb5.Multiline  = $True
$tb5.Font       = New-Object System.Drawing.Font("Courier New", "8")
$tb5.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$tb5.ForeColor  = "White"
$tb5.BackColor  = "Black" 
$mainGrid.Controls.Add($tb5)
$TEXTBOX = $tb5


# ROW 6 - Status Label
$r6 = CreateRowGrid 2 $False
$mainGrid.Controls.Add($r6, 0, 6)
$l60 = CreateLabel "Status:"
$r6.Controls.Add($l60, 0, 0)
$tb61 = New-Object System.Windows.Forms.TextBox
$tb61.ReadOnly = $True
$tb61.Dock = "Fill"
$tb61.AutoSize = $True
$tb61.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor `
              [System.Windows.Forms.AnchorStyles]::Top    -bor `
			  [System.Windows.Forms.AnchorStyles]::Left   -bor `
              [System.Windows.Forms.AnchorStyles]::Right
$r6.Controls.Add($tb61, 1, 0)
$STATUSTEXT = $tb61

# ROW 7 - Risky Buttons
$r7 = CreateRowGrid 7
$r7.Anchor =[System.Windows.Forms.AnchorStyles]::Bottom
$mainGrid.Controls.Add($r7, 0, 7)
$b70 = CreateButton $Risky_Colour {ResetIE} `
	"Reset Internet Explorer Options"  `
	"Resets all Internet Explorer options to the default"
$r7.Controls.Add($b70, 0, 0)
$b76 = CreateButton $Safe_Colour {EmailTextWindow} `
	"Email TextBox"                                `
	"Creates an Outlook Email Windows to send the content of the Textbox to the support engineer"
$r7.Controls.Add($b76, 6, 0)
	

$objForm.Topmost = $False

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

#------------------------------------------------------------------------
#                                               the fat lady is singing
#------------------------------------------------------------------------