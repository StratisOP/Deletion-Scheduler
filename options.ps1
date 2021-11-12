$AssemblyFullName = 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
$Assembly = [System.Reflection.Assembly]::Load($AssemblyFullName)
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.AddExtension = $false
$OpenFileDialog.CheckFileExists = $false
$OpenFileDialog.DereferenceLinks = $true
$OpenFileDialog.Filter = "Folders|`n"
$OpenFileDialog.Multiselect = $false
$OpenFileDialog.Title = "Select folder"
$OpenFileDialogType = $OpenFileDialog.GetType()
$FileDialogInterfaceType = $Assembly.GetType('System.Windows.Forms.FileDialogNative+IFileDialog')
$IFileDialog = $OpenFileDialogType.GetMethod('CreateVistaDialog', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($OpenFileDialog, $null)
$null = $OpenFileDialogType.GetMethod('OnBeforeVistaDialog', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($OpenFileDialog, $IFileDialog)
[uint32]$PickFoldersOption = $Assembly.GetType('System.Windows.Forms.FileDialogNative+FOS').GetField('FOS_PICKFOLDERS').GetValue($null)
$FolderOptions = $OpenFileDialogType.GetMethod('get_Options', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($OpenFileDialog, $null) -bor $PickFoldersOption
$null = $FileDialogInterfaceType.GetMethod('SetOptions', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $FolderOptions)
$VistaDialogEvent = [System.Activator]::CreateInstance($AssemblyFullName, 'System.Windows.Forms.FileDialog+VistaDialogEvents', $false, 0, $null, $OpenFileDialog, $null, $null).Unwrap()
[uint32]$AdviceCookie = 0
$AdvisoryParameters = @($VistaDialogEvent, $AdviceCookie)
$AdviseResult = $FileDialogInterfaceType.GetMethod('Advise', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $AdvisoryParameters)
$AdviceCookie = $AdvisoryParameters[1]
$Result = $FileDialogInterfaceType.GetMethod('Show', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, [System.IntPtr]::Zero)
$null = $FileDialogInterfaceType.GetMethod('Unadvise', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $AdviceCookie)
if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
	$FileDialogInterfaceType.GetMethod('GetResult', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $null)
}if ($Result -eq -2147023673) { 
	Remove-Item  DSoptions.ini -Force -ErrorAction Ignore
	break
}
if (test-path DSoptions.ini) { remove-item DSoptions.ini -Force }
Write-Output $OpenFileDialog.FileName | out-file .\DSoptions.ini -Force
(get-item DSoptions.ini).Attributes += 'Hidden'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$iconBase64 = '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'
$iconBytes = [Convert]::FromBase64String($iconBase64)
$stream = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)

$form = New-Object Windows.Forms.Form -Property @{
	StartPosition   = [Windows.Forms.FormStartPosition]::CenterScreen
	FormBorderStyle = 'FixedDialog'
	MaximizeBox     = $false;
	MinimizeBox     = $false;
	Size            = New-Object Drawing.Size 233, 223
	Text            = 'Select a Date'
	Topmost         = $true
	Icon            = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))
}

$calendar = New-Object Windows.Forms.MonthCalendar -Property @{
	ShowTodayCircle   = $false
	MaxSelectionCount = 1
}
$form.Controls.Add($calendar)

$okButton = New-Object Windows.Forms.Button -Property @{
	Location     = New-Object Drawing.Point 43, 165
	Size         = New-Object Drawing.Size 75, 23
	Text         = 'OK'
	DialogResult = [Windows.Forms.DialogResult]::OK
}
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object Windows.Forms.Button -Property @{
	Location     = New-Object Drawing.Point 118, 165
	Size         = New-Object Drawing.Size 75, 23
	Text         = 'Cancel'
	DialogResult = [Windows.Forms.DialogResult]::Cancel
}
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$result = $form.ShowDialog()

if ($result -eq [Windows.Forms.DialogResult]::OK) {
	$date = $calendar.SelectionStart
	$current = Get-date
	$diff = New-TimeSpan -Start $current -End $date
	$diff.Days | add-content .\DSoptions.ini
}if ($result -eq 'Cancel') { 
	Remove-Item  DSoptions.ini -Force -ErrorAction Ignore
	break
}
Copy-Item -Path .\*.ini -Destination $OpenFileDialog.FileName -Force
$DS = @'
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function ShowConsole
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 5)
}

function HideConsole
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0)
}
HideConsole | Out-Null
if (test-path DSoptions.ini) {
	$diff = Get-Content DSoptions.ini -Tail 1 
	$time = (Get-Date).AddDays($diff)
	$path = Get-Content DSoptions.ini -Head 1
	$LogFile = ".\log.txt"
	$prompt = new-object -comobject wscript.shell 
	$answer = $prompt.popup("Deleting files in $path created before $time`n", 90, "Deletion Scheduler", 1)   
	if ($answer -eq 1) {
		$files = Get-ChildItem -Path $path -Force -Recurse -exclude DSoptions.ini |
		Where-Object { $_.CreationTime -lt $time }
		foreach ($File in $Files) { 
			if ($Null -ne $File) { 
				$myDate = Get-Date -UFormat "%m-%d-%Y-%R "
				$content = $myDate + $File.FullName
				Add-Content $LogFile $content
			} 
		} 
		Get-ChildItem -Path $path -Force -exclude DSoptions.ini |
		Where-Object { $_.CreationTime -lt $time } |
		Remove-Item  -Force -Recurse			
	}if ($answer -eq -1) {
		$files = Get-ChildItem -Path $path -Force -Recurse -exclude DSoptions.ini |
		Where-Object { $_.CreationTime -lt $time }
		foreach ($File in $Files) { 
			if ($Null -ne $File) { 
				$myDate = Get-Date -UFormat "%m-%d-%Y-%R "
				$content = $myDate + $File.FullName
				Add-Content $LogFile $content
			} 
		} 
		Get-ChildItem -Path $path -Force -exclude DSoptions.ini |
		Where-Object { $_.CreationTime -lt $time } |
		Remove-Item  -Force -Recurse
	}
	else { break }
}
else {
	Write-Host "Config file is missing!"
	Exit
}
'@
$DS | Out-File .\DS.ps1
New-Item -Path "$($env:appdata)\Deletion` Scheduler" -ItemType directory -Force | Out-Null
Copy-Item -Path .\DS.ps1 -Destination "$($env:appdata)\Deletion` Scheduler\" -Force
remove-item DS.ps1 -Force -ErrorAction Ignore
remove-item DSoptions.ini -Force -ErrorAction Ignore
$time = (Get-Date).AddDays($diff.Days)
$path = $OpenFileDialog.FileName
$stream.Dispose()
$Form.Dispose()
if (test-path -Path "$($env:appdata)\Deletion` Scheduler\log.txt") {
	$LogFile = "$($env:appdata)\Deletion` Scheduler\log.txt"
}
else { $LogFile = New-Item -Path "$($env:appdata)\Deletion` Scheduler\log.txt" -Force }

$last2parts = $($OpenFileDialog.FileName)
$last2parts = $last2parts.replace('\', '-')
$last2parts = $last2parts.replace(':', '')
$last2parts = $last2parts.Split("-") | Select-Object -Last 3

$datedesc = "$($diff.Days)" +" days"
if ( $($diff.Days) -eq 0) {
	$datedesc = "today"
}
if ($($diff.Days) -lt 0) {
	$datedesc =  $datedesc.replace('-', '')
}

$TaskAction = New-ScheduledTaskAction -Execute PowerShell.exe -Argument '-file "%appdata%\Deletion Scheduler\DS.ps1"' -WorkingDirectory $OpenFileDialog.FileName
$TaskTrigger = New-ScheduledTaskTrigger -At 12:00pm -Daily
$TaskRegister = Register-ScheduledTask -TaskName "DeletionScheduler{$($last2parts -join "-" )}" -Description "Deleting files with creation date older than $datedesc at $($OpenFileDialog.FileName)" -Trigger $TaskTrigger -Action $TaskAction -Force

$prompt = new-object -comobject wscript.shell 
$answer = $prompt.popup("Delete files in $path created before $time ?`n", 90, "Deletion Scheduler", 4)   
if ($answer -eq 6) {
	$files = Get-ChildItem -Path $path -Force -Recurse -exclude DSoptions.ini |
	Where-Object { $_.CreationTime -lt $time }
	foreach ($File in $Files) { 
		if ($Null -ne $File) { 
			$myDate = Get-Date -UFormat "%m-%d-%Y-%R "
			$content = $myDate + $File.FullName
			Add-Content $LogFile $content
		} 
	} 
	Get-ChildItem -Path $path -Force -exclude DSoptions.ini |
	Where-Object { $_.CreationTime -lt $time } |
	Remove-Item  -Force -Recurse -ErrorAction SilentlyContinue -ErrorVariable removeErrors
	$removeErrors | where-object { $_.Exception.Message -notlike '*it is being used by another process*' }
			
}if ($answer -eq -1) {
	$files = Get-ChildItem -Path $path -Force -Recurse -exclude DSoptions.ini |
	Where-Object { $_.CreationTime -lt $time }
	foreach ($File in $Files) { 
		if ($Null -ne $File) { 
			$myDate = Get-Date -UFormat "%m-%d-%Y-%R "
			$content = $myDate + $File.FullName
			Add-Content $LogFile $content
		} 
	} 
	Get-ChildItem -Path $path -Force -exclude DSoptions.ini |
	Where-Object { $_.CreationTime -lt $time } |
	Remove-Item  -Force -Recurse -ErrorAction SilentlyContinue -ErrorVariable removeErrors
	$removeErrors | where-object { $_.Exception.Message -notlike '*it is being used by another process*' }
}
else { break }	
