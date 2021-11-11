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

$iconBase64 = 'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjAuMTnU1rJkAAAB50lEQVRIS7WWzytEURTHZ2FhaWFhYWFhYWFhaWFh6c+wsGCapJBJU0hRSrOgLBVSmkQoSpqyUJISapIFJU0i1KQp1PG9826vO9+Z97Pr22dz3pxzv/PO/fUSkvxfOLYOx9bh2DocBzPZKlPtku2VuS7JtMlwIydUw7EnGO50Rd4ehPRdlru87GWUMZVU4LgOqLza0cP56PdHNga4NthgsUc+i3qIQOH9qDzAYLZTyiVd7Ah/8/lGztclNyLbY5JfksKxeuiKRvAzwOy93OsyR+9Pam6HajLHm5UfXhTQT34GqDH1eCGjTZxjkmqQmQ5+6Gdgth5NQLsoIRwca7AoTZ1k1UM0p7Y/QXCsof4sdOvRrRlgeZgyu49eY/0cTDO76ShzcJnTQ0O0NvA2XioWqjIrcKy5PdQ1kLl90CKsVC9F2GiYVVPuiQaDiRb1TzGWw9eHzoEiGGwO6hpHWFRe04vuu4pggCPIFPwowSWmAXa/qdKr5zaOaQBwyps6W+UEh/gGtasFlrjCKC2+ATia15WucHpf76vna/2ylVLntnmeRzbApsUQ4RXZwAFnAC7eMMLNSrWhDEC6RfUac1D3+sRew87HhVzvC4PjYLBecRxhDsByn/qEoYRqOLYOx9bh2DocWyaZ+APgBBKhVfsHwAAAAABJRU5ErkJggg=='
$iconBytes = [Convert]::FromBase64String($iconBase64)
$stream = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)

$form = New-Object Windows.Forms.Form -Property @{
	StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
	Size          = New-Object Drawing.Size 243, 230
	Text          = 'Select a Date'
	Topmost       = $true
	Icon          = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))
}

$calendar = New-Object Windows.Forms.MonthCalendar -Property @{
	ShowTodayCircle   = $false
	MaxSelectionCount = 1
}
$form.Controls.Add($calendar)

$okButton = New-Object Windows.Forms.Button -Property @{
	Location     = New-Object Drawing.Point 38, 165
	Size         = New-Object Drawing.Size 75, 23
	Text         = 'OK'
	DialogResult = [Windows.Forms.DialogResult]::OK
}
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object Windows.Forms.Button -Property @{
	Location     = New-Object Drawing.Point 113, 165
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
$DS1 = '
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
'
$DS1 | Out-File .\DS.ps1
invoke-ps2exe DS.ps1 .\DS.exe -noConsole -title 'Deletion Scheduler' -version '0.1.0.5' |Out-Null
Copy-Item -Path .\DS.exe -Destination "$($env:appdata)\Deletion` Sceduler\" -Force
remove-item DS.ps1 -Force -ErrorAction Ignore
remove-item DS.exe -Force -ErrorAction Ignore
remove-item DSoptions.ini -Force -ErrorAction Ignore
$time = (Get-Date).AddDays($diff.Days)
$path = $OpenFileDialog.FileName
$stream.Dispose()
$Form.Dispose()
if (test-path -Path "$($env:appdata)\Deletion` Sceduler\log.txt") {
	$LogFile = "$($env:appdata)\Deletion` Sceduler\log.txt"
}
else { $LogFile = New-Item -Path "$($env:appdata)\Deletion` Sceduler\log.txt" -Force }

$test = New-ScheduledTaskAction -Execute "$($env:appdata)\Deletion Sceduler\DS.exe" -WorkingDirectory $OpenFileDialog.FileName
$yo = New-ScheduledTaskTrigger -At 12:00pm -Daily
$task = Register-ScheduledTask -TaskName "Deletion Scheduler" -Trigger $yo -Action $test -Force

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