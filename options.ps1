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

$iconBase64 = 'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAACXBIWXMAAAsTAAALEwEAmpwYAAAEsElEQVRYw72XbWiVZRjHf9d9nnO2s50zNxjq1rayfCnUCjJRCW0JfaiYYX0qPxSBIalkUeCauLUckSSiBClBn5TobWhYLXJZ2sIpQcjIsIx82XRb7O3sbDvPy9WHs7ezc9Z5JqsL7i/32/9/vTz/634En6Yf14YG/x6oFKhSdBloiSAlKILQodAhKm3A8Twn0SzbD474uVeybRg89HIJjrVbRJ9RiPrkGwOO4gbq8re9035LBPTD2tx4PFYD3g6QPG7JNI7I/rxwtF6erx32TWDw4OuliNsIrGR2rJWg+2T+i/s7shKIHXjlPhHzJVDK7Fq7qvdYZPu+X6YlEHvvtfni6XmU2wAkJ4ypWIzk3GIGEsO4f11ER4bGSeA6K/J3TERinIAe2JYT19xToKuSK4bcZ1/FFBWP7hQ/NZteBQO9DB95F02MlYC25hVE143VhDW2Me6GdyK6ahxElcTJT5Hc/OSU44B6hNY/jRQU+SYg0UJM2ULcP9rGZlbG++M1QM14BGL7q+cJid9BImMHQ+uqsO5/KM2bxLefkLPhBTDGfyaaPsK5+PPkqbgGuSvy0t4bFoB4iRpkAhxjsJavzuhN4PYl2OeaCZQv9IcezkeH46Aps3lisxvYInpoczA+WNgFzBlfDobI21o/TVIV99IFvIFefzUQ68O50Ap2YurnNxDuzyu2BgeKHhHROamn/i2pQmDxvQRmUIiBO+5m5LMPpkJEByNDlUbU24COgk4es2iBioVgrDQMQassxSyVqYAqs8tABCQA6k6Jgiy1UEr4vyzNUS2xJBMB/Y/ANa0PlFqSUd18pMB18TqTimrmlkAgkKXvZ1ZSC6UDWOQ7Ao6N/cM3aH9vEhhwzn6PFBQSXPsoWEHfEQDaZ0bAdUk0HsVatRZTviBlybvyJ4nGo4Q2bsocjcwEOoyotKFC2shg9pmTWCvWjIO7ly/hXr6UTEPFAgIPrME+c3KaCGTA8EybcVWOpWlAwkk2n6l3dHViFkwK1vBQcox973cuQrs6M6TNSd45BceD4ya/uLtZlL6URc/DbjmVCt7Xi0TnpJMaik/pFwVof6pM2y2nwPNSwFXpz7dD3wlA/K2dB1HZmla5hUVItGAUaAhTVk7oiY0T4CPD2E0nkEiU4Lr1EAiQ+OJzvOtXkXB4tIP2o709mbTp/XBNwxYLwHVy6gMm8RwQSfGupwftmTjs6ZgLyRqRnFxCVU/h/vYrXsd1TFk53rUraHdXNikZdF27DsAARGtrOwXdm1VLurtwzrema/2SezBlFTjnz6LdXdk1SdkXqd17Y5wAQK6X8zZqfsxYrZOG3fQVTstpUE250Wk5jd30NdnOo+anvKKBPZkfpXuq5xlbzgHlWftLJIIpLUumpv0aGov5EeSrXlAfjLzRcHP6Z/muXctNwDuBZicxs47IVc81j0fq6y9kFf1YdfU8Y5lGYPUswbd4jrcx0jDh+fjrL9PuSEPDzXBPX6VCHUos44PF34gp1IWNVZkJ3FfbG6itnRuw7V2IbAIKff7x9qnqEU/1zemAZ9B3Rwt98+bgSPH8hz3YgNFlow+ZktE7OkA6RGlz4Vh+941mOXzY9nPvP/2/IUmX17Y+AAAAAElFTkSuQmCC'
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

$test = New-ScheduledTaskAction -Execute PowerShell.exe -Argument '-file "%appdata%\Deletion Scheduler\DS.ps1"' -WorkingDirectory $OpenFileDialog.FileName
$yo = New-ScheduledTaskTrigger -At 12:00pm -Daily
$task = Register-ScheduledTask -TaskName "Deletion Scheduler $($diff.Days) days" -Trigger $yo -Action $test -Force

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
