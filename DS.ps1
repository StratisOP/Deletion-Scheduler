if (test-path DSoptions.ini) {
  $diff = Get-Content DSoptions.ini -Tail 1 
  $time = (Get-Date).AddDays($diff)
  $path = Get-Content DSoptions.ini -Head 1
  #$id = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
  $LogFile = ".\log.txt"
  $prompt = new-object -comobject wscript.shell 
  $answer = $prompt.popup("Deleting files in $path created before $time`n", 90, "Deletion Scheduler", 1)   
  if ($answer -eq 1) {
				$files = Get-ChildItem -Path $path -Force -Recurse -exclude DSoptions.ini |
				Where-Object { $_.CreationTime -lt $time }
				foreach ($File in $Files) { 
      if ($Null -ne $File) { 
								#$status = " User "
								$myDate = Get-Date -UFormat "%m-%d-%Y-%R "
								$content = $myDate + $File.FullName# + $status + $id
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
								#$status = " Timed Out "
								$myDate = Get-Date -UFormat "%m-%d-%Y-%R "
								$content = $myDate + $File.FullName# + $status + $id
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