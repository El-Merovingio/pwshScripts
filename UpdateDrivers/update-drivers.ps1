# This script automates the process of checking for, downloading, and installing available driver updates.
# It uses the Microsoft Update Service Manager to interact with the Windows Update API.
$UpdateSvc = New-Object -ComObject Microsoft.Update.ServiceManager
$UpdateSvc.AddService2("7971f918-a847-4430-9279-4a52d1efe18d",7,"")
$Session = New-Object -ComObject Microsoft.Update.Session
$Searcher = $Session.CreateUpdateSearcher() 

# Verificar si la creación de objetos Update.ServiceManager y Update.Session fue exitosa antes de proceder.
if (-not $UpdateSvc) { 
    Write-Host("Error creating Update Service Manager.") -Fore Red 
    exit 
}
if (-not $Session) { 
    Write-Host("Error creating Update Session.") -Fore Red 
    exit 
}

$Searcher.ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d'
$Searcher.SearchScope =  1 # MachineOnly
$Searcher.ServerSelection = 3 # Third Party
          
$Criteria = "IsInstalled=0 and Type='Driver'"
Write-Host('Searching Driver-Updates...') -Fore Green     
$SearchResult = $Searcher.Search($Criteria)          
$Updates = $SearchResult.Updates

#if([string]::IsNullOrEmpty($Updates)){
#  Write-Host "No pending driver updates."
#}

if ($Updates.Count -eq 0) {
    Write-Host "No pending driver updates."
}

else{
  #Show available Drivers...
  $Updates | select Title, DriverModel, DriverVerDate, Driverclass, DriverManufacturer | fl
  $UpdatesToDownload = New-Object -Com Microsoft.Update.UpdateColl
  $updates | % { $UpdatesToDownload.Add($_) | out-null }
  Write-Host('Downloading Drivers...')  -Fore Green
  $UpdateSession = New-Object -Com Microsoft.Update.Session
  $Downloader = $UpdateSession.CreateUpdateDownloader()
  $Downloader.Updates = $UpdatesToDownload

#  $Downloader.Download()

  try {
      $Downloader.Download()
  } catch {
      Write-Host("Error during download: $_") -Fore Red
      exit
  }

  $UpdatesToInstall = New-Object -Com Microsoft.Update.UpdateColl
  $updates | % { if($_.IsDownloaded) { $UpdatesToInstall.Add($_) | out-null } }

  Write-Host('Installing Drivers...')  -Fore Green
  $Installer = $UpdateSession.CreateUpdateInstaller()
  $Installer.Updates = $UpdatesToInstall
  $InstallationResult = $Installer.Install()
  if($InstallationResult.RebootRequired) { 
  Write-Host('Reboot required! Please reboot now.') -Fore Red
  } else { Write-Host('Done.') -Fore Green }
  $updateSvc.Services | ? { $_.IsDefaultAUService -eq $false -and $_.ServiceID -eq "7971f918-a847-4430-9279-4a52d1efe18d" } | % { $UpdateSvc.RemoveService($_.ServiceID) }
}
