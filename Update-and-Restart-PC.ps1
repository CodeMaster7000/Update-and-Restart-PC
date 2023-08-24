$WUAgent = New-Object -ComObject Microsoft.Update.ServiceManager
$Searcher = $WUAgent.CreateUpdateSearcher()
$Searcher.ServerSelection = 3
$Searcher.ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d' # Windows Update
$Searcher.IncludePotentiallySupersededUpdates = $false
$Searcher.Criteria = 'IsInstalled=0 and IsHidden=0 and Type="Software" and CategoryIDs contains "SecurityUpdates"'
$SearchResult = $Searcher.Search()
if ($SearchResult.Updates.Count -gt 0) {
    Write-Host "Available security updates:"
    $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($Update in $SearchResult.Updates) {
        $UpdatesToInstall.Add($Update)
        Write-Host "----------------------------------"
        Write-Host "Title: $($Update.Title)"
        Write-Host "Description: $($Update.Description)"
        Write-Host "KB Article ID: $($Update.KBArticleIDs)"
        Write-Host "----------------------------------"
    }
    $Installer = $WUAgent.CreateUpdateInstaller()
    $Installer.Updates = $UpdatesToInstall
    $InstallationResult = $Installer.Install()
    if ($InstallationResult.ResultCode -eq 2) {
        Write-Host "Security updates successfully installed!"
        Write-Host "Restarting your PC..."
        Restart-Computer -Force
    } else {
        Write-Host "Failed to install updates. Check for sources of errors and try again."
    }
}
else {
    Write-Host "No security updates available."
}
