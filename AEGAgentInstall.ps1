﻿#!ps
#timeout=900000
#maxlength=9000000
$FQDN = 'executech.hostedrmm.com'                   # Enter Automate Server FQDN                             Example: 'company.hostedrmm.com'
$LocationID = '4808'                                    # Enter Location ID (Location ID '1' = New Computers)
$SoftwarePath = "C:\Executech\automate"                # Enter Software Path in order to download the files to  Example: "C:\Support\Automate"
$ForceRip = $False                                 # Force Agent Uninstall / Install                        Example: $True or $False
& {Start-Transcript -Path "$($env:windir)\Temp\AutomateLogon.txt" -Force
    $AutomateURL = "HTTPS://$($FQDN)" 
    $AutomateSrvAddrReg = (Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service" -ErrorAction SilentlyContinue).'Server Address'
    $AutomateCompIDReg = (Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service" -ErrorAction SilentlyContinue).ID
    if (($AutomateSrvAddrReg -like "*$($FQDN)*") -and ($AutomateCompIDReg -ne $null) -and ($ForceRip -ne $True)) {
        Start-Service ltservice,ltsvcmon -PassThru
        Write-Host ""
        Write-Host "The Automate Agent is already installed." -ForegroundColor Green
        Write-Host "The Automate Agent is checking-in to: $($AutomateSrvAddrReg)" -ForegroundColor Green
        Write-Host "ComputerID: $($AutomateCompIDReg)" -ForegroundColor Green 
        Write-Host " "
        } else {
    # Remove Existing Automate Agent
        if (Test-Path "$($env:windir)\ltsvc") {
        $DownloadPath = "https://s3.amazonaws.com/assets-cp/assets/Agent_Uninstall.exe"
        $Filename = [System.IO.Path]::GetFileName($DownloadPath)
        $SoftwareFullPath = "$($SoftwarePath)\$Filename"
        $wc = New-Object System.Net.WebClient
        if (!(Test-Path $SoftwarePath)) {md $SoftwarePath}
        Set-Location $SoftwarePath
        if ((Test-Path $SoftwareFullPath)) {Remove-Item $SoftwareFullPath}
        $wc.DownloadFile($DownloadPath, $SoftwareFullPath)
        Write-Host "Removing existing Automate Agent..."
        Stop-Process -Name "ltsvcmon","lttray","ltsvc","ltclient" -Force -PassThru -ErrorAction SilentlyContinue
        Stop-Service ltservice,ltsvcmon -Force -ErrorAction SilentlyContinue
        cmd /c $SoftwareFullPath
        Write-Host " "
        Write-Host "Waiting 60 seconds to continue..." -ForegroundColor Gray
        Start-Sleep 60
            if (Test-Path "$($env:windir)\ltsvc\lterrors.txt") {
            Write-Host "  still waiting..." -ForegroundColor Gray
            Start-Sleep 90}
            if (Test-Path "$($env:windir)\ltsvc\lterrors.txt") {
            Write-Host "$($env:windir)\LTSVC folder still exists" -ForegroundColor Red} else {
            Write-Host "The Automate Agent Removed Successfully" -ForegroundColor Green
            Write-Host " "
            }}
    # Install Automate Agent
        $DownloadPath2 = "https://executech.hostedrmm.com/LabTech/Deployment.aspx?InstallerToken=e2f3aedb065448ebbeafb04aaa6f9cb1"
        $Filename2 = "AutomateAgent.msi"
        $SoftwareFullPath2 = "$SoftwarePath\$Filename2"
        if (!(Test-Path $SoftwarePath)) {md $SoftwarePath}
        Set-Location $SoftwarePath
        if ((test-path $SoftwareFullPath2)) {remove-item $SoftwareFullPath2}
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($DownloadPath2, $SoftwareFullPath2)
        Write-Host "Installing Automate Agent on $($AutomateURL)" -ForegroundColor Yellow
        msiexec.exe /i $($SoftwareFullPath2) /quiet /norestart LOCATION=$($LocationID)
        Write-Host "Waiting 3 minutes for Automate to install..." -ForegroundColor Gray
        Start-Sleep -s 180
        Write-Host "Starting Automate Services..." -ForegroundColor Yellow
        Start-Service ltservice,ltsvcmon -PassThru
        $AutomateSrvAddrReg = (Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service" -ErrorAction SilentlyContinue).'Server Address'
        $AutomateCompIDReg = (Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service" -ErrorAction SilentlyContinue).ID
            if (($AutomateSrvAddrReg -like "*$($FQDN)*") -and ($AutomateCompIDReg -ne $null)) {
            Write-Host " "
            Write-Host "The Automate Agent is checking-in to:" -ForegroundColor Green
            Write-Host "$($AutomateSrvAddrReg)" -ForegroundColor Green
            Write-Host "ComputerID: $($AutomateCompIDReg)" -ForegroundColor Green 
            Write-Host " "
            } else {
                Write-Host " "
                Write-Host "The Automate Agent is NOT installed and/or checking-in" -ForegroundColor Red
                Write-Host "Server Address: $($AutomateSrvAddrReg)" -ForegroundColor Yellow
                Write-Host "ComputerID: $($AutomateCompIDReg)" -ForegroundColor Yellow
                Write-Host " "
            }}
  Stop-Transcript }