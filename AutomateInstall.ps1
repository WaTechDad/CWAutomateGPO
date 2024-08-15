param(
    [Parameter(Mandatory)]
    [String]$CompanyName,

    [Parameter()]
    [String]$AgentURL = "https://github.com/WaTechDad/CWAutomateGPO/raw/main/Agent_Install.msi",

    [Parameter(Mandatory)]
    [String]$LocationID,

    [Parameter()]
    [Bool]$CleanUp = $False,

    [Parameter()]
    [String]$SoftwarePath = "C:\$($CompanyName)\Automate"
)

$FQDN = "Executech.hostedrmm.com"
$PassWord = "yULNxfLhgvy8N50yla58RcAHJWBGuP6L"
$UninstallerPath = "$($SoftwarePath)\Agent_Uninstall.exe"
$InstallerPath = "$($SoftwarePath)\LTAgent.msi"
$TryCount = 0
$wc = New-Object System.Net.WebClient
                        
& {
    Start-Transcript -Path "$($SoftwarePath)\Log\AutomateLogon.txt" -Force
    
    #Checks if reg keys match expected values for agent
    function Confirm-AgentStatus {
        $AutomateSrvAddrReg = (Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service" -ErrorAction SilentlyContinue).'Server Address'
        $AutomateCompIDReg = (Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service" -ErrorAction SilentlyContinue).ID
        $AutomateLocIDReg = (Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service" -ErrorAction SilentlyContinue).LocationID

        Write-Host "Checking agent status...`n" -ForegroundColor Gray

        if (($AutomateSrvAddrReg -like "*$($FQDN)*") -and ($AutomateCompIDReg -ne $null) -and ($AutomateLocIDReg -eq $LocationID)) {
            Write-Host "The Automate Agent is checking-in to: $($AutomateSrvAddrReg)"-ForegroundColor Green
            Write-Host "ComputerID: $($AutomateCompIDReg)" -ForegroundColor Green
            Write-Host "LocationID: $($AutomateLocIDReg)`n" -ForegroundColor Green
            return $True
        } elseif (Test-Path "$($env:windir)\ltsvc") {
            Write-Host "The Automate Agent is checking-in to: $($AutomateSrvAddrReg)" -ForegroundColor Red
            Write-Host "ComputerID: $($AutomateCompIDReg)" -ForegroundColor Red
            Write-Host "LocationID: $($AutomateLocIDReg)`n" -ForegroundColor Red
            return $False
        } else {
            return $False
        }
    }

    #Downloads Automate generic uninstaller
    function Get-AgentUninstall {
        $DownloadPath = "https://s3.amazonaws.com/assets-cp/assets/Agent_Uninstall.exe"
        $wc = New-Object System.Net.WebClient

        Write-Host "Downloading Automate generic uninstaller...`n" -ForegroundColor Gray

        #Checks if uninstaller exists and downloads if not
        if (Test-Path $UninstallerPath) {
            Write-Host "Uninstaller already exists...`n" -ForegroundColor Green
        } else {
            Set-Location $SoftwarePath
            $wc.DownloadFile($DownloadPath, $UninstallerPath)

            if (!(Test-Path $UninstallerPath)) {
                Write-Host "Failed to download uninstaller, exiting script...`n" -ForegroundColor Red
                Stop-Transcript
                Exit
            } else {
                Write-Host "Downloaded installer successfully, proceeding to unsinstall...`n" -ForegroundColor Green
            }
        }
    }

    #Runs Automate uninstaller
    function Remove-LTAgent {
        Get-AgentUninstall
        Write-Host "Stopping services...`n" -ForegroundColor Gray
        Stop-Process -Name "ltsvcmon","lttray","ltsvc","ltclient" -Force -ErrorAction SilentlyContinue
        Stop-Service ltservice,ltsvcmon -Force -ErrorAction SilentlyContinue
        Write-Host "Removing existing Automate Agent...`n" -ForegroundColor Gray
        cmd /c $UninstallerPath
        Start-Sleep 30
        Write-Host "Confirming agent removal...`n" -ForegroundColor Gray
        Confirm-LTRemoval
    }

    #Checks if Automate agent removal was successful
    function Confirm-LTRemoval{
        If (Test-Path "$($env:windir)\ltsvc") {
            $TryCount++
            Write-Host "Agent removal failed, checking again in 10 seconds...`n" -ForegroundColor Red
            Start-Sleep 10

            If ($TryCount -le 6) {
                Write-Host "Checking again...(Attempt $($TryCount)/6)`n"-ForegroundColor Gray
                Confirm-LTRemoval
            } else {
                Write-Host "Agent removal exceeded timeout, exiting script...`n"-ForegroundColor Red
                Stop-Transcript
                Exit
            }
        } else {
            Write-Host "Agent removal was successful!`n"-ForegroundColor Green
        }
    }

    #Installs specified Automate agent
    function Install-LTAgent {
        $wc.DownloadFile($AgentURL, $InstallerPath)

        Write-Host "Installing Automate Agent on $($FQDN)`n" -ForegroundColor Gray
        msiexec.exe /i $InstallerPath SERVERADDRESS="$FQDN" SERVERPASS="$PassWord" LOCATION="$LocationID" /quiet /norestart
        Start-Sleep -s 180
        Write-Host "Starting Automate Services...`n" -ForegroundColor Gray
        Start-Service ltservice,ltsvcmon
        #Write-Host "`n"

        if (Confirm-AgentStatus) {
            Write-Host "Success!`n" -ForegroundColor Green
        } else {
            Write-Host "The Automate Agent is NOT installed and/or checking-in:`n" -ForegroundColor Red
        }
    }

    #Checks if temp folder for installers exists and creates folder if not
    if (!(Test-Path $SoftwarePath)) {
        Write-Host "Creating folder...`n" -ForegroundColor Gray
        md $SoftwarePath
    }

    #Checks if correct agent is installed, remove incorrect agent if installed, and installs correct agent
    if (Confirm-AgentStatus) {
        Start-Service ltservice,ltsvcmon -PassThru
        Write-Host "The Automate agent is already installed.`n" -ForegroundColor Green
    } elseif (Test-Path "$($env:windir)\ltsvc") {
        Remove-LTAgent
        Install-LTAgent
    } else {
        Write-Host "No conflicting agent detected, proceeding to agent install...`n" -ForegroundColor Gray
        Install-LTAgent
    }

    Stop-Transcript

    if ($CleanUp) {
        Remove-Item $SoftwarePath -Force
    }
}