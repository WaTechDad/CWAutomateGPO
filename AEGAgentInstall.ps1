param(
    [Parameter(Mandatory)]
    [String]$CompanyName,

    [Parameter(Mandatory)]
    [String]$AgentURL,

    [Parameter()]
    [String]$SoftwarePath = "C:\$($CompanyName)\Automate"
)

$FQDN = "$($CompanyName).hostedrmm.com"
$UninstallerPath = "$($SoftwarePath)\Agent_Uninstall.exe"
$InstallerPath = "$($SoftwarePath)\LTAgent.msi"
$TryCount = 0
$wc = New-Object System.Net.WebClient
                        
& {Start-Transcript -Path "$($SoftwarePath)\Log\AutomateLogon.txt" -Force
    
    function Confirm-AgentStatus {
        $AutomateSrvAddrReg = (Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service" -ErrorAction SilentlyContinue).'Server Address'
        $AutomateCompIDReg = (Get-ItemProperty "HKLM:\SOFTWARE\LabTech\Service" -ErrorAction SilentlyContinue).ID

        if (($AutomateSrvAddrReg -like "*$($FQDN)*") -and ($AutomateCompIDReg -ne $null)) {
            return $True
        } else {
            return $False
        }
    }

    function Write-AgentStatus {
        Write-Host "The Automate Agent is checking-in to: $($AutomateSrvAddrReg)"
        Write-Host "ComputerID: $($AutomateCompIDReg)`n"
    }

    function Get-AgentUninstall {

        $DownloadPath = "https://s3.amazonaws.com/assets-cp/assets/Agent_Uninstall.exe"
        
        $wc = New-Object System.Net.WebClient

        if (Test-Path $UninstallerPath) {
            Write-Host "Uninstaller already exists...`n" -ForegroundColor Black
        } else {
            Write-Host "Downloading uninstaller...`n" -ForegroundColor Black
            Set-Location $SoftwarePath
            $wc.DownloadFile($DownloadPath, $UninstallerPath)
        }
    }

    function Install-LTAgent {
    
        $wc.DownloadFile($AgentURL, $InstallerPath)
        Write-Host "Installing Automate Agent on $($AgentURL)" -ForegroundColor Yellow
        msiexec.exe /i $InstallerPath /quiet /norestart #LOCATION=$($LocationID)
        Write-Host "Waiting 3 minutes for Automate to install..." -ForegroundColor Gray
        Start-Sleep -s 180
        Write-Host "Starting Automate Services..." -ForegroundColor Yellow
        Start-Service ltservice,ltsvcmon -PassThru

            if (Confirm-AgentStatus) {
                Write-Host "Success!"
                Write-AgentStatus -ForegroundColor Green
            } else {
                Write-Host "The Automate Agent is NOT installed and/or checking-in:" -ForegroundColor Red
                Write-AgentStatus -ForegroundColor Red
            }
    }

    function Remove-LTAgent {
    
        Get-AgentUninstall
        Write-Host "Removing existing Automate Agent..."
        Stop-Process -Name "ltsvcmon","lttray","ltsvc","ltclient" -Force -PassThru -ErrorAction SilentlyContinue
        Stop-Service ltservice,ltsvcmon -Force -ErrorAction SilentlyContinue
        cmd /c $UninstallerPath
        
        Start-Sleep 10
        Write-Host "Confirming agent removal...`n"
        Confirm-LTRemoval
    }

    function Confirm-LTRemoval{
        If (Test-Path "$($env:windir)\ltsvc") {
            $TryCount++
            Write-Host "Agent removal failed...`n"
            If ($TryCount -le 6) {
                Write-Host "Retrying removal...(Attempt $($TryCount)/6)`n"
                Confirm-LTRemoval
            } else {
                Write-Host "Agent removal exceeded timeout, exiting script...`n"
                Exit
            }
        } else {
            Write-Host "Agent removal was successful!`n"
        }
    }

if (!(Test-Path $SoftwarePath)) {
    Write-Host "Creating folder...`n" -ForegroundColor Black
    md $SoftwarePath
}

if (Confirm-AgentStatus) {
    # Checks if LabTech reg keys match the expected tenant
    Start-Service ltservice,ltsvcmon -PassThru
    Write-Host "The Automate Agent is already installed.`n" -ForegroundColor Green
    Write-AgentStatus -ForegroundColor Green
} elseif (Test-Path "$($env:windir)\ltsvc") {
    Remove-LTAgent
} else {
    Write-Host "No conflicting agent detected, proceeding to agent install..."
    Install-LTAgent
}

  Stop-Transcript }
