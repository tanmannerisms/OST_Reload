
class OstReload {
    static [string[]] $exemptUsers = "Administrator", "LogMeInRemoteUser", "Public", "gblackburnadmin"

    hidden [string]$machineName 
    hidden [string] $currentUser 
    hidden [string[]]$localUsers
    hidden [string[]]$userList

    ## Constructor ##
    Device() {
        $this.currentUser = ((Get-CimInstance -ClassName Win32_ComputerSystem).Username).Split('\')[1]
        $this.machineName = (Get-CimInstance -ClassName Win32_ComputerSystem).Name
        $this.localUsers = @((Get-ChildItem C:\Users).Name)
    }

    [void] start() {
        Write-Host "Current user set to:" ($this.getCurrentUser()).toUpper()
        [OstReload]::prompt()
    }
   
    static [void] prompt() {
	    Write-Host "Is this the correct user? Enter y/n`n"
	    $bool = Read-Host
        
	    if (($bool -eq 'y') -or ($bool -eq 'yes')) {
		    Write-Host "Let's begin"
            [OstReload]::deleteOst()
	    }
	    elseif (($bool -eq 'n') -or ($bool -eq 'no')) {
    		
	    }
	    else {
		    Write-Host "Oops, wrong option. Try again."
            [OstReload]::prompt()
	    }
    }


    hidden [void] deleteOst() {
        [OstReload]::stopOutlook()
        cd C:\Users\$this.currentUser\Appdata\Local\Microsoft\Outlook
        Get-ChildItem -Filter "OLD - $this.currentUser*" | Remove-Item
        Start-Sleep -m 1000
        Get-ChildItem -Filter "$this.currentUser*" | Rename-Item -NewName {$_.Name -replace "^", "OLD - "}
        [OstReload]::startOutlook()
    }

    static [void] startOutlook() {
        Start-Process Outlook.exe
    }

    static [void] stopOutlook() {
        Stop-Process -Name "Outlook" -Force
    }
    
    ## Getters ##

    static [int] getSelection([int[]]$options) {
        return 0
    }


    hidden [string] getCurrentUser() {
	    return $this.currentUser
    }
    
    hidden [string[]] getLocalUsers($currentUser) {
        $this.localUsers = @((Get-ChildItem C:\Users).Name)

        Write-Host

        for ($i = 0; $i -lt $this.localUsers.Length; $i++) {
            if ($this.localUsers[$i] -contains [OstReload]::exemptUsers -or $this.currentUser) {
                
            }
        }
        return $this.localUsers
    }
}

[OstReload]$session = [OstReload]::new()
$session.start()


## $session.start()