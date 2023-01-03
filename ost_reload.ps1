
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
   
    static [void] prompt() {
	    Write-Host "Is this the correct user? Enter y/n`n"
	    $bool = Read-Host
        
	    if (($bool -eq 'y') -or ($bool -eq 'yes')) {
		    Write-Host 
	    }
	    elseif (($bool -eq 'n') -or ($bool -eq 'no')) {
    		
	    }
	    else {
		    Write-Host "Oops, wrong option. Try again."
            [OstReload]::prompt()
	    }
    }

    static [int] getSelection([int[]]$options) {
        return 0
    }


    [void] start() {
        Write-Host "Current user set to:" ($this.getCurrentUser()).toUpper()
        [OstReload]::prompt()
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
    }
}

[OstReload]$session = [OstReload]::new()
$session.setCurrUser()


## $session.start()