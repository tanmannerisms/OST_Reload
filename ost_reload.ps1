
class OstReload {
    static [string[]] $exemptUsers = "Administrator", "LogMeInRemoteUser", "Public", "gblackburnadmin"

    hidden [string]$machineName 
    hidden [string] $currentUser 
    hidden [string[]]$localUsers
    hidden [string[]]$userList

    ## Constructor ##
    OstReload() {
        $this.currentUser = ((Get-CimInstance -ClassName Win32_ComputerSystem).Username).Split('\')[1]
        $this.machineName = (Get-CimInstance -ClassName Win32_ComputerSystem).Name
        $this.localUsers = @((Get-ChildItem C:\Users).Name)
    }

    [void] start() {
        clear
        Write-Host "Current user set to:" ($this.getCurrentUser()).toUpper()
        $this.prompt()
    }
   
    hidden [void] prompt() {
	    Write-Host "Is this the correct user? Enter y/n"
	    $bool = Read-Host
        
	    if (($bool -eq 'y') -or ($bool -eq 'yes')) {
		    Write-Host "Let's begin"
            $this.deleteOst()
	    }
	    elseif (($bool -eq 'n') -or ($bool -eq 'no')) {
            Write-Host $this.getUserOptions()
	    }
	    else {
		    Write-Host "Oops, wrong option. Try again."
            $this.prompt()
        }
    }


    hidden [void] deleteOst() {
        [string]$dir = 'C:\Users\' + $this.currentUser + '\Appdata\Local\Microsoft\Outlook'
        [string]$backup = "OLD - " + $this.currentUser + "*"
        [string]$ostFile = $this.currentUser + '*'

        [OstReload]::stopOutlook()
        
        cd $dir
        Get-ChildItem -Filter $backup | Remove-Item
        Start-Sleep -m 1000
        Get-ChildItem -Filter $ostFile | Rename-Item -NewName {$_.Name -replace "^", "OLD - "}
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
    
    hidden [string[]] getUserOptions() {
        clear

        [int]$menuCount = 0
        [string[]]$userOptions = @(
            [string]$user = 'null'

            for ($i = 0; $i -lt $this.localUsers.Length; $i++) {
                if (($this.localUsers[$i] -eq $this.currentUser) -or ([OstReload]::exemptUsers -contains $this.localUsers[$i])) {
                    continue
                }
                else {
                    $user = '[' + $menuCount + ']' + $this.localUsers[$i]
                    $menuCount++
                    $user
                }
            }
        )
        return $userOptions
    }
}

[OstReload]$session = [OstReload]::new()
$session.start()


## $session.start()