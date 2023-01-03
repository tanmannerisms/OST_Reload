
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
        Clear-Host
        $this.prompt()
    }
   
    hidden [void] prompt() {
        $this.printCurrentUser()
	    Write-Host "Is this the correct user? Enter y/n"
	    $bool = Read-Host
        
	    if (($bool -eq 'y') -or ($bool -eq 'yes')) {
		    Write-Host "Let's begin"
            $this.deleteOst()
	    }
	    elseif (($bool -eq 'n') -or ($bool -eq 'no')) {
            $this.userList = $this.getUserOptions()
            $this.currentUser = $this.userList[[OstReload]::getSelection($this.userList)]
            $this.printCurrentUser()
	    }
	    else {
		    Write-Host "Oops, wrong option. Try again."
            $this.start()
        }
    }


    hidden [void] deleteOst() {
        [string]$dir = 'C:\Users\' + $this.currentUser + '\Appdata\Local\Microsoft\Outlook'
        [string]$backup = "OLD - " + $this.currentUser + "*"
        [string]$ostFile = $this.currentUser + '*'

        [OstReload]::stopOutlook()
        Set-Location $dir
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
    
############### Start Getters ###############

    hidden [string] getCurrentUser() {
	    return $this.currentUser
    }
    
    hidden [string[]] getUserOptions() {
        Clear-Host

        [string[]]$userOptions = @(
            for ($i = 0; $i -lt $this.localUsers.Length; $i++) {
                if (($this.localUsers[$i] -eq $this.currentUser) -or ([OstReload]::exemptUsers -contains $this.localUsers[$i])) {
                    continue
                }
                else {
                    $this.localUsers[$i]
                }
            }
        )
        return $userOptions
    }

################ End Getters ###########################

    hidden static [string[]] makeMenu([string[]]$menuItems) {
        [int]$i = 0
        [string[]]$menu = @(
            foreach ($item in $menuItems) {
                $item = '[' + $i + '] ' + $item
                $item
            }
        )

        return $menu
    }

    static [int] getSelection([string[]]$menuOptions) {
        [OstReload]::printMenu($menuOptions)
        [int]$answer = Read-Host 'Select an option from above by entering the number associated with the desired selection (default is 0)'
        if ([OstReload]::validateMenuSelection($answer, $menuOptions) -eq $false) {
            Clear-Host
            Write-Host "Let's try that again..."
            Start-Sleep -m 500
            return [OstReload]::getSelection($menuOptions)
        }
        return $answer
    }

    hidden [void] printCurrentUser() {
        Write-Host "Current user set to:" ($this.getCurrentUser()).toUpper()
    }

    static [void] printMenu([string[]]$menuOptions) {
        foreach ($option in [OstReload]::makeMenu($menuOptions)) {
            Write-Host $option
        }
    }

    static [boolean] validateMenuSelection([int]$answer, [string[]]$menuOptions) {
        if (($answer -lt $menuOptions.Length) -and ($answer -ge 0)) {
            return $true
        }
        else { return $false }
    }
}

[OstReload]$session = [OstReload]::new()
$session.start()


## $session.start()