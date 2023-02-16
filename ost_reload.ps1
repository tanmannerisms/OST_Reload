class Session {
    hidden [string]$machineName 
    hidden [string] $currentUser 
    hidden [string[]]$localUsers
    hidden [string[]]$localAdmins

    Session() {
        $this.currentUser = ((Get-CimInstance -ClassName Win32_ComputerSystem).Username).Split('\')[1]
        $this.machineName = (Get-CimInstance -ClassName Win32_ComputerSystem).Name
        $this.localUsers = @((Get-ChildItem C:\Users).Name)
        $this.localAdmins = "Administrator", "AzureAdmin", "LogMeInRemoteUser"
    }

    hidden [void] printCurrentUser() {
        Write-Host "Current user set to:" ($this.getCurrentUser()).toUpper()
    }

    [string] getCurrentUser() {
	    return $this.currentUser
    }
}

class OstReload : Session {
    static hidden [string[]]$exemptionArray = "Public", "gblackburnadmin"

    hidden [string[]]$exemptUsers 
    hidden [string[]]$applicableUserList
    hidden [String[]]$runningOfficeApps

    ## Constructor ##
    OstReload() {
        $this.exemptUsers = [OstReload]::exemptionArray
        $this.addExemptUser($this.localAdmins)
        $this.addExemptUser("$($this.machineName)$")
    }
########## Instance Methods ##########
    [void] start() {
        $initialDirectory = Get-Location
        Clear-Host
        $this.prompt()
        Set-Location $initialDirectory
    }

    hidden [void] addExemptUser([string]$exemptUser) {
        $this.exemptUsers += $exemptUser
    }

    hidden [void] addExemptUser([string[]]$exemptUser) {
        foreach ($item in $exemptUser) {
            $this.exemptUsers += $item
        }
    }
   
    hidden [void] prompt() {
        Clear-Host
        $this.printCurrentUser()
	    Write-Host "Is this the correct user? Enter y/n"
	    $bool = Read-Host
        
	    if (($bool -eq 'y') -or ($bool -eq 'yes')) {
		    Write-Host "Let's begin"
            $this.deleteOst()
	    }
	    elseif (($bool -eq 'n') -or ($bool -eq 'no')) {
            $this.addExemptUser($this.currentUser)
            $this.applicableUserList = $this.getUserOptions()
            $this.currentUser = $this.applicableUserList[[OstReload]::getSelection($this.applicableUserList)]
            $this.prompt()
	    }
	    else {
            Clear-Host
		    Write-Host "Oops, wrong option. Try again."
            $this.prompt()
        }
    }

    hidden [void] deleteOst() {
        [string]$dir = 'C:\Users\' + $this.currentUser + '\Appdata\Local\Microsoft\Outlook'
        [string]$backup = "OLD - " + $this.currentUser + "*"
        [string]$ostFile = $this.currentUser + '*'

        $this.stopOffice()
        Set-Location $dir
        Get-ChildItem -Filter $backup | Remove-Item
        Start-Sleep -m 1000
        Get-ChildItem -Filter $ostFile | Rename-Item -NewName {$_.Name -replace "^", "OLD - "}
        $this.startOffice()
    }
    [void] startOffice() {
        foreach ($process in $this.runningOfficeApps) {
            Start-Process $process
        }
    }

    [void] stopOffice() {
        [String[]]$processes =  "OUTLOOK", "EXCEL", "WINWORD", "POWERPNT", "ONENOTE", "MSPUB", "MSACCESS"
        foreach($process in $processes) {
            if ( Get-Process -Name $process -ErrorAction SilentlyContinue ) {
                $this.runningOfficeApps += $process
                Stop-Process -Name $process -Force -ErrorAction Stop
            }
            else {
                continue
            }
        }
    }
    
    ##### Start Getters #####
    
    hidden [string[]] getUserOptions() {
        Clear-Host

        [string[]]$userOptions = @(
            foreach ($user in $this.localUsers) {
                if ($this.exemptUsers -contains $user) {
                    continue
                }
                else {
                    $user
                }
            }
        )
        return $userOptions
    }
    ##### End Getters #####
}

class Menu {
    hidden [String[]]$menuItems
    hidden [String[]]$menu

    [void] setMenuOptions([string[]]$menuItems) {
        $this.menuItems = $menuItems
        $this.makeMenu
    }

    [void] makeMenu() {
        [int]$i = 0
        $this.menu = @(
            foreach ($item in $this.menuItems) {
                $item = '[' + $i + '] ' + $item
                $item
                $i++
            }
        )
    }

    [void] printMenu() {
        if ($this.menu.length -gt 0) {
            foreach ($option in $this.menu) {
                Write-Host $option
            }
        }
        else {
            Write-Host "The list of other possible options that you could select is empty.`nPlease contact tmills@clydeinc.com if you believe this to be incorrect."
            Start-Sleep -Seconds 15
            exit
        }
    }

    [int] getSelection() {
        [int]$answer = Read-Host 'Select an option from above by entering the number associated with the desired selection (default is 0)'
        if ($this.validateMenuSelection($answer) -eq $false) {
            Clear-Host
            Write-Host "Let's try that again..."
            Start-Sleep -m 500
            return $this.getSelection()
        }
        return $answer
    }

    [boolean] validateMenuSelection([int]$answer) {
        if (($answer -lt $this.menu.Length) -and ($answer -ge 0)) {
            return $true
        }
        else { return $false }
    }
}

[OstReload]$session = [OstReload]::new()
$session.start()