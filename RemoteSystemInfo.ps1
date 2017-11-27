$output = @()

# Sets computer list file location to to variable
$computerList = ".\Computers.txt"

# Removes spaces from Computer List to prevent errors in query
$b = Get-Content -Path $computerList
$b | ForEach {$_.TrimEnd()} | ? {$_.trim() -ne '' } > $computerList

# Queries each computer in computer list and formats data to $computerInfo array
foreach ($computer in Get-Content $computerList) {

    # Tests if $computer fails to respond to pings and adds error to $computerInfo 
    If (-NOT (Test-Connection -ComputerName $computer -count 1 -Quiet)) {

        # Sets error element to none
        $scanError = "No Ping Response"

        Format-Data

        }

    Else {

        # Tests if WMI is availible
        $testWMI = Get-WmiObject Win32_ComputerSystem -ComputerName $computer

        If ($testWMI) {
        
            # Gathers data from remote workstation
            $computerSystem = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer
            $computerOS = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer

            # Sets error element to none
            $scanError = "None"

            # Checks if the OS install date is more than two years old and generates true or false
            $replaceDate = ((Get-Date).AddYears(-2))
            $replace = ($replaceDate) -gt ($computerOS.ConvertToDateTime($computerOS.InstallDate))

            # Formats OS isntall date and adds it to the $script:computerInfo Array
            $osInstall = $computerOS.ConvertToDateTime($computerOS.InstallDate)

            Format-Data

            }

        Else {
            
            # Sets error element to none
            $scanError = "Not Windows or WMI Not Enabled"

            Format-Data

            }

        }

    Function Format-Data {
        
        # Formats simple data and adds to $script:computerInfo Array
        $script:computerInfo = @{}
        $script:computerInfo.Add("Scan Error", $scanError)
        $script:computerInfo.Add("IP Scanned", $computer)
        $script:computerInfo.Add("Computer Name", $computerSystem.Name)
        $script:computerInfo.Add("Computer Description", $computerOS.Description)
        $script:computerInfo.Add("OS Version", $computerOS.Caption)
        $script:computerInfo.Add("Manufacturer", $computerSystem.Manufacturer)
        $script:computerInfo.Add("Model", $computerSystem.Model)
        $script:computerInfo.Add("Install Date", $osInstall)
        $script:computerInfo.Add("Needs Replacement", $replace)

        }

    # Adds $computerInfo data from current loop to $output array
    $output += New-Object PSObject -Property $computerInfo
    
    # Clears created variables to prevent spilling into other array elements
    Clear-Variable -Name computerInfo,computerSystem,osInstall,computerOS,replaceDate,replace,scanError,testWMI
    # Clears previous powershell output to keep loop display clean
    Clear-Host
    # Oupts progress
    Write-Verbose ($output | Out-String) -Verbose

    }

# Exports $output array to CSV
$output | Export-Csv ComputerList.csv -NoTypeInformation