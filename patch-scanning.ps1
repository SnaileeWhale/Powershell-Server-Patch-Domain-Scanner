<#
-----------------------------------------

Patch Scanning Utility 
-targets either 2008r2 or 2012r2 for a specified KB 

-----------------------------------------
#>
#Environmentals
$OutputArray = @()
$filelocation = "C:\Temp\Patch-Report.csv"
$names = Get-content E:\computers.txt # breakout text array for testing 
#loop Start
do
    {
#Menu Display
write-host " "
write-host "Server Patch Scanning utility" 
write-host " "
write-host "Select Operating System"
write-host " "
write-host " 1) Server 2012r2" 
write-host " 2) Server 2008r2"
write-host " "
#Switch Call
$OperatingSystem = Read-Host " "
    if(($OperatingSystem -eq "1") -or ($OperatingSystem -eq "2") -or ($OperatingSystem -eq "3"))
    {
        #switch start to build array for foreach
        switch($OperatingSystem)
        {
            "1"
            {
            $2012r2 = Get-ADComputer -Filter {OperatingSystem -like "*Windows Server 2012 R2*"} -Properties Operatingsystem | select Name, OperatingSystem | sort name
            $names = $2012r2.Name 
            $ServerType = "2012r2"
            }
            "2"
            {
            $2008r2 = Get-ADComputer -Filter {OperatingSystem -like "*Windows Server 2008 R2*"} -Properties Operatingsystem | select Name, OperatingSystem | sort name
            $names = $2008r2.Name
            $ServerType = "2008r2"
            }
            "3"
            {
            $names = Get-content E:\computers.txt
            $ServerType = "Textfile"
            }
        }
#Call for Patch required to Scan, currently no intelligence to detect if input is valid, scan based on whatever is entered
$patchID = Read-host "Enter Patch KB e.g 'KB4012213'" 
#Start foreach loop
foreach ($name in $names)
    {
            #Begin building new oject table object/results array
            $ServerObject = New-Object PSObject
            #Test for connectivity and continue if available  
            $testconnection = Test-Connection -ComputerName $name -Count 1 -ErrorAction SilentlyContinue
            if ($testconnection)
            {
                    $connectionOK = Write-Host -NoNewline "$name`t`t`tOnline" -ForegroundColor Green
                    $connected = "Yes"
                          Try
                          {
                          #test to see if patch is install
                          $Getpatch = Get-HotFix -Id $PatchID -ComputerName $name -ev Err -ErrorAction SilentlyContinue
                                    if($Getpatch.PSComputerName -eq $name)
                                    {
                                    $PatchedOK = Write-Host "`t`tPatched" -ForegroundColor Green
                                    $Patched = "Yes"
                                    $connectionOK
                                    $PatchedOK
                                    }
                                    Else
                                    {
                                    $PatchedNotOK = Write-Host "`t`tNot patched" -ForegroundColor Red
                                    $Patched = "No"
                                    $connectionOK
                                    $PatchedNotOK
                                    }
                          }
                          Catch
                          {
                          Write-Host "`t`tNot responding to query, possible firewall, invalid patch id, WinRM or RPC error" -ForegroundColor yellow
                          $Patched = "Unknown"
                          }
             }
             else
             {
             Write-Host "$name`t`t`tOffline" -ForegroundColor darkred
             $connected = "No"
             $patched = "Unknown"
             }
#Builds result line to new psobject
Add-Member -InputObject $ServerObject -MemberType NoteProperty -Name 'Computer Name' -Value $name 
Add-Member -InputObject $ServerObject -MemberType NoteProperty -Name 'Connected' -Value $connected
Add-Member -InputObject $ServerObject -MemberType NoteProperty -Name 'Patched' -Value $Patched
$OutputArray += $ServerObject
}
$OutputArray | sort patched | FT -AutoSize
#start loop to contain  try export csv and catch whether export is possible and try catch email is possible  
do{
        try
            {
            $OutputArray | sort patched | Export-Csv $filelocation -NoTypeInformation
            write-host " "
            Write-host "Exported CSV to C:\Temp\Patch-Report.csv" -ForegroundColor green
            write-host " "
            try
            {
                $SendEmail = Read-host "Send Email? (y/n)?"
                if($SendEmail -eq "y")
                    {
                    write-host " "
                    $emailaddress = Read-host "Enter Email Address"
                    write-host " "
                    $smtp = Read-host "Enter SMTP Mail Server e.g 'mail'" 
                    send-mailmessage -From $emailaddress -subject "Server $ServerType machines missing $patchID" -To $emailaddress -Attachments $filelocation -smtpserver $smtp 
                    write-host " "
                    "Email Sent"
                    write-host " "
                    $retry = "n"
                    }
                else
                    {
                    write-host " "
                    "No Email Sent"
                    write-host " "
                    $retry = "n"
                    }
            }            
            Catch
            {
            write-host " "
            Write-host "Error invalid SMTP Server" -ForegroundColor Red
            write-host " "
            $retry = read-host "Retry? (y/n)"
            write-host " "
            }           
            }
       Catch
            {
            write-host " "
            write-host "Error Exporting to C:\Temp\Patch-Report.csv" -ForegroundColor Red
            write-host " "
            $retry = read-host "Retry? (y/n)"
            write-host " "
            }
}
while($retry -eq "y")
}
    else
            {
            write-host " "
            "Invalid Selection"
            write-host " "
            }
$repeat = Read-Host "Exit? (y/n)"
}
while($repeat -ne "y")
write-host " "
write-host "Exiting..."
