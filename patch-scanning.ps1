<#




-----------------------------------------

PowerShell Script 

Patch Scanning Utility 

-----------------------------------------

#>

#Environmentals

$smtpserver = #Enter smtp server
$Email = #Enter EMail Address 
$filelocation = C:\temp\Patched.CSV #Enter File Location
$OutputArray = @()


#loop start

do
{


#switch loop

write-host " "
write-host "Server Patch Scanning utility" 
write-host " "
write-host "Select Operating System"
write-host " "
write-host " 1) Server 2012r2" 
write-host " 2) Server 2008r2"
write-host " "
$OperatingSystem = Read-Host " "



if(($OperatingSystem -eq "1") -or ($OperatingSystem -eq "2"))

{

switch($OperatingSystem)
{
"1"
{

$2012r2 = Get-ADComputer -Filter {OperatingSystem -like "*Windows Server 2012 R2*"} -Properties Operatingsystem | select Name, OperatingSystem | sort name

$names = $2012r2.Name 

}
"2"
{


$2008r2 = Get-ADComputer -Filter {OperatingSystem -like "*Windows Server 2008 R2*"} -Properties Operatingsystem | select Name, OperatingSystem | sort name

$names = $2008r2.Name

}


}



$patchID = Read-host "Enter Patch KB" 




#For each loop

foreach ($name in $names)

{

            $ServerObject = New-Object PSObject

            $testconnection = Test-Connection -ComputerName $name -Count 1 -ErrorAction SilentlyContinue

            if ($testconnection)

            {


                    Write-Host "$name is connected" -ForegroundColor Green

                    $connected = "Yes"

                          Try
    
                          {

                          $Getpatch = Get-HotFix -Id $PatchID -ComputerName $name -ev Err -ErrorAction SilentlyContinue

                                    if($Getpatch.PSComputerName -eq $name)

                                    {

                                    Write-Host "---$name is patched" -ForegroundColor Green

                                    $Patched = "Yes"


                                    }

                                    Else

                                    {

                                    Write-Host "---$name is not patched" -ForegroundColor Red

                                    $Patched = "No"

                                    }

                          }

                          Catch

                          {
                          Write-Host "---$name is not responding to query possibly firewalled,WinRM not Enabled or RPC error" -ForegroundColor yellow

                          $Patched = "Unknown"

                          }

             }

             else
 
             {

             Write-Host "$name is not online" -ForegroundColor darkred

             $connected = "No"
             $patched = "Unknown"
    
             }

Add-Member -InputObject $ServerObject -MemberType NoteProperty -Name 'Computer Name' -Value $name 
Add-Member -InputObject $ServerObject -MemberType NoteProperty -Name 'Connected' -Value $connected
Add-Member -InputObject $ServerObject -MemberType NoteProperty -Name 'Patched' -Value $Patched

$OutputArray += $ServerObject
  
}

$OutputArray | sort patched | FT -AutoSize

$OutputArray | sort patched | Export-Csv $filelocation -NoTypeInformation

send-mailmessage -From $Email -subject "Machines missing patch $PatchID " -To $Email -Attachments $filelocation -smtpserver $smtpserver 

}

else
{
""
"Invalid Selection"
""}



$repeat = Read-Host "Exit? (y/n)"


}
while($repeat -ne "y")
