#Compress the Software folder and transfer to the root directory of your PC
Compress-Archive -Path "\\sw\swdepot$\PowerShell\Software" -DestinationPath "C:\Software"
#Expand the Software folder to the root directory
Expand-Archive -LiteralPath "C:\Software.Zip" -DestinationPath "C:\"

#Variable $pe is used to store the plant of the computer
$pe = Read-Host "Which plant will this PC deployed to? (PE1, PE3, PE5, PE6, PE7)"
<#The while loop technically will loop forever, but will stop by the break statement
in each if/elseif statement, if there is no input meets the condition for the if/elseif statement,
the code in the else statement will be executed, which is asking for input again, 
restarting the loop#>
While ($pe)
{
    <#The first line of code in the if/elseif statement will compress respective printer driver folder and transfer to
    the remote computer, the second line which will unzip the folder on the local Administrator's Desktop, the third
    line will execute the installer, installing the printer driver software silently on the remote installer and
    remove the unwanted zip file after installation#>
    if ($pe -eq "PE1")
    {
        Compress-Archive C:\Software\NEWPE1 -DestinationPath "C:\Users\Administrator\Desktop\NEWPE1"
        Expand-Archive -LiteralPath "C:\Users\Administrator\Desktop\NEWPE1.Zip" -DestinationPath "C:\Users\Administrator\Desktop"
        Start-Process C:\Users\Administrator\Desktop\NEWPE1\execPkg.exe -ArgumentList "/quiet" -Wait
        RM C:\Users\Administrator\Desktop\NEWPE1.Zip
        break
     }
     elseif ($pe -eq "PE3")
     {
        Compress-Archive C:\Software\NEWPE3 -DestinationPath "C:\Users\Administrator\Desktop\NEWPE3"
        Expand-Archive -LiteralPath "C:\Users\Administrator\Desktop\NEWPE3.Zip" -DestinationPath "C:\Users\Administrator\Desktop"
        Start-Process C:\Users\Administrator\Desktop\NEWPE3\execPkg.exe -ArgumentList "/quiet" -Wait
        RM C:\Users\Administrator\Desktop\NEWPE3.Zip
        break
     }
     elseif ($pe -eq "PE5")
     {
        Compress-Archive C:\Software\NEWPE5 -DestinationPath "C:\Users\Administrator\Desktop\NEWPE5"
        Expand-Archive -LiteralPath "C:\Users\Administrator\Desktop\NEWPE5.Zip" -DestinationPath "C:\Users\Administrator\Desktop"
        Start-Process C:\Users\Administrator\Desktop\NEWPE5\execPkg.exe -ArgumentList "/quiet" -Wait
        RM C:\Users\Administrator\Desktop\NEWPE5.Zip
        break
     }
     elseif ($pe -eq "PE6")
     {
        Compress-Archive C:\Software\NEWPE6 -DestinationPath "C:\Users\Administrator\Desktop\NEWPE6"
        Expand-Archive -LiteralPath "C:\Users\Administrator\Desktop\NEWPE6.Zip" -DestinationPath "C:\Users\Administrator\Desktop"
        Start-Process C:\Users\Administrator\Desktop\NEWPE6\execPkg.exe -ArgumentList "/quiet" -Wait
        RM C:\Users\Administrator\Desktop\NEWPE6.Zip
        break
     }
     elseif ($pe -eq "PE7")
     {
        Compress-Archive C:\Software\NEWPE7 -DestinationPath "C:\Users\Administrator\Desktop\NEWPE7"
        Expand-Archive -LiteralPath "C:\Users\Administrator\Desktop\NEWPE7.Zip" -DestinationPath "C:\Users\Administrator\Desktop"
        Start-Process C:\Users\Administrator\Desktop\NEWPE7\execPkg.exe -ArgumentList "/quiet" -Wait
        RM C:\Users\Administrator\Desktop\NEWPE7.Zip
        break
     }
     else
     {
        Write-Host "Invalid input. Please ensure input or spelling is correct."
        $pe = Read-Host "Which plant will this PC deployed to? (PE1, PE3, PE5, PE6, PE7)"
     }
}

#Execute the installer to install the softwares silently on the computer
Start-Process C:\Software\Sophos.exe -ArgumentList "--quiet"        
Start-Process C:\Software\Firefox.exe -ArgumentList "/silent /install" -Wait
Start-Process C:\Software\Chrome.exe -ArgumentList  "/silent /install" -Wait
Start-Process C:\Software\Adobe.exe -ArgumentList "/sAll /rs" -Wait
Start-Process msiexec.exe -ArgumentList "/i C:\Software\FortiClientVPN.msi REBOOT=ReallySuppress /qn" -Wait
Start-Process msiexec.exe -ArgumentList "/i C:\Software\Zoom.msi /quiet /norestart" -Wait
Start-Process C:\Software\UltraVNC.exe -ArgumentList "/verysilent /norestart /loadinf=`"C:\Software\UltraVNC.ini`"" -Wait
Start-Process C:\Software\Java.exe -ArgumentList "/s" -Wait 
Start-Process C:\Software\7-Zip.exe -ArgumentList "/S" -Wait
Start-Process C:\Software\Skype.exe -ArgumentList "/VERYSILENT /NORESTART /SUPPRESSMSGBOXEX /DL=1" -Wait
Start-Process C:\Software\AnyDesk.exe -ArgumentList '--install "C:\Program Files (x86)" --start-with-win --create-shortcuts --create-desktop-icon --silent' -Wait
Start-Process msiexec.exe -ArgumentList "/i C:\Software\SamEndPoint.msi /quiet /norestart" -Wait

<#There are times where Sophos had been installed, but if you take a look in task manager,
the installer is still running but its CPU usage is 0%, this code snippet is used to overcome
this issue#>
#Check if Sophos is installed, check if the installer is still running, if it does, kill the process
$sophosInstalled = Test-Path -Path "C:\Program Files\Sophos"
if($sophosInstalled)
{
    $sophosProcess = Get-Process Sophos -Erroraction SilentlyContinue
    if($sophosProcess)
    {
        Stop-Process -Name Sophos -Force
    }
}
#Remove files and folders after installation
RM C:\Software -Recurse
RM C:\Software.zip