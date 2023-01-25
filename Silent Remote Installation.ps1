#Create a text file named Computers at root directory
New-Item C:\Computers.txt

#The variable $continue will be null upon declaration, the variable acts as an controller for the while loop
$continue
#The while loop is used for continuous input if there is more than one PC to setup
while($continue -ne 'N' -OR $continue -ne 'N')
{
    #Variable $pc is used to stores the PC's hostname
    $pc = Read-Host "Enter PC's hostname"
    #Value of $pc is then added to Computers.txt
    Add-Content C:\Computers.txt $pc
    <#Variable $continue will store input from user, if user enter 'N' or 'n', the loop will break because
    it does not meet the condition to continue the loop#>
    $continue = Read-Host "Do you want to continue? (Press N or n to stop)"
}

#Store the contents of Computers.txt into variable $computers
$computers = Get-Content "C:\Computers.txt"

#Compress the Software folder and transfer to the root directory of your PC
Compress-Archive -Path "\\sw\swdepot$\PowerShell\Software" -DestinationPath "C:\Software"
#Expand the Software folder to the root directory
Expand-Archive -LiteralPath "C:\Software.Zip" -DestinationPath "C:\"

#Stores the location of individual installer in a variable
$firefoxFile = "C:\Software\Firefox.exe"
$chromeFile = "C:\Software\Chrome.exe"
$adobeFile = "C:\Software\Adobe.exe"
$forticlientFile = "C:\Software\FortiClientVPN.exe"
$zoomFile = "C:\Software\Zoom.msi"
$vncFile = "C:\Software\UltraVNC.exe"
$vnciniFile = "C:\Software\UltraVNC.ini"
$7zipFile = "C:\Software\7-Zip.exe"
$javaFile = "C:\Software\Java.exe" 
$sophosFile = "C:\Software\Sophos.exe"
$samendpointFile = "C:\Software\SamEndPoint.msi"
$skypeFile = "C:\Software\Skype.exe"
$anydeskFile = "C:\Software\AnyDesk.exe"
$pe1File = "C:\Software\NEWPE1"
$pe3File = "C:\Software\NEWPE3"
$pe5File = "C:\Software\NEWPE5"
$pe6File = "C:\Software\NEWPE6"
$pe7File = "C:\Software\NEWPE7"

<#Foreach loop is used to remote install the 13 standard softwares onto each computer loop 
by loop#>
Foreach ($computer in $computers)
{
    <#Variable $pe is used to store the plant of the computer#>
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
            Compress-Archive -Path $pe1File -DestinationPath "\\$computer\c$\Users\Administrator\Desktop\NEWPE1" 
            Expand-Archive -LiteralPath "\\$computer\c$\Users\Administrator\Desktop\NEWPE1.Zip" -DestinationPath "\\$computer\c$\Users\Administrator\Desktop"
            Invoke-Command -ComputerName $computer -ScriptBlock{
                Start-Process -FilePath "C:\Users\Administrator\Desktop\NEWPE1\execPkg.exe" -Args "/silent /install" -Verb RunAs -Wait
                RM "C:\Users\Administrator\Desktop\NEWPE1.Zip"
            }   
            break
        }
        elseif ($pe -eq "PE3")
        {
            Compress-Archive -Path $pe3File -DestinationPath "\\$computer\c$\Users\Administrator\Desktop\NEWPE3" 
            Expand-Archive -LiteralPath "\\$computer\c$\Users\Administrator\Desktop\NEWPE3.Zip" -DestinationPath "\\$computer\c$\Users\Administrator\Desktop" 
            Invoke-Command -ComputerName $computer -ScriptBlock{
                Start-Process -FilePath "C:\Users\Administrator\Desktop\NEWPE3\execPkg.exe" -Args "/silent /install" -Verb RunAs -Wait
                RM "C:\Users\Administrator\Desktop\NEWPE3.Zip"
            } 
            break
        }
        elseif ($pe -eq "PE5")
        {
            Compress-Archive -Path $pe5File -DestinationPath "\\$computer\c$\Users\Administrator\Desktop\NEWPE5" 
            Expand-Archive -LiteralPath "\\$computer\c$\Users\Administrator\Desktop\NEWPE5.Zip" -DestinationPath "\\$computer\c$\Users\Administrator\Desktop"
            Invoke-Command -ComputerName $computer -ScriptBlock{
                Start-Process -FilePath "C:\Users\Administrator\Desktop\NEWPE5\execPkg.exe" -Args "/silent /install" -Verb RunAs -Wait
                RM "C:\Users\Administrator\Desktop\NEWPE5.Zip"
            }   
            break
        }
        elseif ($pe -eq "PE6")
        {
            Compress-Archive -Path $pe6File -DestinationPath "\\$computer\c$\Users\Administrator\Desktop\NEWPE6" 
            Expand-Archive -LiteralPath "\\$computer\c$\Users\Administrator\Desktop\NEWPE6.Zip" -DestinationPath "\\$computer\c$\Users\Administrator\Desktop" 
            Invoke-Command -ComputerName $computer -ScriptBlock{
                Start-Process -FilePath "C:\Users\Administrator\Desktop\NEWPE6\execPkg.exe" -Args "/silent /install" -Verb RunAs -Wait
                RM "C:\Users\Administrator\Desktop\NEWPE6.Zip"
            }  
            break
        }
        elseif ($pe -eq "PE7")
        {
            Compress-Archive -Path $pe7File -DestinationPath "\\$computer\c$\Users\Administrator\Desktop\NEWPE7" 
            Expand-Archive -LiteralPath "\\$computer\c$\Users\Administrator\Desktop\NEWPE7.Zip" -DestinationPath "\\$computer\c$\Users\Administrator\Desktop"
            Invoke-Command -ComputerName $computer -ScriptBlock{
                Start-Process -FilePath "C:\Users\Administrator\Desktop\NEWPE7\execPkg.exe" -Args "/silent /install" -Verb RunAs -Wait
                RM "C:\Users\Administrator\Desktop\NEWPE7.Zip"
            }   
            break
        }
        else
        {
            Write-Host "Invalid input. Please ensure input or spelling is correct."
            $pe = Read-Host "Which plant will this PC deployed to? (PE1, PE3, PE5, PE6, PE7)"
        }
    }

    #Copy each installer to the root directory of the remote computer
    Copy-Item $firefoxFile -Destination "\\$computer\c$\Firefox.exe"
    Copy-Item $chromeFile -Destination "\\$computer\c$\Chrome.exe"
    Copy-Item $adobeFile -Destination "\\$computer\c$\Adobe.exe"
    Copy-Item $forticlientFile -Destination "\\$computer\c$\FortiClientVPN.exe"
    Copy-Item $zoomFile -Destination "\\$computer\c$\Zoom.msi"
    Copy-Item $vncFile -Destination "\\$computer\c$\UltraVNC.exe"
    Copy-Item $vnciniFile -Destination "\\$computer\c$\UltraVNC.ini"
    Copy-Item $javaFile -Destination "\\$computer\c$\Java.exe"
    Copy-Item $7zipFile -Destination "\\$computer\c$\7-Zip.exe"
    Copy-Item $sophosFile -Destination "\\$computer\c$\Sophos.exe"
    Copy-Item $samendpointFile -Destination "\\$computer\c$\SamEndPoint.msi"
    Copy-Item $skypeFile -Destination "\\$computer\c$\Skype.exe"
    Copy-Item $anydeskFile -Destination "\\$computer\c$\AnyDesk.exe" 
    
    #Execute the installer to install the softwares silently on the remote computer      
    Invoke-Command -ComputerName $computer -ScriptBlock{
        Start-Process -FilePath "C:\Firefox.exe" -Args "/silent /install" -Verb RunAs -Wait 
        Start-Process -FilePath "C:\Chrome.exe" -Args "/silent /install" -Verb RunAs -Wait
        Start-Process -FilePath "C:\Adobe.exe" -Args "/silent /install" -Verb RunAs -Wait
        Start-Process -FilePath "C:\FortiClientVPN.exe" -Args "/quiet /norestart"
        msiexec /i "C:\Zoom.msi" /quiet /qn /norestart /log install.log
        Start-Process -FilePath "C:\UltraVNC.exe" -Args "/verysilent /norestart /loadinf=`"C:\UltraVNC.ini`"" -Wait -NoNewWindow
        Start-Process -FilePath "C:\Java.exe" -ArgumentList "/s" -Verb RunAs -Wait
        Start-Process -FilePath "C:\7-Zip.exe" -Args "/S /install" -Verb RunAs -Wait
        Start-Process -FilePath "C:\Sophos.exe" -ArgumentList "--quiet" -Verb RunAs -Wait
        msiexec /i "C:\SamEndPoint.msi" /quiet
        Start-Process -FilePath "C:\Skype.exe" -ArgumentList "/VERYSILENT /NORESTART /SUPPRESSMSGBOXEX /DL=1" -Verb RunAs -Wait
        Start-Process -FilePath "C:\AnyDesk.exe" -ArgumentList '--install "C:\Program Files (x86)" --start-with-win --create-shortcuts --create-desktop-icon --silent' -Wait
    
    #Remove the installer on the remote computer after installation
        RM C:\Firefox.exe 
        RM C:\Chrome.exe 
        RM C:\Adobe.exe
        RM C:\FortiClientVPN.exe 
        RM C:\Zoom.msi
        RM C:\UltraVNC.exe
        RM C:\UltraVNC.ini 
        RM C:\Java.exe 
        RM C:\7-Zip.exe 
        RM C:\Sophos.exe 
        RM C:\SamEndPoint.msi
        RM C:\Skype.exe
        RM C:\AnyDesk.exe   
    }
}

#Remove files and folders after multiple remote installation
RM C:\Computers.txt
RM C:\Software -Recurse
RM C:\Software.zip 