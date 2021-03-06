### Global ###

#outdatedOS list, used to check against multiple old build ID's at once
#UPDATE AS NEEDED IN FORMAT "10.0 (XXXXX)"
$script:OutOfDateOS = ("10.0 (16299)", "10.0 (17134)")
#UpdatedOS list, used to check against multiple "acceptable" build ID's at once 
#UPDATE AS NEEDED IN FORMAT "10.0 (XXXXX)"
$script:UpdatedOS = ("10.0 (18363)", "10.0 (19041)")
$script:compArray = @()
$script:1709_Count = 0
$script:1803_Count = 0
$script:LCM_Count = 0
$script:No_AD_Count = 0
$script:Updated_Count = 0
$script:Outdated_Count = 0
$script:Percent = 0

## Function Definitions ## 
Function Get-VPNStatus()
{
    $netAdapter = Get-NetAdapter -InterfaceDescription *AnyConnect* | Select-Object -ExpandProperty status
    While ($netAdapter -ne "Up") {
        Write-Host -ForegroundColor Yellow "* Connection failed *"
        Start-Process -FilePath "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"
        Read-Host "AnyConnect VPN will now open to verify login. Press 'ENTER' here to try again"
        $netAdapter  =  Get-NetAdapter -InterfaceDescription *AnyConnect* | Select-Object -ExpandProperty status
    }  
    Write-Host -ForegroundColor Green "* Connection to VPN successful *" 
}

Function Get-BuildID() 
{
    $path = "C:\Users\mgray40\Documents\Powershell Scripts\ExcelTest\ExcelDaily.xlsx"
    $computers = import-Excel -Path $path
    foreach ($computer in $computers)
    {
        #create object
        $CompObj = New-Object System.Object | Select-Object "OSBuildID", "Name", "Technician", "Status", "TASK", "User"

        #fill object
        #Try to find machine name in AD, if not found, catch by adding to LCM array
        try
        {
            #fill Name variable using get-adcomputer searching via $computer (computer serial string from input file) and selecting the name property only
            $CompObj.Name = Get-ADComputer $computer.'Name' -Properties Name | Select-Object -ExpandProperty Name
            
            #fill OS version build id searching get-adcomputer provided serialand selecting only OS build version
            $CompObj.OSBuildID = Get-ADComputer $computer.'Name' -Properties OperatingSystemVersion | Select-Object -ExpandProperty OperatingSystemVersion
              
            #Fill with remaining columns from input file
            $CompObj.Technician = $computer.'Technician'      
            $CompObj.Status = $computer.'Flag'      
            $CompObj.Task = $computer.'TASK'      
            $CompObj.User = $computer.'User'

            #Check machine OU for retiring Computers
            $OU = get-ADComputer $computer.'Name' -Properties DistinguishedName | Where-Object {$_.DistinguishedName -like "*,OU=Retiring Computers*"}
      
            #if the variable OU is not empty (meaning it is in retiring computers), add to the LCM'd array
            if ($null -ne $OU)
            {
                $CompObj.OSBuildID = "LCM"
            }
            #If machine was updated to acceptable OS - update the flag
            if ($CompObj.OSBuildID -in $script:UpdatedOS)
            {
                $CompObj.Status = "Updated"
                $script:Updated_Count++
            }
            if ($CompObj.OSBuildID -in $script:OutOfDateOS)
            {
                $script:Outdated_Count++
            }

            $script:CompArray = $script:CompArray + $CompObj
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
        {            
            #If device is not in AD or can't be found Assign object name from list, mark the build ID as not valid and add to list
            $CompObj.Name = $computer.'Name' #if not in AD just assign the name from input file to object
            $CompObj.OSBuildID = "N/A"
            $script:CompArray = $script:CompArray + $CompObj
        }

        #wipe object clear for next loop iteration. This avoids resource hogging/bloat
        $CompObj = $null
    }

    Get-Counts
}

Function Get-Counts()
{
    #count variables - These 3 should be correct/stable for all Win 10 OS builds
    $script:LCM_Count = ($script:compArray | Where-Object {$_.OSBuildID -like "LCM"}).Count
    $script:No_AD_Count = ($script:compArray | Where-Object {$_.OSBuildID -like "N/A"}).Count
    #$script:Updated_Count = ($script:compArray | Where-Object {$_.Status -like "Updated"}).Count
    #The two below are current project. Can be removed or used as template for future project
    $script:1709_Count = ($script:compArray | Where-Object {$_.OSBuildID -like "10.0 (16299)"}).Count
    $script:1803_Count = ($script:compArray | Where-Object {$_.OSBuildID -like "10.0 (17134)"}).Count
}
Function DisplayToConsole()
{
    #Machines Updated!
    "`nMachines Successfully Updated: " + ($script:Updated_Count + $script:LCM_Count)
    #Write-Output $script:compArray | Where-Object Flag -like "Updated" -or Where-Object OSBuildID -like "LCM" | Format-Table -AutoSize
    Write-Output ($script:compArray | Where-Object {$_.Status -like "Updated" -or $_.OSBuildID -like "LCM"}) | Format-Table -AutoSize
    
    
    #Machines needing Updates: 
    "`nMachines Needing Updates: " + ($script:Outdated_Count)

    "`nMachines on Hold: " + ($script:compArray | Where-Object {$_.Status -like "true"}).Count
    
    #1709 output
    "`nNumber of Machines on 1709: " + $1709_Count 

    #Write-Output "Machines on 1709: "
    #write-output $script:compArray | Where-Object OSBuildID -like "10.0 (16299)" | Format-Table -AutoSize
    
    #1803 output
    "`nNumber of Machines on 1803: " + $1803_Count
    #" `n Machines on 1803"
    #write-output $script:compArray | Where-Object OSBuildID -like "10.0 (17134)" | Format-Table -AutoSize

    #LCM and Missing from AD
    "`nMachines Replaced by LCM: " + $script:LCM_Count
    #write-output $script:compArray | Where-Object OSBuildID -like "LCM" | Format-Table -AutoSize

    "`nMachines not Found in AD: " + $script:No_AD_Count
    write-output $script:compArray | Where-Object OSBuildID -like "N/A" | Format-Table -AutoSize
}

Function ExportReport()
{
    #enter code here
    $xlfile = "C:\Users\mgray40\desktop\Daily Report $(Get-date -f yyyy-MM-dd).xlsx"
    #$xlfile = "$env:TEMP\Report $(Get-date -f yyyy-MM-dd).xlsx"
    #Remove-Item $xlfile -ErrorAction SilentlyContinue

    Get-Date | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 2
    "Machines Updated: " + ($script:Updated_Count + $script:LCM_Count + $script:No_AD_Count) | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 3
    "Percent of machines completed: " + [math]::truncate($script:Percent) +'%' | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 4
    "Machines Needing Updates: " + $script:Outdated_Count  | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 5
    "Number of Machines on 1709: " + $1709_Count | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 6
    "Number of Machines on 1803: " + $1803_Count | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 7
    "Number of Machines on Hold: " + ($script:compArray | Where-Object {$_.Flag -like "true"}).Count | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 8
    $Sheet1 = "Windows IPU Report" | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 1 -PassThru
    $ws = $Sheet1.Workbook.Worksheets['Summary']
    $formatType1 = @{WorkSheet=$ws;Bold=$true;FontSize=18;}
    Set-Format -Range A1 -Value "Windows IPU Report" @formatType1

    $formatType2 = @{WorkSheet=$ws;AutoSize=$true}
    Set-Format -Range A  @formatType2

    Close-ExcelPackage $Sheet1 
    
    $Sheet2 =  $script:compArray | Where-Object OSBuildID -like "10.0 (16299)" | Export-Excel $xlfile -AutoSize -WorksheetName 1709 -Numberformat 'Text' -NoNumberConversion * -startrow 1 -TableName Comp1709 -PassThru
    Close-ExcelPackage $Sheet2 

    $Sheet3 =  $script:compArray | Where-Object OSBuildID -like "10.0 (17134)" | Export-Excel $xlfile -AutoSize -WorksheetName 1803 -Numberformat 'Text' -NoNumberConversion * -startrow 1 -TableName Comp1803 -PassThru
    Close-ExcelPackage $Sheet3 

    $Sheet4 =  $script:compArray | Where-Object OSBuildID -like "LCM" | Export-Excel $xlfile -AutoSize -WorksheetName LCM -Numberformat 'Text' -NoNumberConversion * -startrow 1 -TableName CompLCM -PassThru
    Close-ExcelPackage $Sheet4 

    $Sheet5 =  $script:compArray | Where-Object OSBuildID -like "N/A" | Export-Excel $xlfile -AutoSize -WorksheetName Not_in_AD -Numberformat 'Text' -NoNumberConversion * -startrow 1 -TableName CompAD -PassThru
    Close-ExcelPackage $Sheet5 -Show

    #Call email function to send report to supervisors
    #EmailReport
}

Function EmailReport()
{
    Write-Host -ForegroundColor Red "VPN ADAPTER IS NOW BEING DISABLED. PLEASE WAIT..."
    Start-Sleep 3
    Get-NetAdapter -InterfaceDescription "Cisco AnyConnect Secure Mobility Client Virtual Miniport Adapter for Windows x64" | Disable-NetAdapter -Verbose -Confirm:$false
    Start-Sleep 3
    Send-MailMessage -SmtpServer mail.liberty.edu -From mgray40@liberty.edu -To mgray40@liberty.edu -Subject "Windows IPU Project - Daily Progress" -Body "$date Update report attached" -Attachments $xlfile
}

## Main ##
Get-VPNStatus
Get-BuildID
DisplayToConsole
#ExportReport