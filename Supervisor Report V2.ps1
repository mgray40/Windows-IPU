### Global ###

#outdatedOS list, used to check against multiple old build ID's at once
#UPDATE AS NEEDED IN FORMAT "10.0 (XXXXX)"
$script:OutOfDateOS = ("10.0 (16299)", "10.0 (17134)", "10.0 (17763)" , "10.0 (18362)")
#UpdatedOS list, used to check against multiple "acceptable" build ID's at once 
#UPDATE AS NEEDED IN FORMAT "10.0 (XXXXX)"
$script:UpdatedOS = ("10.0 (18363)", "10.0 (19041)")
$script:compArray = @()
$script:1709_Count = 0
$script:1803_Count = 0
$script:1903_Count = 0
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
    $path = "C:\Users\mgray40\Documents\Powershell Scripts\ExcelTest\ExcelMaster.xlsx"
    $computers = import-Excel -Path $path
    foreach ($computer in $computers)
    {
        #create object
        $CompObj = New-Object System.Object | Select-Object "OSBuildID", "Name", "Technician", "Flag", "TASK", "User"

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
            $CompObj.Flag = $computer.'Flag'      
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
            if ($CompObj.OSBuildID -in $UpdatedOS)
            {
                $CompObj.Flag = "Updated"
            }
            if ($CompObj.OSBuildID -in $OutOfDateOS)
            {
                $script:Outdated_Count++
                if($null -eq $CompObj.Flag) 
                {
                    $CompObj.Flag = "Outdated"
                }
            }

            $script:CompArray = $script:CompArray + $CompObj
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
        {            
            #If device is not in AD or can't be found Assign object name from list, mark the build ID as not valid and add to list
            $CompObj.Name = $computer.'Name' #if not in AD just assign the name from input file to object
            $CompObj.OSBuildID = "N/A"
            $CompObj.Task = $computer.'TASK'
            $CompObj.Flag = $computer.'Flag'
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
    $script:Updated_Count = ($script:compArray | Where-Object {$_.Flag -like "Updated"}).Count
    $script:Percent =  ((($script:Updated_Count + $script:LCM_Count + $script:No_AD_Count)/$compArray.count)*100)
    #The two below are current project. Can be removed or used as template for future project
    $script:1709_Count = ($script:compArray | Where-Object {$_.OSBuildID -like "10.0 (16299)"}).Count
    $script:1803_Count = ($script:compArray | Where-Object {$_.OSBuildID -like "10.0 (17134)"}).Count
    $script:1903_Count = ($script:compArray | Where-Object {$_.OSBuildID -like "10.0 (18362)"}).Count
}
Function DisplayToConsole()
{
    #Machines Updated!
    "`nMachines Successfully Updated: " + ($script:Updated_Count + $script:LCM_Count)
    #Write-Output $script:compArray | Where-Object Flag -like "Updated" | Format-Table -AutoSize
    
    
    #Machines needing Updates: 
    "`nMachines Needing Updates: " + ($script:Outdated_Count)

    "`nMachines on Hold: " + ($script:compArray | Where-Object {$_.Flag -like "true"}).Count
    
    #1709 output
    "`nNumber of Machines on 1709: " + $1709_Count 

    #Write-Output "Machines on 1709: "
    #write-output $script:compArray | Where-Object OSBuildID -like "10.0 (16299)" | Format-Table -AutoSize
    
    #1803 output
    "`nNumber of Machines on 1803: " + $1803_Count
    #" `n Machines on 1803"
    #write-output $script:compArray | Where-Object OSBuildID -like "10.0 (17134)" | Format-Table -AutoSize

    #1903 output
    "`nNumber of Machines on 1903: " + $1903_Count
    #" `n Machines on 1903"
    #write-output $script:compArray | Where-Object OSBuildID -like "10.0 (18362)" | Format-Table -AutoSize

    #LCM and Missing from AD
    "`nMachines Replaced by LCM: " + $script:LCM_Count
    #write-output $script:compArray | Where-Object OSBuildID -like "LCM" | Format-Table -AutoSize

    "`nMachines not Found in AD: " + $script:No_AD_Count
    #write-output $script:compArray | Where-Object OSBuildID -like "N/A" | Format-Table -AutoSize

    "`nPercent of machines completed: " + [math]::truncate($script:Percent) +'%'
}

Function ExportReport()
{
    #Remove-Item "C:\Users\mgray40\desktop\IPU Report $(Get-date -f yyyy-MM-dd).xlsx"
    #create excel file
    $xlfile = "C:\Users\mgray40\desktop\IPU Report $(Get-date -f yyyy-MM-dd).xlsx"

    foreach ($comp in $script:CompArray) {
        if ($comp.OSBuildID -like "10.0 (16299)")
        {
            $comp.OSBuildID = "1709"
        }
        elseif ($comp.OSBuildID -like "10.0 (17134)")
        {
            $comp.OSBuildID = "1803"
        }
        elseif ($comp.OSBuildID -like "10.0 (17763)")
        {
            $comp.OSBuildID = "1809"
        }
        elseif ($comp.OSBuildID -like "10.0 (18362)")
        {
            $comp.OSBuildID = "1903"
        }
        elseif ($comp.OSBuildID -like "10.0 (18363)")
        {
            $comp.OSBuildID = "1909"
        }
        elseif ($comp.OSBuildID -like "10.0 (19041)")
        {
            $comp.OSBuildID = "2004"
        }
    }
    
    #write placeholder data, this will get overwritten. It's just to create the first worksheet
    "placeholder" | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 1 #-PassThru
    
    $Sheet2 =  $script:compArray | Where-Object OSBuildID -like "1709" | Export-Excel $xlfile -AutoSize -WorksheetName 1709 -Numberformat 'Text' -NoNumberConversion * -startrow 1 -TableName Comp1709 -PassThru
    Close-ExcelPackage $Sheet2 

    $Sheet3 =  $script:compArray | Where-Object OSBuildID -like "1803" | Export-Excel $xlfile -AutoSize -WorksheetName 1803 -Numberformat 'Text' -NoNumberConversion * -startrow 1 -TableName Comp1803 -PassThru
    Close-ExcelPackage $Sheet3 

    $Sheet4 =  $script:compArray | Where-Object OSBuildID -like "1903" | Export-Excel $xlfile -AutoSize -WorksheetName 1903 -Numberformat 'Text' -NoNumberConversion * -startrow 1 -TableName Comp1903 -PassThru
    Close-ExcelPackage $Sheet4 

    $Sheet5 =  $script:compArray | Where-Object OSBuildID -like "LCM" | Export-Excel $xlfile -AutoSize -WorksheetName LCM -Numberformat 'Text' -NoNumberConversion * -startrow 1 -TableName CompLCM -PassThru
    Close-ExcelPackage $Sheet5 

    $Sheet6 =  $script:compArray | Where-Object OSBuildID -like "N/A" | Export-Excel $xlfile -AutoSize -WorksheetName Not_in_AD -Numberformat 'Text' -NoNumberConversion * -startrow 1 -TableName CompAD -PassThru
    Close-ExcelPackage $Sheet6 

    $Sheet7 =  $script:compArray | Where-Object Flag -like "Enrollment" | Export-Excel $xlfile -AutoSize -WorksheetName Enrollment -Numberformat 'Text' -NoNumberConversion * -startrow 1 -TableName CompHold -PassThru
    Close-ExcelPackage $Sheet7 #-Show

    $Sheet8 =  $script:compArray | Export-Excel $xlfile -AutoSize -WorksheetName Master -NoNumberConversion * -startrow 1 -TableName MasterTable -PassThru #-IncludePivottable  -PivotRows OSBuildID -PivotData @{"OSBuildID"="COUNT"} -PivotColumns Flag #-IncludePivotChart -PivotChartType Pie -ShowPercent
    
    Close-ExcelPackage $Sheet8 
    
    $Sheet1 ="Machines Updated: " + ($script:Updated_Count + $script:LCM_Count + $script:No_AD_Count) | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 3
    $Sheet1 ="Percent of machines completed: " + [math]::truncate($script:Percent) +'%' | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 4
    $Sheet1 ="Machines Needing Updates: " + $script:Outdated_Count  | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 5
    $Sheet1 ="Number of Machines on 1709: " + $1709_Count | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 6
    $Sheet1 ="Number of Machines on 1803: " + $1803_Count | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 7
    $Sheet1 ="Number of Machines on 1903: " + $1903_Count | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 8
    #$Sheet1 ="Number of Machines on Hold: " + ($script:compArray | Where-Object {$_.Flag -like "true"}).Count | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 9
    $Sheet1 = Get-Date | Export-Excel $xlfile -AutoSize -WorksheetName Summary -startrow 2 -PassThru
    $ws = $sheet1.workbook.worksheets['Summary']
    $formatType1 = @{WorkSheet=$ws;Bold=$true;FontSize=18;}
    Set-Format -Range A1 -Value "Windows IPU Report" @formatType1

    #adds rules for pivot table 
    $pivotTableParams = @{
        PivotTableName  = "TestTable"
        Address         = $sheet1.Summary.cells["C1"]
        SourceWorkSheet = $sheet1.Master
        PivotRows       = @("OSBuildID")
        PivotData       = @{"OSBuildID"="COUNT"}
        PivotColumns    = @("Flag")
        
        PivotChartDefinition =@{
            Title="IPU Pivot Chart"
            ChartType = 'ColumnStacked'
            Column = 9
        }        
    }
    #declares pivot table
    Add-PivotTable @pivotTableParams -PassThru

    #writes new format over range A1:8
    $formatType2 = @{WorkSheet=$ws;AutoSize=$true}
    Set-Format -Range A1:A8  @formatType2

    Close-ExcelPackage $Sheet1 #-Show
    #Call email function to send report to supervisors
    EmailReport
}

Function EmailReport()
{
    Write-Host -ForegroundColor Red "VPN ADAPTER IS NOW BEING DISABLED. PLEASE WAIT..."
    Start-Sleep 3
    Get-NetAdapter -InterfaceDescription "Cisco AnyConnect Secure Mobility Client Virtual Miniport Adapter for Windows x64" | Disable-NetAdapter -Verbose -Confirm:$false
    Start-Sleep 3
    Send-MailMessage -SmtpServer mail.liberty.edu -From mgray40@liberty.edu -To mgray40@liberty.edu, cseavers@liberty.edu, sebrooks@liberty.edu, bjday@liberty.edu, jdclemons@liberty.edu -Subject "Windows IPU Project - Daily Progress" -Body "$date Update report attached." -Attachments $xlfile
}

## Main ##
Get-VPNStatus
Get-BuildID
DisplayToConsole
ExportReport
