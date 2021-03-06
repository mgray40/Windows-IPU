﻿      ### Global ###
      $date = Get-Date
      $compArray = @() #Create array of zero or 1 object. Will be used to store results and export data
      $updatedArray = @() #create array of zero objects, will store completed update count
      $otherArray = @() #create array to store unknown builds
      $holdArray = @() #array to store machines on hold
      $Count_1709 = 0
      $Count_1803 = 0
       
       ### Function Definition ###
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
          #prompt user for file containing list of serials
          #$path = Read-Host -Prompt "Please enter the path to your file"
      
          $path = "C:\users\mgray40\desktop\1909script\dailyupdate.csv"
          #Read input file into variable $computers
          $computers = import-Csv -Path $path
          foreach ($computer in $computers) 
          {
              #create object
              $myObj = New-Object System.Object | Select "OSBuildID", "Name", "Technician", "Hold", "TASK", "User"
      
              #fill object
              #Try to find machine name in AD, if not found, catch by adding to LCM array
              try
              {
                 
                 ###$OU = get-ADComputer $computer.'Name' -Properties DistinguishedName | select @{label='OU';expression={$_.DistinguishedName.Split(',')[1].Split('=')[1]}}
                 $OU = get-ADComputer $computer.'Name' -Properties DistinguishedName | where {$_.DistinguishedName -like "*,OU=Retiring Computers*"}
      
      
                 #fill Name variable using get-adcomputer searching via $computer (computer serial string from input file) and selecting the name property only
                 $myObj.Name = Get-ADComputer $computer.'Name' -Properties Name | Select-Object -ExpandProperty Name 
      
                  #fill OS version build id searching get-adcomputer provided serialand selecting only OS build version
                  $myObj.OSBuildID = Get-ADComputer $computer.'Name' -Properties OperatingSystemVersion | Select-Object -ExpandProperty OperatingSystemVersion
              
                  #fill Name variable using get-adcomputer searching via $computer (computer serial string from input file) and selecting the name property only
                  $myObj.Technician = $computer.'Technician'
      
                  $myObj.Hold = $computer.'Hold'
      
                  $myObj.Task = $computer.'TASK'
      
                  $myObj.User = $computer.'User'
      
                  #if the variable OU is not empty (meaning it is in retiring computers), add to the LCM'd array
                  if ($OU -ne $null)
                  {
                      $Global:otherArray = $Global:otherArray + $myObj
                  }
                  else 
                  { 
                      #add object into proper array. If updated add to 'updatedArray' if not add to compArray
                      If ($myObj.OSBuildID -like "10.0 (18363)" -or $myObj.OSBuildID -like "10.0 (17763)" -or $myObj.OSBuildID -like "10.0 (19041)") 
                      {
                          $Global:updatedArray = $Global:updatedArray + $myObj 
                      }
                      If ($myObj.OSBuildID -like "10.0 (16299)" -or $myObj.OSBuildID -like "10.0 (17134)") 
                      {
                        $Global:compArray = $Global:compArray + $myObj
                        If ($myObj.Hold -ne "")
                        {
                            $Global:holdArray = $Global:holdArray + $myObj

                          #If ($myObj.Hold -like "True") 
                         # {
                          #$Global:holdArray = $Global:holdArray + $myObj
                          #}
                        }
                          If ($myObj.OSBuildID -like "10.0 (16299)")
                          {
                              $Global:Count_1709 ++
      
                          }
                          if ($myObj.OSBuildID -like "10.0 (17134)")
                          {
                              $Global:Count_1803 ++
                          }
                      } 
                  } 
                     
              }
              catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] 
              {
                  $myObj.Name = $computer.'Name' #if not in AD just assign the name from input file to object
                  $Global:otherArray = $Global:otherArray + $myObj 
              }
      
              #wipe object clear for next loop iteration. This avoids resource hogging/bloat
              $myObj = $null
          }
      }
      
      Function DisplaytoConsole()
      {
          #display array
          Write-Output `n "Number of computers successfully updated:" $updatedArray.Count
          Write-Output $Global:updatedArray | Format-Table -AutoSize -Property Name, OSBuildID, Technician, Task
      
          Write-Output "Number of Computers replaced by LCM:" $otherArray.Count
          Write-Output $Global:otherArray | Format-Table -AutoSize -Property Name
      
          Write-Output "Number of Computers needing updates:" $compArray.Count
          Write-Output $Global:compArray | where OSBuildID -like  "10.0 (16299)" | where Hold -Like "" | Format-Table -AutoSize
      
          Write-Output "Devices on 1709: "  $Global:Count_1709 
          Write-Output "Devices on 1803: "  $Global:Count_1803           
      
          Write-Output "Number of Computers on Hold:" $holdArray.Count
          Write-Output $Global:compArray | where OSBuildID -like  "10.0 (16299)" | where Hold -notLike "" | Format-Table -AutoSize
          #Write-Output $Global:holdArray | Format-Table -Property OSBuildID, Name, User -AutoSize 
      }
      
      Function DisplayObjects()
      {
          #Prompt user for output file path/name. Saved to global variable for email attachment later
          #$Global:outPath = Read-Host -Prompt "Please enter desired path/name for your output file"
          
          #Writes array values out to text file
          #$date | Out-File -FilePath $outPath 
      
          $outPath = "C:\Users\mgray40\desktop\1909script\1809 Update Report $(Get-date -f yyyy-MM-dd).txt"
          #Out-File -FilePath $outPath
      
          $date | Out-File -FilePath $outPath 
      
          $percent = ((($Global:updatedArray.Count + $Global:otherArray.Count)/($Global:updatedArray.Count + $Global:otherArray.Count + $Global:compArray.Count))*100)
          "Percent of machines completed: " + [math]::truncate($percent) +'%' | Out-File -FilePath $outPath  -Append
      
          'Number of computers successfully updated: ' + $Global:updatedArray.Count | Out-File -FilePath $outPath  -Append
          'Number of computers needing updates: ' + $Global:compArray.Count | Out-File -FilePath $outPath  -Append
          'Devices on 1709: ' + $Global:Count_1709 | Out-File -FilePath $outPath  -Append
          'Devices on 1803: ' + $Global:Count_1803 | Out-File -FilePath $outPath  -Append
          'Devices replaced by LCM: ' + $Global:otherArray.Count | Out-File -FilePath $outPath  -Append
          'Devices on Hold due to department: ' + $Global:holdArray.Count  + "`r`n" + "`r`n"  | Out-File -FilePath $outPath  -Append
      }
      
      Get-VPNStatus
      Get-BuildID
      #DisplayObjects
      DisplaytoConsole