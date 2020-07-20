#processTCP.ps1
#Get TCP Connected Processes of Local IP Address
#Sort Output by [1=OwningProcess][2=LocalPort][3=RemoteAddress][4=RemotePort][5=State][6=ProcessName][7=ServiceName][8=LocalAddress]
#Sort Order  by [1=Ascending][2=Descending]
#Filter IP Address by [1=IP][2=127.0.0.1][3=0.0.0.0][0=ALL]
#
#Run from cmd
#powershell .\processTCP.ps1                             Sort by Owning Process in Ascending  Order
#powershell .\processTCP.ps1 -sortorderno 2              Sort by Owning Process in Descending Order
#powershell .\processTCP.ps1 -sortorderno 2              Sort by Owning Process in Descending Order
#powershell .\processTCP.ps1 2                           Sort by Local Port     in Ascending  Order
#powershell .\processTCP.ps1 -sortno 3 -sortorderno 1    Sort by Remote Address in Ascending  Order
#powershell .\processTCP.ps1 -ipno 0                     Sort by Owning Process in Ascending  Order with ALL IP
#powershell .\processTCP.ps1 -sortno 5 -ipno 1           Sort by State          in Ascending  Order with 127.0.0.1 IP
#powershell .\processTCP.ps1 8 2 0                       Sort by Local Address  in Descending Order with ALL IP

#1st param = How will the output be sorted
#2nd param = Ascending or Descending
#3rd param = Filter by which IP Address
#Default param = Sorted by OwningProcess in Ascending Order
param($sortno = 1,$sortorderno = 1,$ipno = 1)

#Get Date MM/dd/yyyy HH:mm:ss 
$mm,$dd,$yy,$hh,$mn,$ss = (get-date -Format "MM/dd/yyyy/HH/mm/ss") -split "/"

#Set Location of Files
$dirpath = Get-Location
$fileloc = "D:\projects\batch\"
if($dirpath -ne $fileloc){
   $result = Test-Path $fileloc
   if($result){
      Set-Location $fileloc
      $dirpath = Get-Location
   }
}
write-host ("Current location of files: " + $dirpath)

#Set File Path
$procpath = "" + $dirpath + "\processTCP" + $yy + $mm + $dd + ".txt"
$procnow = "" + $dirpath + "\processTCPnow.txt"

#Create new temp file
Out-File -FilePath $procnow

#Get IP Address
$ip = Get-NetIPAddress -AddressFamily IPv4 -PrefixLength 24 -PrefixOrigin Dhcp | Select-Object -ExpandProperty IPAddress

#Set File Header
$fileheader = "" + $ip + " " + $mm + "/" + $dd + "/" + $yy + " " + $hh + ":" + $mn + ":" + $ss

#Set IP Address
if($ipno -eq 2){
   $ip = "127.0.0.1"
}elseif($ipno -eq 3){
   $ip = "0.0.0.0"
}elseif($ipno -eq 0){
   $ip = "*"
}

Write-Output ($fileheader) >> $procnow

#Set Sorted By
if($sortno -eq 2){
   $sortby = "LocalPort"
}elseif($sortno -eq 3){
   $sortby = "RemoteAddress"
}elseif($sortno -eq 4){
   $sortby = "RemotePort"
}elseif($sortno -eq 5){
   $sortby = "State"
}elseif($sortno -eq 6){
   $sortby = "Process"
}elseif($sortno -eq 7){
   $sortby = "Service"
}elseif($sortno -eq 8){
   $sortby = "LocalAddress"
}else{
   $sortby = "OwningProcess"
}

#Set Sorted Order
if($sortorderno -eq 2){
   $sortorder = @{Expression=$sortby;Descending=$true}
}else{
   $sortorder = $sortby
}

#Get ProcessID, Local Port, Remote Address and Port, State of TCP Connections
#$tcp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -eq $ip} | Where-Object{$_.RemoteAddress -ne "0.0.0.0"} | Sort-Object OwningProcess | Select-Object OwningProcess,LocalPort,RemoteAddress,RemotePort,State
if($sortno -eq 2){
   # Sort by Local Port and IP Address 
   if($ipno -ne 1){
      $tmp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      if($sortorderno -eq 2){
         $tcp = $tmp | Sort-Object $sortorder,@{Expression={$_.LocalAddress -as [VERSION]};Descending=$true}
         #$tcp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Sort-Object $sortorder,@{Expression={$_.LocalAddress -as [VERSION]};Descending=$true} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }else{
         $tcp = $tmp | Sort-Object $sortorder,{$_.LocalAddress -as [VERSION]}
         #$tcp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Sort-Object $sortorder,{$_.LocalAddress -as [VERSION]} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }
   }else{
      $tmp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      if($sortorderno -eq 2){
         $tcp = $tmp | Sort-Object $sortorder,@{Expression={$_.LocalAddress -as [VERSION]};Descending=$true}
         #$tcp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Sort-Object $sortorder,@{Expression={$_.LocalAddress -as [VERSION]};Descending=$true} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }else{
         $tcp = $tmp | Sort-Object $sortorder,{$_.LocalAddress -as [VERSION]} 
         #$tcp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Sort-Object $sortorder,{$_.LocalAddress -as [VERSION]} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }
   }
}elseif($sortno -eq 3){
   # Sort by Remote IP Address 
   if($ipno -ne 1){
      $tmp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      if($sortorderno -eq 2){
         $tcp = $tmp | Sort-Object @{Expression={$_.$sortby -as [VERSION]};Descending=$true}
         #$tcp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Sort-Object @{Expression={$_.$sortby -as [VERSION]};Descending=$true} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }else{
         $tcp = $tmp | Sort-Object {$_.$sortby -as [VERSION]}
         #$tcp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Sort-Object {$_.$sortby -as [VERSION]} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }
   }else{
      $tmp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      if($sortorderno -eq 2){
         $tcp = $tmp | Sort-Object @{Expression={$_.$sortby -as [VERSION]};Descending=$true}
         #$tcp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Sort-Object @{Expression={$_.$sortby -as [VERSION]};Descending=$true} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }else{
         $tcp = $tmp | Sort-Object {$_.$sortby -as [VERSION]}
         #$tcp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Sort-Object {$_.$sortby -as [VERSION]} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }
   }
}elseif($sortno -eq 4){
   # Sort by Remote Port and IP Address 
   if($ipno -ne 1){
      $tmp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      if($sortorderno -eq 2){
         $tcp = $tmp | Sort-Object $sortorder,@{Expression={$_.RemoteAddress -as [VERSION]};Descending=$true}
         #$tcp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Sort-Object $sortorder,@{Expression={$_.RemoteAddress -as [VERSION]};Descending=$true} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }else{
         $tcp = $tmp | Sort-Object $sortorder,{$_.RemoteAddress -as [VERSION]}
         #$tcp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Sort-Object $sortorder,{$_.RemoteAddress -as [VERSION]} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }
   }else{
      $tmp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      if($sortorderno -eq 2){
         $tcp = $tmp | Sort-Object $sortorder,@{Expression={$_.RemoteAddress -as [VERSION]};Descending=$true}
         #$tcp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Sort-Object $sortorder,@{Expression={$_.RemoteAddress -as [VERSION]};Descending=$true} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }else{
         $tcp = $tmp | Sort-Object $sortorder,{$_.RemoteAddress -as [VERSION]}
         #$tcp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Sort-Object $sortorder,{$_.RemoteAddress -as [VERSION]} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }
   }
}elseif($sortno -eq 6){
   # Sort by Process Name
   if($ipno -ne 1){
      $tcp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State,Process
   }else{
      $tcp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State,Process
   }
   Foreach($i in $tcp){
      $process = Get-Process -Id $i.OwningProcess | Select-Object -ExpandProperty ProcessName
      $i.Process = $process
   }
   $tcp = $tcp | Sort-Object $sortorder
}elseif($sortno -eq 7){
   # Sort by Service Name
   if($ipno -ne 1){
      $tcp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State,Service
   }else{
      $tcp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State,Service
   }
   Foreach($i in $tcp){
      $service = Get-WmiObject -Class Win32_Service | Where-Object {$_.processid -eq $i.OwningProcess} | select-Object -ExpandProperty name
      $i.Service = $service
   }
   $tcp = $tcp | Sort-Object $sortorder
}if($sortno -eq 8){
   # Sort by Local IP Address 
   if($ipno -ne 1){
      $tmp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      if($sortorderno -eq 2){
         $tcp = $tmp | Sort-Object @{Expression={$_.$sortby -as [VERSION]};Descending=$true}
         #$tcp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Sort-Object @{Expression={$_.$sortby -as [VERSION]};Descending=$true} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }else{
         $tcp = $tmp | Sort-Object {$_.$sortby -as [VERSION]}
         #$tcp = Get-NetTCPConnection | Where-Object{$_.LocalAddress -like $ip} | Sort-Object {$_.$sortby -as [VERSION]} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }
   }else{
      $tmp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      if($sortorderno -eq 2){
         $tcp = $tmp | Sort-Object @{Expression={$_.$sortby -as [VERSION]};Descending=$true}
         #$tcp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Sort-Object @{Expression={$_.$sortby -as [VERSION]};Descending=$true} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }else{
         $tcp = $tmp | Sort-Object {$_.$sortby -as [VERSION]}
         #$tcp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Sort-Object {$_.$sortby -as [VERSION]} | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
      }
   }
}else{
   # Sort by Owning Process
   if($ipno -ne 1){
      $tcp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip)} | Sort-Object $sortorder | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
   }else{
      $tcp = Get-NetTCPConnection | Where-Object{($_.LocalAddress -like $ip) -And ($_.RemoteAddress -ne "0.0.0.0") -And ($_.RemoteAddress -ne "::")} | Sort-Object $sortorder | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort,State
   }
}
$cnt = $tcp.length

#Write Header to file
Write-Output ("{0,-5} {1,-59} {2,-15}" -f " Pid", " ", "Process") >> $procnow
Write-Output (" {0,-4} {1,-15}:{2,-5} {3,-3} {4,-15}:{5,-5} {6,-11} {7,-15} {8,-21} {9,-4}" -f "Path", "Local Address", "Port", "TCP", "Remote Address", "Port", "State", "Version", "Service", "Name") >> $procnow
Write-Output ("{0,-5} {1,-15} {2,-5} {3,-3} {4,-15} {5,-5} {6,-11} {7,-15} {8,-21} {9,-5}" -f "-----", "--------------", "-----", "---", "--------------", "-----", "------", "-------", "--------", "-----") >> $procnow

$j=1
Foreach($i in $tcp){
   
   #Write Header to file after every 10 records
   $mod = $j % 21
   if($mod -eq 0){
      Write-Output ("`n{0,-5} {1,-59} {2,-15}" -f " Pid", " ", "Process") >> $procnow
      Write-Output (" {0,-4} {1,-15}:{2,-5} {3,-3} {4,-15}:{5,-5} {6,-11} {7,-15} {8,-21} {9,-4}" -f "Path", "Local Address", "Port", "TCP", "Remote Address", "Port", "State", "Version", "Service", "Name") >> $procnow
      Write-Output ("{0,-5} {1,-15} {2,-5} {3,-3} {4,-15} {5,-5} {6,-11} {7,-15} {8,-21} {9,-5}" -f "-----", "--------------", "-----", "---", "--------------", "-----", "------", "-------", "--------", "-----") >> $procnow
   }
   
   #Progress Bar
   if($cnt -ge 1){
      $k = ($j / $cnt) * 100
      $l = [math]::Round($k)
      Write-Progress -Activity "Searching Processes..." -Status "$l% Complete:" -PercentComplete $l
   }
   
   #Get Process Name, Path, Version from ProcessID
   $ErrorActionPreference = "Stop"
   try{
      $process = Get-Process -Id $i.OwningProcess | Select-Object ProcessName, Path, ProductVersion
      $proc = $process.ProcessName
      $path = $process.Path
      $ver = $process.ProductVersion
   }catch{
   }
   $ErrorActionPreference = "continue"
   
   $procid = $i.OwningProcess
   $laddr = $i.LocalAddress
   $lport = $i.LocalPort
   $raddr = $i.RemoteAddress
   $rport = $i.RemotePort
   $state = $i.State
   
   #Get Service Name, Path, Display
   $service = Get-WmiObject -Class Win32_Service | Where-Object {$_.processid -eq $procid} | select-Object name, pathname, displayname

   #Path in Service > Path in Process > Path in Registry
   if($null -ne $service){
      $name = ""
      $display = ""
      if($procid -ne 0){
         #Find Path in Service
         $pathtmp = $service.pathname -split " "
         Foreach($m in $pathtmp){
            if($null -ne $m){
               $path = $path + " " + $m
            }
         }
         $nametmp = $service.name -split " "
         Foreach($n in $nametmp){
            if($null -ne $n){
               $name = $name + " " + $n
            }
         }
         $displaytmp = $service.displayname -split " "
         Foreach($o in $displaytmp){
            if($null -ne $o){
               $display = $display + " " + $o
            }
         }
      }
   }else{
      if($null -eq $path){
         #Find Path in Registry
         $tempregpath = "HKCR\" + $proc + "\shell\open\command"
         $result = Test-Path $tempregpath
         if($result){
            $path = (Get-ItemProperty Registry::$tempregpath)."(Default)"
         }
      }
      $name = ""
      $display = ""
   }
   
   #Write Record to file
   if($j -eq 1){
      $prevpid = $procid
      $prevpath = $path
      $prevver = $ver
      Write-Output ("{0,-5} {1,-15}:{2,-5} TCP {3,-15}:{4,-5} {5,-11} {6,-14} {7,-21} {8}" -f $procid, $laddr, $lport, $raddr, $rport, $state, $proc, $name, $display) >> $procnow   
   }else{
      if($prevpid -ne $procid){
      
         if($null -ne $prevpath -Or $null -ne $prevver){
            Write-Output (" {0,-50} {1,-15}" -f $prevpath, $prevver) >> $procnow
         }
      
         Write-Output ("{0,-5} {1,-15}:{2,-5} TCP {3,-15}:{4,-5} {5,-11} {6,-14} {7,-21} {8}" -f $procid, $laddr, $lport, $raddr, $rport, $state, $proc, $name, $display) >> $procnow
         $prevpid = $procid
         $prevpath = $path
         $prevver = $ver
      }else{
         Write-Output ("{0,-5} {1,-15}:{2,-5} TCP {3,-15}:{4,-5} {5,-11} {6,-14} {7,-21} {8}" -f " ", $laddr, $lport, $raddr, $rport, $state, $proc, $name, $display) >> $procnow
      }
      if($j -eq $cnt){
         Write-Output (" {0,-50} {1,-15}" -f $path, $ver) >> $procnow
      }
   }
   $j++
}

#Write Output to File for Archive
$var = Get-Content $procnow
Write-Output ($var) >> $procpath
#Display Processes
Get-Content $procnow

#Read-Host -Prompt "Press Enter to continue"