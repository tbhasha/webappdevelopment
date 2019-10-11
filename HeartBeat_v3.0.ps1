########################################
# Tool   : HeartBeat                   #   
# Author : M. Dhayanidhi               #
# Version: 3.0                         #
########################################

    
#Check LogFile.
    
$UserFolerPath ="C:\Users\PALLATH"
$LoggedInUser = $env:USERNAME.ToUpper()
$file_exist=[System.IO.File]::Exists($UserFolerPath+"Desktop\test\Idle_log.txt")
if($file_exist)
{
    Remove-Item –path "$UserFolerPath\Desktop\test\Idle_log.txt"
}

#initializing the variables for file operations.

$today=Get-Date -UFormat "%Y-%m-%d"
$datePattern='\d\d\d\d-\d\d-\d\d \d\d:\d\d:\d\d'
$sum_Time=0
$file_from =$UserFolerPath+"\AppData\Roaming\Sapience\Log"
$files=Get-ChildItem "$UserFolerPath\AppData\Roaming\Sapience\Log" -Filter *.log
$file_dest="$UserFolerPath\Desktop\test"
$sapf=Get-ChildItem "$UserFolerPath\Desktop\test" -Filter *.log

#Loop to get the Sap ID from the log

for ($i=0; $i -lt $files.Count; $i++)
{
    robocopy  $file_from  $file_dest $files[$i]
}
 
for ($z=0; $z -lt $sapf.Count; $z++) 
{
    $infile=$sapf[$z].FullName
    $file=Get-Content $infile
    for($i=0; $i -lt $file.Length ;$i++)
    {
        if($file[$i] -match $today)
        {
            if($file[$i] -match 'Sapience Idle Message id =')
            {
                $id_index=$file[$i].IndexOf("Sapience Idle Message id =")

                $sap_str=$file[$i].Substring($id_index)
                $sap_len=$file[$i].Substring($id_index).Length

                $sap_id=$sap_str.Substring(26,($sap_len-26))
            }
        }
    }

    #Removing Spaces in Sap ID

    $sap_id=$sap_id.replace(' ' , '')
    $f_mat="Sapience Idle Message ($s_id)"
    $file1 = Get-Content $infile

   
    # ********** Get VM Inactive Hours Idle ********** Formula1

    for($i=0; $i -lt $file1.Length ;$i++)
    {
        if($file1[$i] -match $today)
        {
            if($file1[$i] -match "Sapience Idle Message \($sap_id\)")
            {
                echo $file1[$i] >> $UserFolerPath\Desktop\test\Idle_log.txt
                $id_index1=$file1[$i].IndexOf("Sapience Idle Message ($sap_id) StartTime:")
                $result=$file1[$i] | Select-String $datePattern -AllMatches
                $result=$result.Matches.Value
                $start_Time=$result[1]
                $end_Time=$result[2]
                $start_Time=[datetime]($start_Time)
                $end_Time=[datetime]($end_Time)
                
                $final_Time=$end_Time-$start_Time
                
                $sum_Time=$sum_Time+$final_Time
            }
        }

    }
}

# ********** Get VM Login Time ********** 

$myarray=@()
$datePattern='\d\d\d\d-\d\d-\d\d \d\d:\d\d:\d\d'
for ($z=0; $z -lt $sapf.Count; $z++) 
{
    $infile=$sapf[$z].FullName
    $file=Get-Content $infile
    for($j=0; $j -lt $file.Length ;$j++)
    {
        if($file[$j] -match $datePattern )
        {
            if($file[$j] -match $today)
            {
                $result=$file[$j] | Select-String $datePattern -AllMatches
               
                $date=[datetime]($result.Matches.Value)
                $myarray=$myarray+($date)
               
                break
            }
        }
    }
}

if(($myarray.Length) -gt 1)
{
    $myarr1=$myarray | Sort-Object
    $logon=$myarr1[0]
}
else
{
    $logon=$myarray[0]
}

# ********** Get VM Total Hours ********** Formula2
$CurrentTime= Get-Date

$VmTotHr = $CurrentTime - $logon 
$VmTotHrs = $VmTotHr
#$VmTotHr = $VmTotHr.TotalHours


# ********** Get VDI Active Hours ********** Formlua3

$VdiActvHr = $VmTotHr - $sum_Time 


# ********** Get VDI Expected Hours ********** Formlua4
$VdiExpected = 8 - ($VdiActvHr.TotalHours) 
$VdiExpected =  [timespan]::FromHours($VdiExpected)
$VdiExpected="{0:HH:mm:ss.fff}" -f ([datetime]$VdiExpected.Ticks)

#  ********** Display Data ********** 
function Get-TimeStamp {
    
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    
}

Write-Output $(Get-Date -Format dddd) | Out-File -FilePath C:\Users\pallath\Documents\VDIACTIVEHOURS.txt -NoClobber -Append
Write-Output "$(Get-TimeStamp) VM Logged in Time  : $logon"| Out-File -FilePath C:\Users\pallath\Documents\VDIACTIVEHOURS.txt -NoClobber -Append
Write-Output "$(Get-TimeStamp) VM Total Hours     : $VmTotHr"| Out-File -FilePath C:\Users\pallath\Documents\VDIACTIVEHOURS.txt -NoClobber -Append
Write-Output "$(Get-TimeStamp) VDI Active Hours   : $VdiActvHr"| Out-File -FilePath C:\Users\pallath\Documents\VDIACTIVEHOURS.txt -NoClobber -Append
Write-Output "$(Get-TimeStamp) VDI Inactive Hours : $sum_Time"| Out-File -FilePath C:\Users\pallath\Documents\VDIACTIVEHOURS.txt -NoClobber -Append
Write-Output "$(Get-TimeStamp) VDI Expected Hours : $VdiExpected"| Out-File -FilePath C:\Users\pallath\Documents\VDIACTIVEHOURS.txt -NoClobber -Append


cls
echo "*****************************************"
echo "          Welcome $LoggedInUser"
echo "*****************************************"
echo "VM Logged in Time  : $logon"
echo "VM Total Hours     : $VmTotHr"
echo "VDI Active Hours   : $VdiActvHr"
echo "VDI Inactive Hours : $sum_Time"
echo "VDI Expected Hours : $VdiExpected"
echo "*****************************************"
