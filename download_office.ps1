 Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $wshell=New-Object -ComObject wscript.shell
 Add-Type -AssemblyName System.Windows.Forms
  $checkdouble=(get-process cmd*).HandleCount.count
  $obj0=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv -Encoding UTF8

 if ($checkdouble -eq 1 -and (test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv")){

$link_build="https://partner.microsoft.com/en-us/dashboard/collaborate/packages"
Start-Process chrome.exe $link_build
start-sleep -s 180

######check abnormal login#######

start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait('^a')
start-sleep -s 2

[System.Windows.Forms.SendKeys]::SendWait('^c')
start-sleep -s 2
$webpage=Get-Clipboard

if($webpage -match "Unauthorized Verification"){

 $paramHash = @{
  #To =  "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-ProjectLead@allion.com","wallacelee@allion.com","taira.hayashi@allion.co.jp"
   To="shuningyu17120@allion.com.tw","ytliang@allion.com"
 from = 'Office_Download <edata_admin@allion.com>'
 Body = "Unauthorized Verification happened! Please check"
 Subject = "Office_Download_Alarm"

}


Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 

 (get-process "chrome" -ea SilentlyContinue).CloseMainWindow()

exit
}


######check if include survey question wording#######
if($webpage -match "How likely are you"){
[System.Windows.Forms.SendKeys]::SendWait('{TAB}')
start-sleep -s 5

[System.Windows.Forms.SendKeys]::SendWait('+{TAB}')
start-sleep -s 5

[System.Windows.Forms.SendKeys]::SendWait('~')
start-sleep -s 5
}



######check if include multi "next" wording#######

start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait('^a')
start-sleep -s 2

[System.Windows.Forms.SendKeys]::SendWait('^c')
start-sleep -s 2
$webpage=Get-Clipboard

$checknext_c=($webpage -match "next").count

if($checknext_c -ne 0){
$n=0
start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait('^f')
start-sleep -s 5

set-Clipboard -value "next"
[System.Windows.Forms.SendKeys]::SendWait('^v')

do{
[System.Windows.Forms.SendKeys]::SendWait('~')

$n++
}until($n -eq $checknext_c)

[System.Windows.Forms.SendKeys]::SendWait('{ESC}')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('{TAB}')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('{DOWN}{DOWN}')
start-sleep -s 30
[System.Windows.Forms.SendKeys]::SendWait('^a')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^c')
start-sleep -s 5
$webpage=Get-Clipboard

}


#$settings=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\settings.csv"

$ver1=$webpage -notmatch "Binaries" -notmatch "documents" -notmatch "Documents" | where {$_ -match "Published"}

$list_old=get-content  -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\lists.txt
$list_new= $ver1
$comp_list=((Compare-Object $list_old $list_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
 Write-Output "Line 469 comp_logs: comp_list"
<#
foreach($list_new1 in $list_new){
if($list_new1 -notin $list_old){
$comp_list+=$comp_list+@($list_new1)
}
if($list_new1 -in $list_old){
$list_new1
$comp_listsame+=$comp_listsame+@($list_new1)
}

}
#>


if($comp_list){

foreach($comp_list1 in $comp_list){

$topic=(($comp_list1-split '\t')[0]).trim()


###################### check next page ############

Set-Clipboard -value $topic

start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait('^f')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^v')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('{ESC}')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('~')
start-sleep -s 30

###Turn on ALL lists #####

start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait('^a')
start-sleep -s 2

[System.Windows.Forms.SendKeys]::SendWait('^c')
start-sleep -s 5
$webpage2=Get-Clipboard

if($webpage2 -match "51015All"){
$allc0=0
$anly=$webpage2 -match "ALL"

foreach($anl in $anly){
$allc0++
$anly2=0
$anly2= ($anly[0] -split "all").count -2
$anly2
if($anly2 -gt 0){
$anly=$allc0+$anly2
}
if($anlyx -match "51015All"){
$allc=$anly+$anly
}
}



Set-Clipboard -value "ALL"
$cc=0

start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait('^f')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^v')
start-sleep -s 2

if($allc -gt 1){
do{
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('~')
$cc++
}until($cc -eq $allc-1)
}

start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('{ESC}')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('~')
start-sleep -s 10

[System.Windows.Forms.SendKeys]::SendWait('^a')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^c')
start-sleep -s 5
$webpage20=Get-Clipboard
start-sleep -s 5
}

$pattern= "Description(.*?)Files in this Package"
$result_status  = [regex]::match($webpage20, "$pattern").Groups[1].Value
$desc=$result_status.replace(",","，")
$dlcount=($webpage20 -match "\/?\/?\/").count

start-sleep -s 2
Set-Clipboard -value "Last Modified"
start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait('^f')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^v')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('~')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('{ESC}')

$cc=0
do{
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('{tab}')

####start download####

$currentf=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\02.Application G\03.APP Build release\Microsoft Office_Auto Download\*").name

if($currentf.count -eq 0){$currentf="na"}
start-sleep -s 2

echo "Enter...downloading..."
start-sleep -s 5

$dl_p="\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\02.Application G\03.APP Build release\Microsoft Office_Auto Download\$fod\"
 #####################################  Downloading    #####################################

[System.Windows.Forms.SendKeys]::SendWait('~')

#check download complete

 do{
 start-sleep -s 60
 $check_ongoings =(Get-ChildItem -Path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\02.Application G\03.APP Build release\Microsoft Office_Auto Download\*.crdownload").count+(Get-ChildItem -Path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\02.Application G\03.APP Build release\Microsoft Office_Auto Download\*.tmp").count
 #$check_ongoings
  }until($check_ongoings -eq 0)
 
  start-sleep -s 5
 $currentf2=(gci -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\02.Application G\03.APP Build release\Microsoft Office_Auto Download\*").name
 $diff=((compare-object $currentf $currentf2)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject

 $topic1= (((($topic.replace("""","")).replace("(", "")). replace(")", "")).replace("[", "")).replace("]", "") 
 new-item -ItemType directory -path "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\02.Application G\03.APP Build release\Microsoft Office_Auto Download\$topic1\" -erroraction SilentlyContinue

  Move-Item "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\02.Application G\03.APP Build release\Microsoft Office_Auto Download\$diff" -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\02.Application G\03.APP Build release\Microsoft Office_Auto Download\$topic1\" -Force

  #####################################  Downloading    #####################################>
 
 $date_now=get-date -Format "M/d"
  $date_now2=get-date -Format "yyyyMd_HHmm"
 
  #add-content -path  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\logs.txt" -value "$fod, $topic, $select_dl, $date_now"
  copy-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv" "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_$($date_now2).csv" -Force

  "{0},{1},{2},{3},{4}" -f "","","","","" | add-content -path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv -force
   
  $writeto= import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv  -Encoding  UTF8
     $writeto[-1]."type"= "Office"
     $writeto[-1]."item1"= $topic
     $writeto[-1]."file"= $diff
     $writeto[-1]."date"= $date_now
     $writeto[-1]."desc"= $desc	

      $writeto| export-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv -Encoding  UTF8 -NoTypeInformation

$cc++
}until($cc -eq $dlcount)

###back to main page####

start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('%{LEFT}')

}
}

 $obj1=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv -Encoding UTF8
 
 $comp_logs=((Compare-Object $obj0."file" $obj1."file")|Where-Object { $_.SideIndicator -eq "=>"}).InputObject

 if($comp_logs){
 $i=0
 $obj3=@("","","","","")*100
  foreach( $comp_log in  $comp_logs){
   $obj2=$obj1|?{$_."file" -eq $comp_log }|select "type","item1","file","date"
   $obj3[$i]=$obj2
   $i=$i+1
   }
   $obj3=$obj3|?{$_.type -eq "office"}| ConvertTo-Html| Out-String
   

 add-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\lists.txt -value $comp_list

 #$new_logs=get-content -path   "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\logs.txt"
 #copy-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv" -Destination "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\02.Application G\03.APP Build release\Microsoft Office_Auto Download\packages_log.csv" -Force
remove-item -path  "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\02.Application G\03.APP Build release\Microsoft Office_Auto Download\packages_log.csv" -Force
 import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv"|?{$_.type -eq "office"}|export-csv "\\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\02.Application G\03.APP Build release\Microsoft Office_Auto Download\packages_log.csv" -Encoding UTF8 -NoTypeInformation

$end1="<BR>Check the Files: \\192.168.20.20\sto\EO\VD1\Dept-2\nec_tc\02.Application G\03.APP Build release\Microsoft Office_Auto Download"
$end2="<BR>Check the detail description from Query System: https://bu2-query.allion.com 【10_WinImg】<BR>"
$body="Please Check logs:<BR>"+'<font color="blue">'+$obj3+'</font>'+$end1+$end2
 $body= $body -replace  '<table>', '<table border="1">'
 $madd=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\ref\maillists.txt"
  $maillis= $maillis+@($madd)

 $paramHash = @{
 To = $maillis #"NPL-APP@allion.com"#,"NPL-DRV@allion.com","NPL-Preload@allion.com"
 #To="shuningyu17120@allion.com.tw"
 #Cc="shuningyu17120@allion.com.tw"
 from = 'Office_Download <edata_admin@allion.com>'
 BodyAsHtml = $True
 Subject = "<New Office Files> Please Check the download files logs (This is auto mail)"
 Body =$body
}

Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 
}
 (get-process "chrome" -ea SilentlyContinue).CloseMainWindow()

 ####header revised#####

 #copy-Item  -Path 'C:\Users\shuningyu17120\Desktop\csup_sum.csv' -Destination 'C:\Users\shuningyu17120\Desktop\Auto\Query\csup_sum.csv' -force

 $obj=Import-Csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv"

$col_counts=($obj | get-member -type NoteProperty).count

$header_1=  $null
$header_0=  $null
$di1=1

do {

$di2="{0:D2}" -f $di1
$header_1="Col_$di2"

if ($di1 -eq 1){
$header_0=$header_1
#$header_0
}


if ($di1 -gt 1 -and $di1 -le $col_counts){
$header_0=$header_0+","+$header_1
#$header_0
}

$di1++
}until ($di1-gt $col_counts) 

.{$header_0

 Get-Content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv"  | select -Skip 1}| Set-Content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_2.csv" -encoding utf8

$obj=Import-Csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_2.csv"



$header_3=  $null
$d1=$col_counts+1

do {

$d2="{0:D2}" -f $d1

$header_3= "Col_$d2"

$obj|Add-Member -MemberType NoteProperty -Name $header_3  -Value $null
$obj| Export-Csv -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_2.csv" -NoTypeInformation -encoding utf8

$d1++
}until ($d1 -gt 30) 



$obj=Import-Csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_2.csv"  -encoding utf8

 $heads=($obj | Select-Object -First 1).PSObject.Properties |  Select-Object -ExpandProperty Name

foreach($ob in $obj){
 foreach ( $head in  $heads){
 $len1=$ob.$head.length
 $cont1=$ob.$head
if( ($len1 -ne 0) -and ($cont1 -match ",")){
$ob.$head=$cont1 -replace ",","，"

}
}
}


$obj|Export-Csv -Path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_1.csv" -NoTypeInformation -encoding utf8
remove-item -path  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages_2.csv" -Force

 }


$date1= ((get-date).AddHours(-8)).Date
$date2= (get-date).Date
if($date1 -ne $date2){
Add-Content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\reboot_record.txt -Value $date2
shutdown /r /t 0
}