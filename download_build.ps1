 Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
 $wshell=New-Object -ComObject wscript.shell
 Add-Type -AssemblyName System.Windows.Forms
  $checkdouble=(get-process cmd*).HandleCount.count
  $obj0=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv -Encoding UTF8

 if ($checkdouble -eq 1 -and (test-path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\lists.txt")){

$link_build="https://partner.microsoft.com/en-us/dashboard/collaborate/packages"
Start-Process msedge.exe $link_build
#start-sleep -s 180
start-sleep -s 60

######check abnormal login#######

start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait('^a')
start-sleep -s 2

[System.Windows.Forms.SendKeys]::SendWait('^c')
start-sleep -s 2
$webpage=Get-Clipboard


######check need Verify your identity #######

if($webpage -match "Verify your identity"){

 $paramHash = @{
 To="ytliang@allion.com"
 Cc="shuningyu17120@allion.com.tw"
 from = 'Win_Download <edata_admin@allion.com>'
 BodyAsHtml = $True
 Subject = "<Verify your identity> Pleas help verify Microsoft collaborator website  (This is auto mail)"
 Body ="使用AnyDesk遠端  ID/password: 177586082/Cocoon-SIT-014"

}


Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 


 (get-process "msedge" -ea SilentlyContinue).CloseMainWindow()

exit

}

if($webpage -match "Unauthorized Verification" -or $webpage -match "Verify your identity"){

 $paramHash = @{
  #To =  "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-ProjectLead@allion.com","wallacelee@allion.com","taira.hayashi@allion.co.jp"
   To="shuningyu17120@allion.com.tw","ytliang@allion.com"
 from = 'Win_Download <edata_admin@allion.com>'
 Body = "Unauthorized Verification happened! Please check"
 Subject = "使用AnyDesk遠端  ID/password: 177586082/Cocoon-SIT-014"

}


Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 

 (get-process "msedge" -ea SilentlyContinue).CloseMainWindow()

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
start-sleep -s 5
$checknext_c=($webpage -match "next").count
$n=0

set-Clipboard -value "next"
start-sleep -s 5

[System.Windows.Forms.SendKeys]::SendWait('^f')
start-sleep -s 5

[System.Windows.Forms.SendKeys]::SendWait('^v')
start-sleep -s 5

[System.Windows.Forms.SendKeys]::SendWait('~')
if($checknext_c -gt 1){
do{
start-sleep -s 1
[System.Windows.Forms.SendKeys]::SendWait('~')
$n++
}until($n -eq $checknext_c)
}

[System.Windows.Forms.SendKeys]::SendWait('{ESC}')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('{TAB}')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('{DOWN}{DOWN}')
start-sleep -s 20

start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait('^a')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^c')
start-sleep -s 5
$webpage=Get-Clipboard

$settings=import-csv -path "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\settings.csv"

$ver1=$webpage -notmatch "Binaries" -notmatch "documents" -notmatch "Documents" -notmatch "China Only"　-notmatch "Developer Tools & Utilities" | where {$_ -match "Published"}

$list_old=get-content  -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\lists.txt
$list_new= $ver1
$comp_list=((Compare-Object $list_old $list_new)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject
# $old_logs=get-content -path  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\logs.txt"
 if($old_logs.length -eq 0){ $old_logs="na"}

 Write-Output "Line 138 comp_logs: $comp_list"

if($comp_list -and $comp_list.length -gt 0 -and $comp_list.count -gt 0){

foreach($comp_list1 in $comp_list){

$topic=(($comp_list1 -split '\t')[0]).trim()


$f1="1"

if($topic -match "Skip Ahead" -or $topic -match "Beta Channel Preview Build" -or $topic -match "Canary Channel Preview" `
           -and !($topic -match  "ARM64") -and !($topic -match  "China Only") -and !($topic -match  " Business")){
$f1="2"
}

if($topic -match "msu" -or $topic -match "Updated Package"){
$f1="3"
}

if($topic -match "\d{5}" -and $topic -match "kits"){
$f1="4"
}


$f1


###################### check next page ############
start-sleep -s 5
Set-Clipboard -value $topic
$xx=0
$ver1|%{

  if($_.replace("|","/") -match $topic.replace("|","/")){
  $_
  $xx++
  }

}

if($xx -gt 1){
$yy=0

do{
Set-Clipboard -value $topic
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^f')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^v')

if($yy -ne 0){
$y2 =$yy

do{

start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('~')
$y2=$y2-1

}until ($y2 -eq 0)
}

start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('{ESC}')

start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('+{RIGHT}')
start-sleep -s 1
[System.Windows.Forms.SendKeys]::SendWait('+{RIGHT}')
start-sleep -s 1
[System.Windows.Forms.SendKeys]::SendWait('+{RIGHT}')
start-sleep -s 1
[System.Windows.Forms.SendKeys]::SendWait('+{RIGHT}')
start-sleep -s 1
[System.Windows.Forms.SendKeys]::SendWait('+{RIGHT}')
start-sleep -s 1
[System.Windows.Forms.SendKeys]::SendWait('+{RIGHT}')
start-sleep -s 1
[System.Windows.Forms.SendKeys]::SendWait('+{RIGHT}')

start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^c')

$checktitle=get-Clipboard
$yy++
} until ($checktitle -match "builds")
}

else{

Set-Clipboard -value $topic
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^f')
start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('^v')

}

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

foreach($anlyx in $anly){
$allc0++
if($anlyx -match "51015All"){
$allc=$allc0
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
else{$webpage20=$webpage2}

$pattern= "Description(.*?)Files in this Package"
$result_status  = [regex]::match($webpage20, "$pattern").Groups[1].Value
$des1=$result_status.replace(",","，")

$pattern= "MICROSOFT CONFIDENTIAL.(.*?)All rights reserved."
#$pattern

$msrights  ="MICROSOFT CONFIDENTIAL."+ [regex]::match($des1, "$pattern").Groups[1].Value+"All rights reserved."

$desc=($des1 -replace $msrights, "").trim()

###check file names and  definded folder####


foreach ($set in $settings){

$lev1=$set."Level1"

if($f1 -eq $lev1){

$lev2=$set."Level2"
$fod=$set."Folder"


if($webpage20 -match $lev2){

  $webpage22=$webpage20 -match $lev2

  foreach($webpage2 in $webpage22){

  start-sleep -s 5
[System.Windows.Forms.SendKeys]::SendWait('^a')


  if($lev2 -match "iso"){  $select_dl=($webpage2 -split "iso")[0]+"iso"}
    if($lev2 -match "msu"){ $select_dl=($webpage2  -split "msu")[0]+"msu"}
    
    $xd=0
    $d=0
    $select_dl


###check file ranking####

    foreach($web in $webpage20){
    $d++
        if($web -match "Files in this Package"){
     $xd1=$d
        }
   
    if($web -match "$select_dl"){
     $xd2=$d-$xd1-1
        }
    }

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
$cc++
}until($cc -eq $xd2)


####start download####

 remove-item -path "\\192.168.56.48\Windows_ISO\06.AutoDLZone\*.crdownload" -force


$currentf=(gci -path "\\192.168.56.48\Windows_ISO\06.AutoDLZone\*").name

if($currentf.count -eq 0){$currentf="na"}
start-sleep -s 2

echo "Enter...downloading..."
start-sleep -s 5

$dl_p="\\192.168.56.48\Windows_ISO\06.AutoDLZone\$fod\"
 #####################################  Downloading    #####################################
[System.Windows.Forms.SendKeys]::SendWait('~')

#check download complete

$now_t=get-date
 do{
 start-sleep -s 60
 $check_ongoings =(Get-ChildItem -Path "\\192.168.56.48\Windows_ISO\06.AutoDLZone\*.crdownload").count+(Get-ChildItem -Path "\\192.168.56.48\Windows_ISO\06.AutoDLZone\*.tmp").count
 #$check_ongoings
 
 $now_t2=[datetime](get-date)
 
  $now_gap=(NEW-TIMESPAN –Start $now_t –End  $now_t2).TotalMinutes

  }until($check_ongoings -eq 0 -or $now_gap -gt 100)
 
 if($now_gap -lt 100){


  start-sleep -s 5
 $currentf2=(gci -path "\\192.168.56.48\Windows_ISO\06.AutoDLZone\*").name
 $diff=((compare-object $currentf $currentf2)|Where-Object { $_.SideIndicator -eq "=>"}).InputObject

  Move-Item "\\192.168.56.48\Windows_ISO\06.AutoDLZone\$diff" -Destination "\\192.168.56.48\Windows_ISO\06.AutoDLZone\$fod\" -Force

  #####################################  Downloading    #####################################>
 
 $date_now=get-date -Format "M/d"
 
  #add-content -path  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\logs.txt" -value "$fod, $topic, $select_dl, $date_now"

  "{0},{1},{2},{3},{4}" -f "","","","","" | add-content -path  \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv -force
  
  $writeto= import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv  -Encoding  UTF8
     $writeto[-1]."type"= $fod
     $writeto[-1]."item1"= $topic.replace(",","，")
     $writeto[-1]."file"= $diff
     $writeto[-1]."date"= $date_now
     $writeto[-1]."desc"= $desc	

      $writeto| export-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv -Encoding  UTF8 -NoTypeInformation
      }



else{

 $paramHash = @{
   #To =  "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-ProjectLead@allion.com","wallacelee@allion.com"
   To="shuningyu17120@allion.com.tw"
 from = 'Win_Download <edata_admin@allion.com>'
 BodyAsHtml = $True
 Subject = "<Alarm:Windows Files> Please Check the download files might stopped (This is auto mail)"
 Body ="\\192.168.56.48\Windows_ISO\06.AutoDLZone <BR> \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download <BR>" 
}

Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 

exit

}



 }
}


}
}
###back to main page####

start-sleep -s 2
[System.Windows.Forms.SendKeys]::SendWait('%{LEFT}')

}

 $obj1=import-csv -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv -Encoding UTF8
 
 $comp_logs=((Compare-Object $obj0."item1" $obj1."item1")|Where-Object { $_.SideIndicator -eq "=>"}).InputObject|get-unique
 Write-Output "Line 469 comp_logs: $comp_logs"

 if($comp_logs.count -gt 0){
 $i=0
 $obj3=$null

  foreach( $comp_log in  $comp_logs){
   $countx=($obj1|?{ $_."item1" -ne "" -and $_."item1" -eq $comp_log }).count
    $obj3=$obj3+@("")*$countx
   $obj2=$obj1|?{ $_."item1" -ne "" -and $_."item1" -eq $comp_log }|select "type","item1","file","date"
      $obj3=  $obj3+ $obj2
   }
   $obj3=$obj3|?{$_."item1".length -ne 0}| ConvertTo-Html | Out-String
   
 #$new_logs=get-content -path   "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\logs.txt"
 #copy-item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv" -Destination "\\192.168.56.48\Windows_ISO\06.AutoDLZone\packages_log.csv" -Force
 remove-item -path  "\\192.168.56.48\Windows_ISO\06.AutoDLZone\packages_log.csv" -Force
 import-csv -path  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv"|?{$_.type -ne "office"}|export-csv "\\192.168.56.48\Windows_ISO\06.AutoDLZone\packages_log.csv" -Encoding UTF8 -NoTypeInformation


$end1="<BR>Check the Files: \\192.168.56.48\Windows_ISO\06.AutoDLZone"
$end2="<BR>Check the detail description from Query System: https://bu2-query.allion.com 【10_WinImg】<BR>"
$end4="<BR>Check Win10 Public Update News from the web: <BR>https://support.microsoft.com/en-us/topic/windows-10-update-history-1b6aac92-bf01-42b5-b158-f80c6d93eb11
<BR>Check Win11 Public Update News from the web: <BR>https://support.microsoft.com/en-us/topic/windows-11-update-history-a19cd327-b57f-44b9-84e0-26ced7109ba9"
$end3="<BR><BR>The latest Windows Insider Preview Builds Flight Hub:<BR> https://learn.microsoft.com/en-us/windows-insider/flight-hub/"
$end5="<BR><BR> Downlaod Cumulative updates from the web: <BR>https://www.catalog.update.microsoft.com/Home.aspx<BR>"
#$body="Please Check logs:<BR>"+'<font color="blue">'+$obj3+'</font>'+$end1+$end2+$end3+$end4+$end5
$body="Please Check logs:<BR>"+'<font color="blue">'+$obj3+'</font>'+$end1+$end2+$end3+$end5
 $body= $body -replace  '<table>', '<table border="1">'

  $madd=get-content "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\ref\maillists.txt"
  $maillis= $maillis+@($madd)

 $paramHash = @{
  #To =  "NPL-APP@allion.com","NPL-DRV@allion.com","NPL-ProjectLead@allion.com","wallacelee@allion.com","taira.hayashi@allion.co.jp","taira.hayashi@allion.co.jp","nplj_proper@allion.co.jp","nplj_partner@allion.co.jp"
   To= $maillis
   #To="shuningyu17120@allion.com.tw"
 from = 'Win_Download <edata_admin@allion.com>'
 BodyAsHtml = $True
 Subject = "<New Windows Files> Please Check the download files logs (This is auto mail)"
 Body =$body
}
Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 
}

add-content -path \\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\lists.txt -value $comp_list

}
 (get-process "msedge" -ea SilentlyContinue).CloseMainWindow()

########## remove logs of files not existing ##########

import-csv "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv"|?{test-path ("\\192.168.56.48\Windows_ISO\06.AutoDLZone\"+$_.Type+"\"+$_.file)} |export-csv "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packagesx.csv" -Encoding UTF8  -NoTypeInformation
Remove-Item "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv"
Move-Item -Path  "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packagesx.csv" -Destination "\\192.168.20.20\sto\EO\2_AutoTool\ALL\81.WindowsImage_Download\packages.csv" -force

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
   #region check　task normal
 
  $taskcheck_mscol="\\192.168.20.20\sto\EO\2_AutoTool\ALL\65.Query_database\database_generator\ftp\mscol_checktask.txt"
  $lastwriteday=get-date((gci $taskcheck_mscol).LastWriteTime).Date
  $hournow=(get-date).Hour
  $daynow=(get-date).Date
 
  if($hournow -ge 1 -and $daynow -ne $lastwriteday){
   $getmonth=get-date -Format "yyyy/M/d HH:mm"
   Set-Content -path  $taskcheck_mscol -Value "checktask:$getmonth"
  }
 
 
  #endregion