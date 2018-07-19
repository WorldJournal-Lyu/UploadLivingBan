<#

UploadLivingBan.ps1
    
    2017-12-14 Initial Creation
    2017-12-15 Add send newspage proof to ad
    2017-12-18 Only send notification to ad if file is not found in newspage
    2017-12-19 Don't send anything to ad
    2018-01-03 Add 45218, 45221 to xml
    2018-01-05 Rewrite algorithm: use ok file to determine which indd, inx, pdf, ok files to process
    2018-05-22 Implement Modules, code optimization

#>


if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

Import-Module WorldJournal.Log -Verbose -Force
Import-Module WorldJournal.Email -Verbose -Force
Import-Module WorldJournal.Server -Verbose -Force

$scriptPath = $MyInvocation.MyCommand.Path
$scriptName = (($MyInvocation.MyCommand) -Replace ".ps1")
$configXml  = $scriptName+".xml"
$xmlPath    = ((Split-Path ($scriptPath)) + "\" + $configXml)

$zuBanRoot  = (Get-WJPath -Name zuban).Path
$newspage   = (Get-WJPath -Name newspage).Path
$graphic    = (Get-WJPath -Name graphic).Path

$mailMsg    = ""
$newspageFileMsg = ""
$graphicFileMsg  = ""
$errMsg     = ""
$crErMsg    = ""
[System.collections.ArrayList]$expectedNewspage = @()
[System.collections.ArrayList]$expectedGraphic  = @()
[System.collections.ArrayList]$okExtensions     = @(".indd", ".inx", ".pdf", ".ok")



# create log

$newlog = New-Log -Path $scriptPath -LogFormat yyyyMMdd-HHmmss
$log    = $newlog.FullName
Write-Log -Verb "LOG START" -Noun $log -Path $log -Type Long -Status Normal



# create date

if ((Get-Date).Hour -le 6){

    $toDay =  (Get-Date).AddDays(-1).ToString("yyyyMMdd")
    $weekDay = (Get-Date).AddDays(-1).DayOfWeek.value__
    if($weekDay -eq 0){ $weekDay = 7 }
    $workDate = (Get-Date).ToString("yyyyMMdd")

} else {

    $toDay = (Get-Date).ToString("yyyyMMdd")
    $weekDay = (Get-Date).DayOfWeek.value__
    if($weekDay -eq 0){ $weekDay = 7 }
    $workDate = (Get-Date).AddDays(1).ToString("yyyyMMdd")

}

Write-Log -Verb "Today   " -Noun ($toDay + "("+ $weekDay +")") -Path $log -Type Short -Status Normal
Write-Log -Verb "Tomorrow" -Noun $workDate -Path $log -Type Short -Status Normal



# load config xml

if(Test-Path $xmlPath){

    [xml]$xml = Get-Content $xmlPath -Encoding UTF8
    Write-Log -Verb "LOAD CONFIG" -Noun $xml.config.name -Path $log -Type Long -Status Good

    foreach($item in $xml.config.ban){

        $expBan = $item.id

        $expNewspageDay = [int]$item.printday + [int]$item.newspage
        if($expNewspageDay -le 0){ $expNewspageDay = $expNewspageDay + 7 }
        $expGraphicDay = [int]$item.printday + [int]$item.graphic
        if($expGraphicDay -le 0){ $expGraphicDay = $expGraphicDay + 7 }

        Write-Log -Verb $expBan -Noun ("NEWSPAGE " + $expNewspageDay + " GRAPHIC " + $expGraphicDay) -Path $log -Type Short -Status Normal

        if($expNewspageDay -eq $weekDay){ $expectedNewspage += $expBan }
        if($expGraphicDay -eq $weekDay){ $expectedGraphic += $expBan }
    }

    Write-Log -Verb "EXPECTED NEWSPAGE LIST" -Noun $expectedNewspage -Path $log -Type Long -Status Normal
    Write-Log -Verb "EXPECTED GRAPHIC  LIST" -Noun $expectedGraphic -Path $log -Type Long -Status Normal

}else{

    Write-Log -Verb "LOAD CONFIG" -Noun $xml.config.name -Path $log -Type Long -Status Bad

}


Write-Line -Length 50 -Path $log

$banList = @(Get-ChildItem $zuBanRoot | Where-Object{$_.PSIsContainer -and $_.Name -match "^45\d{3}"}) | Where-Object {$xml.config.ban.id -contains ($_.Name).substring(0,5)} 

foreach($ban in $banList){

    Write-Log -Verb "BAN" -Noun $ban.Name -Path $log -Type Long -Status Normal

    $dateList = @(Get-ChildItem $ban.FullName | Where-Object{$_.PSIsContainer -and $_.Name -match "^20\d{6}"})

    foreach($date in $dateList){

        Write-Log -Verb "DATE" -Noun $date.Name -Path $log -Type Long -Status Normal

        # check if folder date is valid

        try{

            $folderDate = [datetime]($date.Name.Substring(0,4)+"/"+$date.Name.Substring(4,2)+"/"+$date.Name.Substring(6,2))
            $dateCheck = $true

            $folderWeekDay = $folderDate.DayOfWeek.value__
            if($folderWeekDay -eq 0){
                $folderWeekDay = 7
            }

        }catch{

            $dateCheck = $false

        }

        if($dateCheck){

            $banData = ($xml.config.ban | Where-Object {$_.id -eq $ban.Name.Substring(0,5)})

            if($banData){

                # data for ban is defined in xml

                $newspageAdj = $banData.newspage
                $graphicAdj  = $banData.graphic
                $printDay    = $banData.printday

                # check if folder date's weekday is correct
                
                if($printDay -eq $folderWeekDay -or $printDay -eq 0){

                    $weekDayCheck = $True

                }else{

                    Write-Log -Verb "DATE INCORRECT" -Noun ($ban.Name + "/" + $date.Name) -Path $log -Type Long -Status Bad

                    $weekDayCheck = $False
                    $errMsg += $ban.Name + "/" + $date.Name + " (Incorrect Date)`n"

                }

            }else{

                # data for ban is NOT defined in xml, use default value

                $newspageAdj = -1
                $graphicAdj  = -1

            }

            # don't process newspage if marked as "X"

            if($banData.newspage -eq "X" -or $weekDayCheck -eq $False){

                $newspageSendDate = "SKIP"

            }else{

                $newspageSendDate = ($folderDate).AddDays($newspageAdj).ToString("yyyyMMdd")

            }

            # don't process graphic if marked as "X"

            if($banData.graphic -eq "X" -or $weekDayCheck -eq $False){

                $graphicSendDate = "SKIP"

            }else{

                $graphicSendDate = ($folderDate).AddDays($graphicAdj).ToString("yyyyMMdd")

            }

            Write-Log -Verb "TO NEWSPAGE" -Noun $newspageSendDate -Path $log -Type Short -Status Normal
            Write-Log -Verb "TO GRAPHIC " -Noun $graphicSendDate -Path $log -Type Short -Status Normal



############# COPY TO NEWSPAGE

            if($newspageSendDate -eq $toDay){

                # newspage date equals current date

                Write-Log -Verb "TO NEWSPAGE" -Noun ($ban.Name + "/" + $date.Name) -Path $log -Type Long -Status Normal
                $expectedNewspage.Remove($ban.Name.Substring(0,5))

                if(Test-Path ($date.FullName+"\*.ok")){

                    if(!(Test-Path ($newspage+"\"+$date.Name.Substring(0,8)))){

                        New-Item ($newspage+"\"+$date.Name.Substring(0,8)) -Type Directory | Out-Null
                        Write-Log -Verb "CREATE NEWSPAGE" -Noun ($newspage+"\"+$date.Name.Substring(0,8)) -Path $log -Type Long -Status Normal

                    }

                    $okfiles = @(Get-ChildItem $date.FullName -filter *.ok)

                    foreach($okfile in $okfiles){

                        $n_banId = ($okfile.Name).SubString(0,(($okfile.Name).length-3))
                        $n_from  = (($okfile.FullName).SubString(0,(($okfile.FullName).length-3))+".pdf")
                        $n_to    = ($newspage+$date.Name.Substring(0,8)+"\"+$n_banId+".pdf")

                        # exception for 45218: only 45218.pdf needs to go to newspage, 45218-8 is for SF only
                        if($n_banId -ne "45218-8"){
                        
                            try{

                                Write-Log -Verb "COPY TO NEWSPAGE FROM" -Noun $n_from -Path $log -Type Long -Status Normal
                                Copy-Item $n_from $n_to -ErrorAction Stop
                                Write-Log -Verb "COPY TO NEWSPAGE" -Noun $n_to -Path $log -Type Long -Status Good

                            }catch{

                                Write-Log -Verb "COPY TO NEWSPAGE" -Noun $n_to -Path $log -Type Long -Status Bad
                                Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Normal
                                $crErMsg += "COPY " + $n_from + " TO " + $n_to + " ERROR"
                                $crErMsg += $_.Exception.Message + "`n"

                            }

                            if( Test-Path $n_to ){

                                Write-Log -Verb "NEWSPAGE FILE CHECK" -Noun $n_to -Path $log -Type Long -Status Good
                                $newspageFileMsg += $n_to + "`n"

                            }else{
                
                                Write-Log -Verb "NEWSPAGE FILE CHECK" -Noun $n_to -Path $log -Type Long -Status Bad
                                $crErMsg += "CHECK FAIL " + $n_to + "`n"

                            }

                        }else{

                            Write-Log -Verb "SKIP" -Noun $n_banId -Path $log -Type Long -Status Normal
                       
                        }

                    }

                    $okfile  = ""
                    $okfiles = @()

                }else{

                    Write-Log -Verb "TO NEWSPAGE STOPPED" -Noun "No ok file" -Path $log -Type Long -Status Bad
                    $crErMsg += $ban.Name + "/" + $date.Name + " (OK Not Found, copy to Newspage stopped)`n"

                }

            }

############# COPY TO GRAPHIC

            if($graphicSendDate -eq $toDay){

                # graphic date equals current date
            
                Write-Log -Verb "TO GRAPHIC" -Noun ($ban.Name + "/" + $date.Name) -Path $log -Type Long -Status Normal
                $expectedGraphic.Remove($ban.Name.Substring(0,5))

                if( Test-Path ($date.FullName+"\*.ok") ){

                    # copy images (and other files) to graphic

                    $files = Get-ChildItem ($zuBanRoot+"/"+$ban.Name+"/"+$date.Name) | Where-Object{ 
                    ($_.Extension -ne ".txt") -and 
                    ($_.Extension -ne ".indd") -and 
                    ($_.Extension -ne ".inx") -and 
                    ($_.Extension -ne ".ok") -and 
                    ($_.Extension -ne ".pdf") -and 
                    ($_.Extension -ne ".idlk") -and (!$_.PSIsContainer) -and ($_.Name -like ($ban.Name.Substring(0,5)+"*.*")) }

                    foreach($file in $files){

                        $g_from_img = $file.FullName
                        $g_to_img   = ($graphic+$file.Name)
                        
                        try{
                        
                            Write-Log -Verb "COPY TO GRAPHIC FROM" -Noun $g_from_img -Path $log -Type Long -Status Normal
                            Copy-Item $g_from_img $g_to_img -ErrorAction Stop
                            Write-Log -Verb "COPY TO GRAPHIC" -Noun $g_to_img -Path $log -Type Long -Status Good
                            $graphicFileMsg += $g_from_img + "`n"

                        }catch{

                            Write-Log -Verb "COPY TO GRAPHIC" -Noun $g_to_img -Path $log -Type Long -Status Bad
                            Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Normal
                            $crErMsg += "COPY " + $g_from_img + " TO " + $g_to_img + " ERROR`n"
                            $crErMsg += $_.Exception.Message + "`n"

                        }
                    
                    }

                    # copy indd, inx, ok, pdf to graphic

                    $okfiles = @(Get-ChildItem $date.FullName -filter *.ok)

                    foreach($okfile in $okfiles){

                        foreach($okExtension in $okExtensions){

                            $g_banId = ($okfile.Name).SubString(0,(($okfile.Name).length-3))
                            $g_from = (($okfile.FullName).SubString(0,(($okfile.FullName).length-3))+$okExtension)
                            $g_to   = ($graphic+$g_banId+$okExtension)

                            try{

                                Write-Log -Verb "COPY TO GRAPHIC FROM" -Noun $g_from -Path $log -Type Long -Status Normal
                                Copy-Item $g_from $g_to -ErrorAction Stop
                                Write-Log -Verb "COPY TO GRAPHIC" -Noun $g_to -Path $log -Type Long -Status Good
                                $graphicFileMsg += $g_from + "`n"

                            }catch{

                                Write-Log -Verb "COPY TO GRAPHIC" -Noun $g_to -Path $log -Type Long -Status Bad
                                Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Normal
                                $crErMsg += "COPY " + $g_from + " TO " + $g_to + " ERROR"
                                $crErMsg += $_.Exception.Message + "`n"

                            }


                        }

                        Write-Log -Verb "REMOVE OK" -Noun $okfile.Name -Path $log -Type Long -Status Normal
                        Remove-Item $okfile.FullName -Force

                    }


                }else{

                    Write-Log -Verb "TO GRAPHIC STOPPED" -Noun "No ok file" -Path $log -Type Long -Status Bad
                    $crErMsg += $ban.Name + "/" + $date.Name + " (OK Not Found, copy to Graphic stopped)`n"

                }

            }

        }else{
        
             Write-Log -Verb "DATE INVALID" -Noun ($ban.Name + "/" + $date.Name) -Path $log -Type Long -Status Bad
            $errMsg += $ban.Name + "/" + $date.Name + " (Invalid Date)`n"

        }

    }

    
    Write-Line -Length 50 -Path $log

}


Write-Log -Verb "EXPECTED NEWSPAGE LIST" -Noun $expectedNewspage -Path $log -Type Long -Status Normal
Write-Log -Verb "EXPECTED GRAPHIC  LIST" -Noun $expectedGraphic -Path $log -Type Long -Status Normal

$mailSbj = $scriptName
$mailMsg += "Upload script completed at: " + (Get-Date).ToString("yyyy/MM/dd HH:mm:ss") + "`n`n"

if(!(Test-Path $xmlPath)){
    $mailMsg += "Config XML not found:`n" + $xmlPath + "`n"
    $mailSbj = "ERROR " + $scriptName
}

if($newspageFileMsg -ne ""){
    $mailMsg += "To Newspage:`n" + $newspageFileMsg + "`n"
}

if($newspageFileMsg -ne ""){
    $mailMsg += "To Graphic:`n" + $graphicFileMsg + "`n"
}

if($crErMsg -ne ""){
    $mailMsg += "File Copy Error:`n" + $crErMsg + "`n"
    $mailSbj = "ERROR " + $scriptName
}

if($errMsg -ne ""){
    $mailMsg += "Warnings:`n" + $errMsg + "`n"
    $mailSbj = "WARNING " + $scriptName
}

if($expectedNewspage.Count -gt 0){
    $mailMsg += "Folder Not Found for Newspage:`n" + $expectedNewspage + "`n"
    $mailSbj = "ERROR " + $scriptName
}

if($expectedGraphic.Count -gt 0){
    $mailMsg += "Folder Not Found for Graphic:`n" + $expectedGraphic + "`n"
    $mailSbj = "ERROR " + $scriptName
}


Write-Log -Verb "LOG END" -Noun "" -Path $log -Type Long -Status Normal

$mailFrom = (Get-WJEmail -Name noreply).MailAddress
$mailPass = (Get-WJEmail -Name noreply).Password
$mailTo = (Get-WJEmail -Name lyu).MailAddress
Emailv2 -From $mailFrom -Pass $mailPass -To $mailTo -Subject $mailSbj -Body $mailMsg -ScriptPath $scriptPath -Attachments $log