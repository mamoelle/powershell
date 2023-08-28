function CopyAA {
    param($autoAttendant)

    ##$autoAttendant = get-csautoattendant
    ##$autoAttendant = Import-Clixml -Path "C:\tmp\backup\AABackup_20230827\AAConfig.xml"


    ## Copy Operator
    $operatorObjectId = $autoAttendant.Operator.Id
    $AutoattendantOperatorType = $autoAttendant.Operator.Type.Value

    switch ($AutoattendantOperatorType)
    {
        User
        {
            $operatorEntity = New-CsAutoAttendantCallableEntity -Identity $operatorObjectId -Type User 
        }
        ExternalPstn
        {
            $operatorEntity = New-CsAutoAttendantCallableEntity -Identity $operatorObjectId -Type ExternalPstn
        }
        ApplicationEndpoint
        {
            $operatorEntity = New-CsAutoAttendantCallableEntity -Identity $operatorObjectId -Type ApplicationEndpoint
        }
    }

    ## Copy inclusionScope and exclusionScope

    if($autoAttendant.DirectoryLookupScope.InclusionScope.GroupScope.GroupIds -ne $null)
    {
        $inclusionScope = New-CsAutoAttendantDialScope -GroupScope -GroupIds $autoAttendant.DirectoryLookupScope.InclusionScope.GroupScope.GroupIds
    }

    if($autoAttendant.DirectoryLookupScope.ExclusionScope.GroupScope.GroupIds -ne $null)
    {
        $exclusionScope = New-CsAutoAttendantDialScope -GroupScope -GroupIds $autoAttendant.DirectoryLookupScope.ExclusionScope.GroupScope.GroupIds
    }


    ## Copy Callflow Timeschedule

    $CallHandlingAssociationArray = @()
    $date = Get-Date -Format "yyyyMMdd"

    foreach ($CallHandlingAssociation in $autoAttendant.CallHandlingAssociations)
    {
        
        $CallHandlingAssociationType = $CallHandlingAssociation.Type.Value
        switch ($CallHandlingAssociationType)
        {
            Holiday
            {
                
                $filteredSchedule = $autoAttendant.Schedules | Where-Object { $_.Id -eq $CallHandlingAssociation.ScheduleId }
                $timeschedule = New-CsOnlineSchedule -Name ($autoAttendant.Name +  $filteredSchedule.Name + $date) -FixedSchedule -DateTimeRanges $filteredSchedule.FixedSchedule.DateTimeRanges
                $CallHandlingAssociationArray += New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $timeschedule.Id -CallFlowId $CallHandlingAssociation.CallFlowId
            }
            AfterHours
            {
            
                ## Get the right schedule 
                $filteredSchedule = $autoAttendant.Schedules | Where-Object { $_.Id -eq $CallHandlingAssociation.ScheduleId }

                ##$filteredSchedule = $autoAttendant.Schedules | Where-Object { $_.Id -eq "664c6bce-e503-49ff-8d59-7466f5e4f9a7" }

                   
                if ($filteredSchedule.WeeklyRecurrentSchedule.MondayHours.Start) {
                    $MondayHours = New-CsOnlineTimeRange -Start $filteredSchedule.WeeklyRecurrentSchedule.MondayHours.Start -end $filteredSchedule.WeeklyRecurrentSchedule.MondayHours.End
                }

                if ($filteredSchedule.WeeklyRecurrentSchedule.TuesdayHours.Start) {
                    $TuesdayHours = New-CsOnlineTimeRange -Start $filteredSchedule.WeeklyRecurrentSchedule.TuesdayHours.Start -end $filteredSchedule.WeeklyRecurrentSchedule.TuesdayHours.End
                }

                if ($filteredSchedule.WeeklyRecurrentSchedule.WednesdayHours.Start) {
                    $WednesdayHours = New-CsOnlineTimeRange -Start $filteredSchedule.WeeklyRecurrentSchedule.WednesdayHours.Start -end $filteredSchedule.WeeklyRecurrentSchedule.WednesdayHours.End
                }

                if ($filteredSchedule.WeeklyRecurrentSchedule.ThursdayHours.Start) {
                    $ThursdayHours = New-CsOnlineTimeRange -Start $filteredSchedule.WeeklyRecurrentSchedule.ThursdayHours.Start -end $filteredSchedule.WeeklyRecurrentSchedule.ThursdayHours.End
                }

                if ($filteredSchedule.WeeklyRecurrentSchedule.FridayHours.Start) {
                    $FridayHours = New-CsOnlineTimeRange -Start $filteredSchedule.WeeklyRecurrentSchedule.FridayHours.Start -end $filteredSchedule.WeeklyRecurrentSchedule.FridayHours.End
                }

                if ($filteredSchedule.WeeklyRecurrentSchedule.SaturdayHours.Start) {
                    $SaturdayHours = New-CsOnlineTimeRange -Start $filteredSchedule.WeeklyRecurrentSchedule.SaturdayHours.Start -end $filteredSchedule.WeeklyRecurrentSchedule.SaturdayHours.End
                }

                if ($filteredSchedule.WeeklyRecurrentSchedule.SundayHours.Start) {
                    $SundayHours = New-CsOnlineTimeRange -Start $filteredSchedule.WeeklyRecurrentSchedule.SundayHours.Start -end $filteredSchedule.WeeklyRecurrentSchedule.SundayHours.End
                }

                $filteredScheduleType = $filteredSchedule.Type.Value
                if ($filteredScheduleType -eq "WeeklyRecurrence") {
                    $timeschedule = New-CsOnlineSchedule -Name ($autoAttendant.Name +  "_AfterHours_" + $date) -WeeklyRecurrentSchedule -MondayHours $MondayHours -TuesdayHours $TuesdayHours -WednesdayHours $WednesdayHours -ThursdayHours $ThursdayHours -FridayHours $FridayHours -SaturdayHours $SaturdayHours -SundayHours $SundayHours
                }
                               
                $CallHandlingAssociationArray += New-CsAutoAttendantCallHandlingAssociation -Type AfterHours -ScheduleId $timeschedule.Id -CallFlowId $CallHandlingAssociation.CallFlowId

            }
        }

    }


    # Define the parameters and values to give to New-CsAutoAttendant.
    $parameters = @{
        Name =$autoAttendant.Name
        LanguageId = $autoAttendant.LanguageId 
        TimeZoneId = $autoAttendant.TimeZoneId 
        VoiceId = $autoAttendant.VoiceId
        DefaultCallFlow  = $autoAttendant.DefaultCallFlow
    }

    # Add optional parameters to hashtable
    if ($CallHandlingAssociationArray) {$parameters.Add("CallHandlingAssociations", $CallHandlingAssociationArray)}
    if ($autoAttendant.CallFlows) {$parameters.Add("CallFlows", $autoAttendant.CallFlows)}
    if ($inclusionScope) {$parameters.Add("InclusionScope", $InclusionScope)}
    if ($ExclusionScope) {$parameters.Add("ExclusionScope", $ExclusionScope)}
    if ($autoAttendant.AuthorizedUsers) {$parameters.Add("AuthorizedUsers", $autoAttendant.AuthorizedUsers)}
    if ($operatorEntity) {$parameters.Add("Operator", $operatorEntity)}
    if ($autoAttendant.VoiceResponseEnabled) {$parameters.Add("EnableVoiceResponse",$true)}

    New-CsAutoAttendant @parameters

}    

function ListAAPrompts {
    param($AutoAttendants)

    $AudioPromptArray = @()
    
    foreach ($autoAttendant in $autoAttendants)
    {
        $AAName = $autoAttendant.Name
        $callflows = $autoAttendant.CallFlows
        Foreach ($callflow in $callflows) 
        {

            $greeting = $callflow.Greetings.AudioFilePrompt
            $menu = $callflow.Menu.Prompts.AudioFilePrompt

            if( $greeting -ne $null)
            {
                
                $greeting |   Add-Member -NotePropertyName "AAName" -NotePropertyValue $autoAttendant.Name -Force
                $greeting |   Add-Member -NotePropertyName "AAIdentity" -NotePropertyValue $autoAttendant.Identity -Force
                $greeting |   Add-Member -NotePropertyName "AAAudioPromptType" -NotePropertyValue "CFG" -Force
                $greeting |   Add-Member -NotePropertyName "AACallFlowName" -NotePropertyValue $callflow.Name  -Force
                $greeting |   Add-Member -NotePropertyName "CallFlowMenuToneDTMFResponse" -NotePropertyValue ""  -Force

                $AudioPromptArray += $greeting                    
            }

            if($menu -ne $null)
            {
                $menu |   Add-Member -NotePropertyName "AAName" -NotePropertyValue $autoAttendant.Name -Force
                $menu |   Add-Member -NotePropertyName "AAIdentity" -NotePropertyValue $autoAttendant.Identity -Force
                $menu |   Add-Member -NotePropertyName "AAAudioPromptType" -NotePropertyValue "CFM" -Force
                $menu |   Add-Member -NotePropertyName "AACallFlowName" -NotePropertyValue $callflow.Name  -Force
                $menu |   Add-Member -NotePropertyName "CallFlowMenuToneDTMFResponse" -NotePropertyValue "" -Force


                $AudioPromptArray += $menu

            }
            
            foreach ($menuoption in $callflow.Menu.MenuOptions) 
            {
                $DtmfResponse = $menuoption.DtmfResponse
                $prompt = $menuoption.Prompt.AudioFilePrompt
                if($prompt -ne $null)
                {
                    $prompt |   Add-Member -NotePropertyName "AAName" -NotePropertyValue $autoAttendant.Name -Force
                    $prompt |   Add-Member -NotePropertyName "AAIdentity" -NotePropertyValue $autoAttendant.Identity -Force
                    $prompt |   Add-Member -NotePropertyName "AAAudioPromptType" -NotePropertyValue "CFMT" -Force
                    $prompt |   Add-Member -NotePropertyName "AACallFlowName" -NotePropertyValue $callflow.Name -Force
                    $prompt |   Add-Member -NotePropertyName "CallFlowMenuToneDTMFResponse" -NotePropertyValue $DtmfResponse  -Force


                    $AudioPromptArray += $prompt

                }
            }
        }

        $callflows = $autoAttendant.DefaultCallFlow
        Foreach ($callflow in $callflows) 
        {

            $greeting = $callflow.Greetings.AudioFilePrompt
            $menu = $callflow.Menu.Prompts.AudioFilePrompt
            
            if($greeting -ne $null)
            {
                $greeting |   Add-Member -NotePropertyName "AAName" -NotePropertyValue $autoAttendant.Name -Force
                $greeting |   Add-Member -NotePropertyName "AAIdentity" -NotePropertyValue $autoAttendant.Identity -Force
                $greeting |   Add-Member -NotePropertyName "AAAudioPromptType" -NotePropertyValue "DCFG" -Force
                $greeting |   Add-Member -NotePropertyName "AACallFlowName" -NotePropertyValue $callflow.Name  -Force
                $greeting |   Add-Member -NotePropertyName "CallFlowMenuToneDTMFResponse" -NotePropertyValue "" -Force


                $AudioPromptArray += $greeting   
            }
        
            if($menu -ne $null)
            {
                $menu |   Add-Member -NotePropertyName "AAName" -NotePropertyValue $autoAttendant.Name -Force
                $menu |   Add-Member -NotePropertyName "AAIdentity" -NotePropertyValue $autoAttendant.Identity -Force
                $menu |   Add-Member -NotePropertyName "AAAudioPromptType" -NotePropertyValue "DCFM" -Force
                $menu |   Add-Member -NotePropertyName "AACallFlowName" -NotePropertyValue $callflow.Name  -Force
                $menu |   Add-Member -NotePropertyName "CallFlowMenuToneDTMFResponse" -NotePropertyValue "" -Force


                $AudioPromptArray += $menu
            }

            foreach ($menuoption in $callflow.Menu.MenuOptions) 
            {
                $DtmfResponse = $menuoption.DtmfResponse
                $prompt = $menuoption.Prompt.AudioFilePrompt

                if($prompt -ne $null)
                {
                    $prompt |   Add-Member -NotePropertyName "AAName" -NotePropertyValue $autoAttendant.Name -Force
                    $prompt |   Add-Member -NotePropertyName "AAIdentity" -NotePropertyValue $autoAttendant.Identity -Force
                    $prompt |   Add-Member -NotePropertyName "AAAudioPromptType" -NotePropertyValue "DCFMT" -Force
                    $prompt |   Add-Member -NotePropertyName "AACallFlowName" -NotePropertyValue $callflow.Name -Force
                    $prompt |   Add-Member -NotePropertyName "CallFlowMenuToneDTMFResponse" -NotePropertyValue $DtmfResponse  -Force

                    $AudioPromptArray += $prompt
                }
            }
        }
    }

    return $AudioPromptArray

}

function UpdateAutoAttendantPromptConfig 
{
    param($AAPromptCSVList)
    
    $i = 0
    Foreach ($AAPrompt in $AAPromptCSVList)
    {
            switch ($AAPrompt.AAAudioPromptType)
            {
                CFG
                {
                    if(($autoAttendants.CallFlows.Greetings.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName -ne $null)
                    {
                        Write-Host "Updating Call Flow Greeting File " ($autoAttendants.CallFlows.Greetings.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName  "to" $AAPrompt.FileNameNew
                        Write-Host "Updating Call Flow Greeting ID" ($autoAttendants.CallFlows.Greetings.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).Id  "to" $AAPrompt.IdNew
                        ($autoAttendants.CallFlows.Greetings.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName  = $AAPrompt.FileNameNew
                        ($autoAttendants.CallFlows.Greetings.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).Id  = $AAPrompt.IdNew
                        $i = $i + 1
                    }
                }
                CFM
                {
                    if(($autoAttendants.CallFlows.Menu.Prompts.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName -ne $null)
                    {
                        Write-Host "Updating Call Flow Menu Prompt File " ($autoAttendants.CallFlows.Menu.Prompts.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName  "to" $AAPrompt.FileNameNew
                        Write-Host "Updating Call Flow Menu Prompt ID" ($autoAttendants.CallFlows.Menu.Prompts.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).Id  "to" $AAPrompt.IdNew
                        ($autoAttendants.CallFlows.Menu.Prompts.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName  = $AAPrompt.FileNameNew
                        ($autoAttendants.CallFlows.Menu.Prompts.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).Id  = $AAPrompt.IdNew
                        $i = $i + 1
                    }
                }
                CFMT
                {
                    if(($autoAttendants.CallFlows.Menu.MenuOptions.Prompt.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName -ne $null)
                    {
                        Write-Host "Updating Call Flow Menu MenuOptions File " ($autoAttendants.CallFlows.Menu.MenuOptions.Prompt.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName  "to" $AAPrompt.FileNameNew
                        Write-Host "Updating Call Flow Menu MenuOptions ID" ($autoAttendants.CallFlows.Menu.MenuOptions.Prompt.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).Id  "to" $AAPrompt.IdNew
                        ($autoAttendants.CallFlows.Menu.MenuOptions.Prompt.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName  = $AAPrompt.FileNameNew
                        ($autoAttendants.CallFlows.Menu.MenuOptions.Prompt.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).Id  = $AAPrompt.IdNew
                        $i = $i + 1
                    }
                }
                DCFG
                {
                    if(($autoAttendants.DefaultCallFlow.Greetings.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName -ne $null)
                    {
                        Write-Host "Updating Default Call Flow Greeting File " ($autoAttendants.DefaultCallFlow.Greetings.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName  "to" $AAPrompt.FileNameNew
                        Write-Host "Updating Default Call Flow Greeting ID" ($autoAttendants.DefaultCallFlow.Greetings.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).Id  "to" $AAPrompt.IdNew
                        ($autoAttendants.DefaultCallFlow.Greetings.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName  = $AAPrompt.FileNameNew
                        ($autoAttendants.DefaultCallFlow.Greetings.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).Id  = $AAPrompt.IdNew
                        $i = $i + 1
                    }
                }
                DCFM
                {
                    if(($autoAttendants.DefaultCallFlow.Menu.Prompts.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName -ne $null)
                    {
                        Write-Host "Updating Default Call Flow Menu Prompt File " ($autoAttendants.DefaultCallFlow.Menu.Prompts.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName  "to" $AAPrompt.FileNameNew
                        Write-Host "Updating Default Call Flow Menu Prompt ID" ($autoAttendants.DefaultCallFlow.Menu.Prompts.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).Id  "to" $AAPrompt.IdNew
                        ($autoAttendants.DefaultCallFlow.Menu.Prompts.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName  = $AAPrompt.FileNameNew
                        ($autoAttendants.DefaultCallFlow.Menu.Prompts.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).Id  = $AAPrompt.IdNew
                        $i = $i + 1
                    }
                }
                DCFMT
                {
                    if(($autoAttendants.DefaultCallFlow.Menu.MenuOptions.Prompt.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName -ne $null)
                    {
                        Write-Host "Updating Default Call Flow Menu MenuOptions File " ($autoAttendants.DefaultCallFlow.Menu.MenuOptions.Prompt.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName  "to" $AAPrompt.FileNameNew
                        Write-Host "Updating Default Call Flow Menu MenuOptions ID" ($autoAttendants.DefaultCallFlow.Menu.MenuOptions.Prompt.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).Id  "to" $AAPrompt.IdNew
                        ($autoAttendants.DefaultCallFlow.Menu.MenuOptions.Prompt.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).FileName  = $AAPrompt.FileNameNew
                        ($autoAttendants.DefaultCallFlow.Menu.MenuOptions.Prompt.AudioFilePrompt |  Where-Object { $_.Id -eq $AAPrompt.Id }).Id  = $AAPrompt.IdNew
                        $i = $i + 1
                    }

                }

            }
    }

    write-host $i "prompts updated" -ForegroundColor Yellow
}


function CreateNewAnnouncementsFromFile
{
    param($Path)
    
    
    $AAPromptCSVList = Import-csv -Path ($Path + "AAPromptList.csv")
    
    $AudioPromptArray = @()



    foreach($AAPrompt in $AAPromptCSVList)
    {

        Write-host "Processing prompt ID" $AAPrompt.Id
        Write-host "Processing promt filename" $AAPrompt.FileName
        Write-host "Processing URL" $AAPrompt.DownloadUri


        $fileBytes = [System.IO.File]::ReadAllBytes(($Path + $AAPrompt.FileNameNew))

        $date = Get-Date -Format "yyyyMMdd"


        if($AAPrompt.CallFlowMenuToneDTMFResponse -ne "")
        {

            $fileName = $AAPrompt.AAName + "_" + $AAPrompt.AACallFlowName + "_" + $AAPrompt.AAAudioPromptType + "_" + $AAPrompt.CallFlowMenuToneDTMFResponse + "_" + $date + ".wav"

            if($fileName.Length -gt 64)
            {
                $AAPromptAACallFlowName = $AAPrompt.AACallFlowName.Substring(0,$AAPrompt.AACallFlowName.Length-($fileName.Length-64))
                Write-Host "Shortening filename due to filename character length > 64:" -ForegroundColor Yellow
                $fileName = $AAPrompt.AAName + "_" + $AAPromptAACallFlowName + "_" + $AAPrompt.AAAudioPromptType + "_" + $AAPrompt.CallFlowMenuToneDTMFResponse + "_" + $date + ".wav"
            }
        }
        else
        {

            $fileName = $AAPrompt.AAName + "_" + $AAPrompt.AACallFlowName + "_" + $AAPrompt.AAAudioPromptType + "_" + $date + ".wav"

            if($fileName.Length -gt 64)
            {
                $AAPromptAACallFlowName = $AAPrompt.AACallFlowName.Substring(0,$AAPrompt.AACallFlowName.Length-($fileName.Length-64))
                Write-Host "Shortening filename due to filename character length > 64:" -ForegroundColor Yellow
                $fileName = $AAPrompt.AAName + "_" + $AAPromptAACallFlowName + "_" + $AAPrompt.AAAudioPromptType + "_" + $date + ".wav"
            }
        }

        

        $UploadCsOnlineAudioFileResult = UploadCsOnlineAudioFile -content $fileBytes -fileName $fileName

        $AAPrompt.FileNameNew =  $UploadCsOnlineAudioFileResult.FileName 
        $AAPrompt.IdNew = $UploadCsOnlineAudioFileResult.Id 


        $AudioPromptArray += $AAPrompt
    }

    
    Write-Host $AudioPromptArray.count "audio prompts processed."
    Write-Host "Exporting CSV" ($path + "AAPromptListNew.csv")
    $AudioPromptArray | Export-Csv -Path ($path + "AAPromptListNew.csv") -NoTypeInformation 
    return $AudioPromptArray
}


function UploadCsOnlineAudioFile
{
    param($content, $fileName)

       
       Write-host "Uploading new file" $fileName "with" $fileName.Length "character filename length"

       $audioFile = Import-CsOnlineAudioFile -ApplicationId "OrgAutoAttendant" -FileName $fileName -Content $content
       Write-Host "Uploaded"  $fileName  "successfully"

       return $audioFile
}


function BackupAAPrompts
{
    param($ListAAPrompts, $Path)
    
    $AudioPromptArray = @()

    $webClient = New-Object System.Net.WebClient

    foreach($AAPrompt in $ListAAPrompts)
    {

        Write-host "Download prompt ID" $AAPrompt.Id
        Write-host "Download promt filename" $AAPrompt.FileName
        Write-host "Download URL" $AAPrompt.DownloadUri

        $AAPrompt | Add-Member -NotePropertyName "FileNameNew" -NotePropertyValue $UploadCsOnlineAudioFileResult.FileName -Force
        $AAPrompt | Add-Member -NotePropertyName "IdNew" -NotePropertyValue $UploadCsOnlineAudioFileResult.Id -Force

        $fileUrl = $AAPrompt.DownloadUri
        $fileBytes = $webClient.DownloadData($fileUrl)

        if($fileBytes.Length -gt 0)
        {
            Write-Host "Downloaded" $fileBytes.Length "Bytes successfully"
        }
        else
        {
            Write-Host "Download failed" -ForegroundColor Red
        }

        $date = Get-Date -Format "yyyyMMdd"


        if($AAPrompt.CallFlowMenuToneDTMFResponse -ne "")
        {

            $fileName = $AAPrompt.AAName + "_" + $AAPrompt.AACallFlowName + "_" + $AAPrompt.Id + "_" + $AAPrompt.AAAudioPromptType + "_" + $AAPrompt.CallFlowMenuToneDTMFResponse + "_" + $date + ".wav"
        }
        else
        {

            $fileName = $AAPrompt.AAName + "_" + $AAPrompt.AACallFlowName + "_" + $AAPrompt.Id + "_" + $AAPrompt.AAAudioPromptType + "_" + $date + ".wav"
        }

        Set-Content -Path ($Path + "\" + $fileName) -Value $fileBytes -Encoding Byte


        $AAPrompt.FileNameNew = $fileName 
        $AAPrompt.IdNew = "" 

        $AudioPromptArray += $AAPrompt
    }
    
    Write-Host $AudioPromptArray.count "audio prompts processed." -ForegroundColor Yellow
    Write-Host "Export AA Prompt List to " $path"AAPromptList.csv"

    $AudioPromptArray | Export-Csv -Path ($path + "\AAPromptList.csv") -NoTypeInformation 
    return $AudioPromptArray
}


function BackupAAConfig
{
    param($Path)
    
    ##$Path ="C:\tmp\backup\"

    Write-host "Connect to Teams Service"
    ConnectTeams

    If(Test-Path -Path ($Path + "AABackup_" + (Get-Date -Format "yyyyMMdd")))
    {
        $BackupFolder = $Path + "AABackup_" + (Get-Date -Format "yyyyMMdd")
    }
    else
    {
        $BackupFolder = New-Item -ItemType Directory -Path ($Path + "AABackup_" + (Get-Date -Format "yyyyMMdd")) -Force
        $BackupFolder = $BackupFolder.FullName
    }
    

    Write-host "Exporting Auto Attendant config to" $BackupFolder"\AAConfig.xml"
    $xml = $BackupFolder + "\AAConfig.xml"
    Get-CsAutoAttendant | Export-Clixml -Path ($xml) -Depth 1000
    Write-host (Get-CsAutoAttendant).count  "Auto Attendants exported" -ForegroundColor Yellow

    Write-host "Check for Auto Attendant Audio Prompts"
    $AAPrompts = ListAAPrompts -AutoAttendants (Import-Clixml -Path $xml)
    

    if(($AAPrompts.Id).Count  -gt 0)
    {
        Write-host ($AAPrompts.Id).Count "audio prompts found" -ForegroundColor Yellow
        BackupAAPrompts -ListAAPrompts $AAPrompts -Path $BackupFolder
    }
    else
    {
        Write-host "No Auto Attendant Audio Prompts found" -ForegroundColor Yellow
    }

}

function ConnectTeams
{
    try
    {
    
        Write-Host "Testing Teams Powershell Connection"
        $testconnection = Get-CsTenant
        if($testconnection)
        {
            Write-Host "Powershell connected" -ForegroundColor Yellow
        }
    }
    catch 
    {
        If($_.Exception.Message)
        {
            Write-Host "Teams Powershell not connected"
            Connect-MicrosoftTeams
        }
    }
}



function RestoreAAConfig
{
    param($Path)


    ##$Path = "C:\tmp\backup\AABackup_20230827\"

    Write-host "Connect to Teams Service"
    ConnectTeams

    Write-host "Import Auto Attendant Config "
    $autoAttendants = Import-Clixml -Path ($path + "AAConfig.xml")
    
    if((ListAAPrompts -AutoAttendants $autoAttendants).ID.count -gt 0)
    {
        CreateNewAnnouncementsFromFile -Path $Path
        UpdateAutoAttendantPromptConfig -AAPromptCSVList (Import-Csv -Path ($path + "AAPromptListNew.csv"))
    }   


    foreach ($autoAttendant in $autoAttendants)
    {
        Write-host "Copy " $autoAttendant.Name
        CopyAA($autoAttendant)
    }

}


function deleteAllAA
{
    foreach ($AA in Get-CsAutoAttendant)
    {
        Remove-Csautoattendant $AA.Id
    }

    foreach ($schedule in Get-CsOnlineSchedule)
    {
        Remove-CsOnlineSchedule $schedule.Id
    }
}



#### Define working folder
$path="C:\tmp\backup\"

##BackupAAConfig -Path "C:\tmp\backup\"

##RestoreAAConfig -Path "C:\tmp\backup\AABackup_20230827\"
