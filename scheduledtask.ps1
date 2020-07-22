function WakeToRun {
    $PowerSchemeGUID = ([Regex]::new('[A-Za-z0-9]+\-[A-Za-z0-9-]+').Match($((powercfg.exe /LIST) | Select-String "power scheme guid" -List))).Value
    $AllowWakeTimersGUID = ([Regex]::new('[A-Za-z0-9]+\-[A-Za-z0-9-]+').Match($((powercfg.exe /q) | Select-String "(Allow wake timers)"))).Value

    $PowerSchemeGUID | Foreach-object {
    (('/SETDCVALUEINDEX {0} SUB_SLEEP {1} 1' -f $_, $AllowWakeTimersGUID), ('/SETACVALUEINDEX {0} SUB_SLEEP {1} 1' -f $_, $AllowWakeTimersGUID)) |
        Foreach-object {
            Start-Process powercfg.exe -ArgumentList $_ -Wait -Verb runas -WindowStyle Hidden
        }
    }
}


function SheduleTask {
    $Trigger = New-ScheduledTaskTrigger -At 8:30am -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday
    $Action = New-ScheduledTaskAction -Execute "powershell" -Argument "-File `"C:\Users\tdsmaug\EasyMorning\wakeup.ps1`" -Verb RunAs"
    $Settings = New-ScheduledTaskSettingsSet -WakeToRun -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Hidden
    $Principal = New-ScheduledTaskPrincipal -UserId "$($env:USERDOMAIN)\$($env:USERNAME)" -LogonType S4U -RunLevel Highest

    Register-ScheduledTask `
        -TaskName "WakeUp" `
        -Trigger $Trigger `
        -Action $Action `
        -Settings $Settings `
        -Principal $Principal `
        -Force
}

WakeToRun
SheduleTask