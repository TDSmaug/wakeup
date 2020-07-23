param(
    [Parameter(Mandatory=$false)]
    [switch]$rm,
    [Parameter(Mandatory=$false)]
    [switch]$alarm
)


function MyVolume {

    Add-Type -TypeDefinition @'

using System.Runtime.InteropServices;
[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume {
    // f(), g(), ... are unused COM method slots. Define these if you care
    int f(); int g(); int h(); int i();
    int SetMasterVolumeLevelScalar(float fLevel, System.Guid pguidEventContext);
    int j();
    int GetMasterVolumeLevelScalar(out float pfLevel);
    int k(); int l(); int m(); int n();
    int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, System.Guid pguidEventContext);
    int GetMute(out bool pbMute);
}
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice {
    int Activate(ref System.Guid id, int clsCtx, int activationParams, out IAudioEndpointVolume aev);
}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator {
    int f(); // Unused
    int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }
public class Audio {
    static IAudioEndpointVolume Vol() {
    var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
    IMMDevice dev = null;
    Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(/*eRender*/ 0, /*eMultimedia*/ 1, out dev));
    IAudioEndpointVolume epv = null;
    var epvid = typeof(IAudioEndpointVolume).GUID;
    Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
    return epv;
    }
    public static float Volume {
    get {float v = -1; Marshal.ThrowExceptionForHR(Vol().GetMasterVolumeLevelScalar(out v)); return v;}
    set {Marshal.ThrowExceptionForHR(Vol().SetMasterVolumeLevelScalar(value, System.Guid.Empty));}
    }
    public static bool Mute {
    get { bool mute; Marshal.ThrowExceptionForHR(Vol().GetMute(out mute)); return mute; }
    set { Marshal.ThrowExceptionForHR(Vol().SetMute(value, System.Guid.Empty)); }
    }
}

'@

}


function Alarm {

    $key      = Get-Content ('{0}\key' -f $PSScriptRoot)
    $url      = 'https://www.googleapis.com/youtube/v3'
    $playlist = 'PLq5DDV1fyL0Rc26gkELyg16cX4-z50IE7'

    $totalResults = (

        Invoke-RestMethod `
            -Uri ('{0}/playlistItems?key={1}&part=snippet&maxResults=50&playlistId={2}' -f $url, $key, $playlist) `
            -Method Get `
            -UseBasicParsing

    ).pageInfo.totalResults

    $resultsPerPage = (

        Invoke-RestMethod `
            -Uri ('{0}/playlistItems?key={1}&part=snippet&maxResults=50&playlistId={2}' -f $url, $key, $playlist) `
            -Method Get `
            -UseBasicParsing

    ).pageInfo.resultsPerPage

    $step = [math]::truncate($totalResults/$resultsPerPage)

    $nextPageToken = (Invoke-RestMethod `
                        -Uri ('{0}/playlistItems?key={1}&part=snippet&maxResults=50&playlistId={2}' -f $url, $key, $playlist) `
                        -Method Get `
                        -UseBasicParsing).nextPageToken

    $tokens += @($null)

    for ($i = 0; $i -lt $step; $i++ ) {

        $tokens += $nextPageToken

        $nextPageToken = (

        Invoke-RestMethod `
            -Uri ('{0}/playlistItems?key={1}&part=snippet&maxResults=50&playlistId={2}&pageToken={3}' -f $url, $key, $playlist, $nextPageToken) `
            -Method Get `
            -UseBasicParsing

        ).nextPageToken

    }

    $songs = $tokens | ForEach-Object {
        (
            Invoke-RestMethod `
                -Uri ('{0}/playlistItems?key={1}&part=snippet&maxResults=50&playlistId={2}&pageToken={3}' -f $url, $key, $playlist, $_) `
                -Method Get `
                -UseBasicParsing

        ).items.snippet
    }

    $number = 0

    while ($true) {

        switch ($number) {

            { 0..1 -contains $_ } { [audio]::Volume = 0.25 }

            { 2..3 -contains $_ } { [audio]::Volume = 0.50 }

            { 4..5 -contains $_ } { [audio]::Volume = 0.75 }

            default { [audio]::Volume = 1.0 }

        }
        $randomSong = (Get-Random $songs)

        $duration = (

            Invoke-RestMethod `
                -Uri ('{0}/videos?id={1}&part=contentDetails&key={2}' -f $url, $(($randomSong).resourceId.videoId), $key) `
                -Method Get `
                -UseBasicParsing

        ).items.contentDetails.duration

        $durationTime  = [Regex]::new('[0-9]+[A-Za-z]+[0-9]+').Matches($duration)
        $durationMin   = [Regex]::new('\b[0-9]+').Matches($durationTime).Value
        $durationSec   = [Regex]::new('\B[0-9]+').Matches($durationTime).Value
        $durationTotal = $([int]$durationMin * 60) + [int]$durationSec

        Start-Process chrome.exe -ArgumentList ('--new-window https://music.youtube.com/watch?v={0}&list=PLq5DDV1fyL0Rc26gkELyg16cX4-z50IE7' -f $(($randomSong).resourceId.videoId))

        Write-Output ('"{0}" is playing now. Enjoy ^_^' -f ($randomSong).title)

        Start-Sleep -s $($durationTotal + 5)

        Get-Process | ForEach-Object {

            if ($_.name -like '*chrome*') {

                Stop-Process $_.id -ErrorAction SilentlyContinue

            }

        }

        $number++

    }

}


function WakeToRun {

    $PowerSchemeGUID     = ([Regex]::new('[A-Za-z0-9]+\-[A-Za-z0-9-]+').Match($((powercfg.exe /LIST) | Select-String "power scheme guid" -List))).Value
    $AllowWakeTimersGUID = ([Regex]::new('[A-Za-z0-9]+\-[A-Za-z0-9-]+').Match($((powercfg.exe /q) | Select-String "(Allow wake timers)"))).Value

    $PowerSchemeGUID | Foreach-object {

    (('/SETDCVALUEINDEX {0} SUB_SLEEP {1} 1' -f $_, $AllowWakeTimersGUID), ('/SETACVALUEINDEX {0} SUB_SLEEP {1} 1' -f $_, $AllowWakeTimersGUID)) |
        Foreach-object {

            Start-Process powercfg.exe -ArgumentList $_ -Wait -Verb runas -WindowStyle Hidden

        }

    }

}


function SheduleTask {

    $taskName = 'WakeUp'
    $Trigger = New-ScheduledTaskTrigger -At 8:15am -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday
    $Action = New-ScheduledTaskAction -Execute "powershell" -Argument "$PSCommandPath -alarm"
    $Settings = New-ScheduledTaskSettingsSet -WakeToRun -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Hidden
    $Principal = New-ScheduledTaskPrincipal -UserId "$($env:USERDOMAIN)\$($env:USERNAME)" -LogonType S4U -RunLevel Highest

    switch ($rm) {

        $true {

            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false

        }

        $false {

            Register-ScheduledTask `
                -TaskName $taskName `
                -Trigger $Trigger `
                -Action $Action `
                -Settings $Settings `
                -Principal $Principal `
                -Force

        }

    }

}

if ($alarm -eq $true) {

    MyVolume -ErrorAction SilentlyContinue
    Alarm

}

else {

    WakeToRun
    SheduleTask

}