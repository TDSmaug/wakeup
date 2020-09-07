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

function Write-Log {
    param(
        [string]$text,
        [string]$tofile
    )

    if (!(Get-ChildItem ('{0}\EasyMorning\logs' -f $ENV:USERPROFILE) -ErrorAction SilentlyContinue)) {

        New-Item -ItemType Directory -Force -Path ('{0}\EasyMorning\logs' -f $ENV:USERPROFILE) | Out-Null

    }

    Write-Output "$(Get-Date) $text" >> ('{0}\EasyMorning\logs\{1} - {2}.txt' `
        -f $ENV:USERPROFILE, (Get-Date).ToString("dd.MM.yyyy"), $tofile)
}


function Invoke-MyRestMethod {
    param(
        [Parameter(Mandatory=$true)]
        [string]$url,
        [Parameter(Mandatory=$true)]
        [string]$key,
        [Parameter(Mandatory=$true)]
        [string]$playlist,
        [Parameter(Mandatory=$false)]
        [string]$nextPageToken
    )

    Invoke-RestMethod `
        -Uri ('{0}/playlistItems?key={1}&part=snippet&maxResults=50&playlistId={2}&pageToken={3}' -f $url, $key, $playlist, $nextPageToken) `
        -Method Get `
        -UseBasicParsing
}


function Alarm {

    MyVolume -ErrorAction SilentlyContinue

    $key      = Get-Content ('{0}\key' -f $PSScriptRoot)
    $url      = 'https://www.googleapis.com/youtube/v3'
    $playlist = 'PLq5DDV1fyL0Rc26gkELyg16cX4-z50IE7'

    $totalResults = (Invoke-MyRestMethod $url $key $playlist).pageInfo.totalResults

    $resultsPerPage = (Invoke-MyRestMethod $url $key $playlist).pageInfo.resultsPerPage

    $step = [math]::truncate($totalResults/$resultsPerPage)

    $nextPageToken = (Invoke-MyRestMethod $url $key $playlist).nextPageToken

    $tokens += @($null)

    for ($i = 0; $i -lt $step; $i++ ) {

        $tokens += $nextPageToken

        $nextPageToken = (Invoke-MyRestMethod $url $key $playlist $nextPageToken).nextPageToken

    }

    $songs = $tokens | ForEach-Object {

        (Invoke-MyRestMethod $url $key $playlist $_).items.snippet

    }

    $channelId = (Invoke-MyRestMethod $url $key $playlist).items.snippet[0].channelId

    $myplaylists = (
        
    Invoke-RestMethod `
        -Uri ('{0}/playlists?key={1}&part=snippet&maxResults=50&channelId={2}' -f $url, $key, $channelId) `
        -Method Get `
        -UseBasicParsing
    ).items

    $playlistTitle = $myplaylists | ForEach-Object {
        if ($_.id -contains $playlist) {
            $_.snippet.title
        }
    }

    $AudioSaveLocation = "$ENV:USERPROFILE\Music\$playlistTitle"
    $CacheFolder = "$ENV:USERPROFILE\EasyMorning\cache"
    $number = 0

    $songs | ForEach-Object {

        if (($_).resourceId.videoId -notcontains '7X1L8_MDj4I' ) {

            if (Get-ChildItem ('{0}\{1}.mp3' -f $AudioSaveLocation, (($_).title -replace '"', '''' -replace '\?', '')) -ErrorAction SilentlyContinue) {
                Write-Log ('"{0}" is present' -f ($_).title) 'logs'
            }

            else {
                Write-Log ('"{0}"...' -f ($_).title) 'logs'

                $URLToDownload = ('https://music.youtube.com/watch?v={0}&list=PLq5DDV1fyL0Rc26gkELyg16cX4-z50IE7' -f $(($_).resourceId.videoId))

                youtube-dl -o "$AudioSaveLocation\%(title)s.%(ext)s" `
                    --ignore-errors --no-mtime --quiet --no-warnings --no-playlist `
                    --cache-dir "$CacheFolder" -x --audio-format mp3 --audio-quality 0 `
                    --metadata-from-title "(?P<artist>.+?) - (?P<title>.+)" `
                    --add-metadata --prefer-ffmpeg "$URLToDownload"

                if (Get-ChildItem ('{0}\{1}.mp3' -f $AudioSaveLocation, (($_).title -replace '"', '''' -replace '\?', '')) -ErrorAction SilentlyContinue) {
                    Write-Log ('"{0}" has been downloaded' -f ($_).title) 'logs'
                }
                else {
                    Write-Log ('"{0}" is unavailable' -f ($_).title) 'errorlogs'
                }
            }
        }
    }

    ((Get-ChildItem $AudioSaveLocation).FullName | Sort-Object {Get-Random}) | ForEach-Object {
        
        $number++

        switch ($number) {

            { 1..9 -contains $_ } { [audio]::Volume = ('0.{0}' -f $_ ) }

            default { [audio]::Volume = 1.0 }

        }
        
        Add-Type -AssemblyName presentationCore
        $mediaPlayer = New-Object system.windows.media.mediaplayer
        $mediaPlayer.open($_)
        Start-Sleep -s 2
        $duration = $mediaPlayer.NaturalDuration.TimeSpan.TotalSeconds
        $mediaPlayer.Play()
        Start-Sleep -s ([Math]::Truncate(($duration + 0.5)))
        $mediaPlayer.Stop()
        $mediaPlayer.Close()
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

    WakeToRun

    $taskName = 'WakeUp'
    $Trigger = New-ScheduledTaskTrigger -At 9:30am -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday
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

    Alarm

}

else {

    SheduleTask

}