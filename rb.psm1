<#
=============================================================================
Libreria di funzioni per script PowerShell.
-----------------------------------------------------------------------------
INSTALLAZIONE: 
    Copiare questo file rbPS.psm1 in uno dei seguenti percorsi:
    %PROGRAMFILES%\WindowsPowerShell\Modules\rbPS\rbPS.psm1
        per rendere le funzioni disponibili a tutti gli Utenti del PC.
    %USERPROFILE%\Documents\WindowsPowerShell\Modules\rbPS\rbPS.psm1
        per rendere le funzioni disponibili all'Utente corrente del PC.
-----------------------------------------------------------------------------
AUTORE: Raffaele Bianco
AGGIORNATO IL: 11/9/2023
=============================================================================
#>

function AggiornaProgressbarOBSOLETA {
    <#
        .SYNOPSIS
            Aggiornamento Progressbar con ETA (Estimated Time of Arrival).
        .DESCRIPTION
            Progressbar avanzata, che visualizza: 
                - conteggio items processati
                - conteggio tempo trascorso
                - stima tempo rimanente
                - stima orario di termine elaborazione
        .EXAMPLE
            If ($ItemsProcessed % 100 -eq 0) { AggiornaProgressbar -PassNumber $ItemsProcessed -TotalNumber $ItemCount }
        .NOTES
            Chiamare questa funzione "non troppo spesso", perché rallenta l'esecuzione.
    #>

    param(
        [int]$PassNumber,
        [int]$TotalNumber,
        [string]$Description = "Esecuzione $ScriptName"
    )
    $now = Get-Date

    # Tempo trascorso dall'avvio dello script:
    $timeSpan = $now - $script:StartTime

    # Secondi per elaborare ciascun elemento:
    $timePer = $timeSpan.TotalSeconds / $PassNumber 
    $ItemsPerSecond = $PassNumber / ($timeSpan.TotalSeconds + 0.1)

    $remainingSeconds = ($TotalNumber - $PassNumber) * $timePer
    $remainingSpan = New-TimeSpan -seconds $remainingSeconds

    # Data e ora di completamento stimata:
    $eta = $now.AddSeconds($remainingSeconds)
    $FormattedETA = $eta # TODO: Formattare per output della data in formato ITA, anche su Windows ENU

    $FormattedTimeSpan = "$($timeSpan.Hours.ToString("00")):$($timeSpan.Minutes.ToString("00")):$($timeSpan.Seconds.ToString("00"))"
    
    $StatusString =                 "[" + [Math]::Round(($PassNumber / $TotalNumber * 100), 0) + "%]  "
    $StatusString = $StatusString + "[${FormattedTimeSpan}]  "
    $StatusString = $StatusString + "[-$($remainingSpan.Hours.ToString("00")):$($remainingSpan.Minutes.ToString("00")):$($remainingSpan.Seconds.ToString("00"))]  "
    $StatusString = $StatusString + "[ETA: $FormattedETA]  "
    $StatusString = $StatusString + "[Item $PassNumber/$TotalNumber]  "
    $StatusString = $StatusString + "[" + [Math]::Round(($ItemsPerSecond), 0) + " item/s]"

    Write-Progress -Activity $Description -Status $StatusString -PercentComplete ($ItemsProcessed++ / $ItemCount * 100)
    $Host.UI.RawUI.WindowTitle = "[" + [Math]::Round(($PassNumber / $TotalNumber * 100), 0) + "%] ${ScriptName}"
}

function Write-LogAdvancedOBSOLETA {
    <#
        .SYNOPSIS
            Scrive un testo in "formato log" nella console e/o in un LogFile.
        .EXAMPLE
            Write-Log -Message "Informazione."
            Write-Log -Message "Errore da registrare nel log definito globalmente a inizio script." -Level "ERROR"
            Write-Log -Message "Warning da registrare nel log specificato come parametro." -Level "WARN" -LogFile "C:\Logfile.log"
        .NOTES
            Aprire il file LOG con VSCode per ottenere lo syntax-highlighting corretto.
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]
        $Message,

        [Parameter(Mandatory=$False)]
        [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
        [String]
        $Level = "INFO",

        [Parameter(Mandatory=$False)]
        [String]
        $LogFile = $Script:LogFile,

        [Parameter(Mandatory=$False)]
        [ValidateSet("Log", "LogConsole", "LogProgressbar", "LogConsoleProgressbar")]
        [String]
        $Mode = "Log"
    )

    $TimeStamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss,fff")
    $Line = "$TimeStamp [$Level] $Message"
    
    If ($script:LogFileMutex.WaitOne(1000)) {
        Add-Content $LogFile -Value $Line
        [void]$script:LogFileMutex.ReleaseMutex()
    }
    Else {
        Write-Host "Impossibile scrivere nel file di log ""$LogFile""!" -BackgroundColor "Red" -ForegroundColor "White"
    }

    # Azioni dedicate alla visualizzazione in realtime nella console:
    If ($Mode -in "LogConsole", "LogConsoleProgressbar") {
        If ($Level -in "ERROR", "FATAL") { Write-Host $Line -BackgroundColor "Red" -ForegroundColor "White"}
        ElseIf ($Level -eq "WARN") { Write-Host $Line -ForegroundColor "Yellow"}
        Else { Write-Host $Line -ForegroundColor "Gray"}
    }

    # Azioni dedicate all'aggiornamento della progressbar:
    If ($Mode -in "LogConsoleProgressbar", "LogProgressbar") {
        # TODO: Incorporare la funzione AggiornaProgressbar con i relativi parametri (aggiungere i parametri in ingresso).
    }
}

function Get-FileEncoding($Path) {
    $bytes = [byte[]](Get-Content $Path -Encoding byte -ReadCount 4 -TotalCount 4)

    if(!$bytes) { return 'utf8' }

    switch -regex ('{0:x2}{1:x2}{2:x2}{3:x2}' -f $bytes[0],$bytes[1],$bytes[2],$bytes[3]) {
        '^efbbbf'   { return 'utf8' }
        '^2b2f76'   { return 'utf7' }
        '^fffe'     { return 'unicode' }
        '^feff'     { return 'bigendianunicode' }
        '^0000feff' { return 'utf32' }
        default     { return 'ascii' }
    }
}

function Out-FileUtf8NoBom {
	<#
    .SYNOPSIS
      Outputs to a UTF-8-encoded file *without a BOM* (byte-order mark).
    .DESCRIPTION
      Mimics the most important aspects of Out-File:
        * Input objects are sent to Out-String first.
        * -Append allows you to append to an existing file, -NoClobber prevents
          overwriting of an existing file.
        * -Width allows you to specify the line width for the text representations
           of input objects that aren't strings.
      However, it is not a complete implementation of all Out-File parameters:
        * Only a literal output path is supported, and only as a parameter.
        * -Force is not supported.
        * Conversely, an extra -UseLF switch is supported for using LF-only newlines.
      Caveat: *All* pipeline input is buffered before writing output starts,
              but the string representations are generated and written to the target
              file one by one.
    .NOTES
      The raison d'être for this advanced function is that Windows PowerShell
      lacks the ability to write UTF-8 files without a BOM: using -Encoding UTF8 
      invariably prepends a BOM.
      Copyright (c) 2017, 2020 Michael Klement <mklement0@gmail.com> (http://same2u.net), 
      released under the [MIT license](https://spdx.org/licenses/MIT#licenseText).
    #>
    
	[CmdletBinding()]
	param(
		[Parameter(Mandatory, Position = 0)] [string] $LiteralPath,
		[switch] $Append,
		[switch] $NoClobber,
		[AllowNull()] [int] $Width,
		[switch] $UseLF,
		[Parameter(ValueFromPipeline)] $InputObject
	)
    
	#requires -version 3
    
	# Convert the input path to a full one, since .NET's working dir. usually
	# differs from PowerShell's.
	$dir = Split-Path -LiteralPath $LiteralPath
	if ($dir) { $dir = Convert-Path -ErrorAction Stop -LiteralPath $dir } else { $dir = $pwd.ProviderPath }
	$LiteralPath = [IO.Path]::Combine($dir, [IO.Path]::GetFileName($LiteralPath))
    
	# If -NoClobber was specified, throw an exception if the target file already
	# exists.
	if ($NoClobber -and (Test-Path $LiteralPath)) {
		Throw [IO.IOException] "The file '$LiteralPath' already exists."
	}
    
	# Create a StreamWriter object.
	# Note that we take advantage of the fact that the StreamWriter class by default:
	# - uses UTF-8 encoding
	# - without a BOM.
	$sw = New-Object System.IO.StreamWriter $LiteralPath, $Append
    
	$htOutStringArgs = @{}
	if ($Width) {
		$htOutStringArgs += @{ Width = $Width }
	}
    
	# Note: By not using begin / process / end blocks, we're effectively running
	#       in the end block, which means that all pipeline input has already
	#       been collected in automatic variable $Input.
	#       We must use this approach, because using | Out-String individually
	#       in each iteration of a process block would format each input object
	#       with an indvidual header.
	try {
		$Input | Out-String -Stream @htOutStringArgs | ForEach-Object { 
			if ($UseLf) {
				$sw.Write($_ + "`n") 
			}
			else {
				$sw.WriteLine($_) 
			}
		}
	}
    finally {
		$sw.Dispose()
	}    
}

function AggiornaProgressbar {
    <#
    .SYNOPSIS
        Aggiornamento Progressbar con ETA (Estimated Time of Arrival).
    .DESCRIPTION
        Progressbar avanzata, che visualizza: 
            - conteggio items processati
            - conteggio tempo trascorso
            - stima tempo rimanente
            - stima orario di termine elaborazione
    .EXAMPLE
        if ($itemsProcessed % 100 -eq 0) {AggiornaProgressbar -PassNumber $itemsProcessed -TotalNumber $itemCount -Description $testo}
    .NOTES
        Chiamare questa funzione "non troppo spesso" (ad es. ogni 100 $itemsProcessed), per evitare di rallentare troppo lo script.
    #>

    param(
        [Parameter(Mandatory = $false)]
        [int] $PassNumber = 0
        ,
        [Parameter(Mandatory = $false)]
        [int] $TotalNumber = 100
        ,
        [Parameter(Mandatory = $false)]
        [string] $Description = "Attendere..."
        ,
        [Parameter(Mandatory = $false)]
        [int] $Id = 0
    )
    $scriptName = Split-Path $PSCommandPath -Leaf
    $now = Get-Date

    # Tempo trascorso dall'avvio dello script:
    $timePassed = $now - $Script:StartTimewatch

    # Secondi per elaborare ciascun elemento:
    $timePer = $timePassed.TotalSeconds / $PassNumber 
    $itemsPerSecond = ([Math]::Round(($PassNumber / ($timePassed.TotalSeconds + 0.1)), 1))
    if ($itemsPerSecond -lt 60) {
        $speed = ($itemsPerSecond * 60).ToString() + "/min"
    } 
    else {
        $speed = $itemsPerSecond.ToString() + "/sec"
    }


    $percentuale = ([Math]::Round(($PassNumber / $TotalNumber * 100), 0)).ToString() + "%"
    $statusString = $percentuale + "  "
    $statusString = $statusString + "$PassNumber/$TotalNumber  "
    $statusString = $statusString + "$($timePassed.Hours.ToString("00")):$($timePassed.Minutes.ToString("00")):$($timePassed.Seconds.ToString("00"))  "
	if (($PassNumber -ge 3) -and ($TotalNumber -ge 10)) {
        # Data e ora di completamento stimata:
        $remainingSeconds = ($TotalNumber - $PassNumber) * $timePer
        $remainingSpan = New-TimeSpan -seconds $remainingSeconds
        $eta = $now.AddSeconds($remainingSeconds)
        if ($now.ToString("d") -eq $eta.ToString("d")) {
            $FormattedETA = $eta.ToString("H:mm:ss")
        }
        else {
            $FormattedETA = $eta.ToString("d/M/yyyy H:mm:ss")
        }
        # $statusString = $statusString + "[-$($remainingSpan.Hours.ToString("00")):$($remainingSpan.Minutes.ToString("00")):$($remainingSpan.Seconds.ToString("00"))]  "
        $statusString = $statusString + "ETA=$FormattedETA  "
        $statusString = $statusString + "$speed"
        $secondiRimanenti = [Math]::Round($remainingSpan.TotalSeconds / 5.0) * 5
        # Write-Progress -Activity $Description -Status $statusString -PercentComplete ($PassNumber / $TotalNumber * 100) -Id $Id -SecondsRemaining $($remainingSpan.TotalSeconds.ToString())
        Write-Progress -Activity $Description -Status $statusString -PercentComplete ($PassNumber / $TotalNumber * 100) -Id $Id -SecondsRemaining $secondiRimanenti
    } else {
        Write-Progress -Activity $Description -Status $statusString -PercentComplete ($PassNumber / $TotalNumber * 100) -Id $Id
    }
    
    $Host.UI.RawUI.WindowTitle = "[" + $percentuale + "] ${ScriptName}"
}

function Write-LogAdvanced {
    <#
      .SYNOPSIS
          Scrive un testo in "formato log" nella console e/o in un LogFile.
      .EXAMPLE
          Write-Log -Message "Informazione."
          Write-Log -Message "Errore da registrare nel log definito globalmente a inizio script." -Level "ERROR"
          Write-Log -Message "Warning da registrare nel log specificato come parametro." -Level "WARN" -LogFile "C:\Logfile.log"
      .NOTES
          Aprire il file LOG con VSCode per ottenere lo syntax-highlighting corretto.
          Per Notepad++: https://darekkay.com/blog/turn-notepad-into-a-log-file-analyzer/
  #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True)]
        [String] $Message
        ,
        [Parameter(Mandatory = $false)]
        [ValidateSet("FATAL", "ERROR", "WARN", "INFO", "DEBUG", "OK")]
        [String] $Level = "INFO"
        ,
        [Parameter(Mandatory = $false)]
        [String] $LogFile = "$(Split-Path $PSCommandPath -Parent)\Log\$(Split-Path $PSCommandPath -Leaf)-$($Script:startTimestamp).log"
        ,
        [Parameter(Mandatory = $false)]
        [ValidateSet("Console", "Log", "Progressbar", "ConsoleLog", "ConsoleProgressbar", "LogProgressbar", "ConsoleLogProgressbar")]
        [String] $Mode = "Console"
        ,
        [Parameter(Mandatory = $false)]
        [Int] $CurrentItem
        ,
        [Parameter(Mandatory = $false)]
        [Int] $TotalItems
        ,
        [Parameter(Mandatory = $false)]
        [int] $ProgressbarId = 0
    )



    $TimeStamp = (Get-Date).ToString("HH:mm:ss,fff")
    $Line = "[$TimeStamp $Level] $Message"



    # CONSOLE:
    if ($Mode.Contains("Console")) {

        if ($Level -in "ERROR", "FATAL") { 
            Write-Host "[$TimeStamp $Level]" -BackgroundColor "red" -ForegroundColor "white" -NoNewline
            if ($Mode.Contains("Progressbar")) { Write-Host " (${CurrentItem}/${TotalItems})" -NoNewLine }
            Write-Host " $Message"
        }
        elseif ($Level -eq "WARN") { 
            Write-Host "[$TimeStamp $Level]" -BackgroundColor "yellow" -ForegroundColor "black" -NoNewline
            if ($Mode.Contains("Progressbar")) { Write-Host " (${CurrentItem}/${TotalItems})" -NoNewLine }
            Write-Host " $Message"
        }
        elseif ($Level -eq "DEBUG") { 
            Write-Host "[$TimeStamp $Level]" -BackgroundColor "cyan" -ForegroundColor "black" -NoNewline
            if ($Mode.Contains("Progressbar")) { Write-Host " (${CurrentItem}/${TotalItems})" -NoNewLine }
            Write-Host " $Message"
        }
        elseif ($Level -eq "INFO") {
            Write-Host "[$TimeStamp $Level]" -BackgroundColor "gray" -ForegroundColor "black" -NoNewline
            if ($Mode.Contains("Progressbar")) { Write-Host " (${CurrentItem}/${TotalItems})" -NoNewLine }
            Write-Host " $Message"
        }
        elseif ($Level -eq "OK") {
            Write-Host "[$TimeStamp  $Level ]" -BackgroundColor "green" -ForegroundColor "black" -NoNewline
            if ($Mode.Contains("Progressbar")) { Write-Host " (${CurrentItem}/${TotalItems})" -NoNewLine }
            Write-Host " $Message"
        }
    }



    # LOG:
    if ($Mode.Contains("Log")) {
        $logFolder = $(Split-Path $LogFile -Parent)
		if (!(Test-Path $logFolder)) { mkdir "$(Split-Path $PSCommandPath -Parent)\Log" | Out-Null }
		Try {
			Add-Content $LogFile -Value $Line
        }
        Catch {
            Write-Host "Impossibile scrivere nel file di log ""$LogFile""!" -BackgroundColor "Red" -ForegroundColor "White"
        }
    }



    # PROGRESSBAR:
    if ($Mode.Contains("Progressbar")) {
        $CurrentItem = [int]$CurrentItem
        $TotalItems = [int]$TotalItems
        AggiornaProgressbar -PassNumber $CurrentItem -TotalNumber $TotalItems -Description $Message -Id $ProgressbarId
    }
}

function Test-ValidFileName {
    param([string]$FileName)

    $indexOfInvalidChar = $FileName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars())

    # IndexOfAny() returns the value -1 to indicate no such character was found
    return $indexOfInvalidChar -eq -1
}

function Test-FolderAccess {
    param([string]$Folder)

    # Verifico se la $Folder esiste:
    if (!(Test-Path $Folder)) { return $false }

    # Verifico se posso leggere dalla $Folder un file qualunque (prendo il più piccolo per rapidità di esecuzione):
    try {
        # TODO: if ( cartella contiene zero file ) { #     crea file vuoto nella cartella         # }
        $myFile = Get-ChildItem "$Folder\*" -File | Sort-Object -Property Length | Select-Object -First 1
        $copiedFile = "$env:Temp\$(Split-Path $myFile -Leaf)"
        Copy-Item "$myFile" -Destination $copiedFile -Force # | Out-Null
        if (Test-Path $copiedFile) {
            Remove-Item -Path "$copiedFile"
        } else {
            # Non sono riuscito a scrivere il file nella %TEMP%
            return $false
        }
    }
    catch {
        return $false
    }
	return $true
}

function Wait-AllJobs {
    $runningJobs = Get-Job | Where-Object { $_.State -eq "Running" }
    $completedJobs = Get-Job | Where-Object { $_.State -eq "Completed" }

    if ($($runningJobs.Count) -le 0) { 
        $script:fase += $completedJobs.Count # Serve per aggiornare correttamente la progress-bar e il log, nel caso i job siano già completati prima di entrare in questa funzione.
    }

    $initialJobs = $($runningJobs.Count)
    $oldCjCount = 0
    $i = 0
    while($($runningJobs.Count) -gt 0) {
        $completedJobs = Get-Job | Where-Object { $_.State -eq "Completed" }
        $runningJobs = Get-Job | Where-Object { $_.State -eq "Running" }

        $percent = [math]::Round((($initialJobs - $($runningJobs.Count)) / $initialJobs * 100), 0) + ($i++)
        if ($percent -gt 99) { $percent = 99 }
        Write-Progress -Activity "Attesa completamento $($runningJobs.Count) Job in esecuzione ($($runningJobs.Name))..." -Status "Completati: $($initialJobs-$($runningJobs.Count))/$initialJobs" -PercentComplete $percent -Id 1 -ParentId 0

        if ($oldCjCount -lt $($completedJobs.Count)) { 
            $script:fase += ($($completedJobs.Count) - $oldCjCount)
            Write-LogAdvanced "Job completati: $($completedJobs.Name)" -Level "INFO" -Mode "ConsoleLogProgressbar" -CurrentItem ($script:fase) -TotalItems ($script:totaleFasi)
            $oldCjCount = $($completedJobs.Count)
            $i = 0
        }

        Start-Sleep -Seconds 1
    }
    Get-Job | Remove-Job
    Write-Progress "Job completati." -Id 1 -Completed
}
