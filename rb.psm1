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
=============================================================================
#>

Function AggiornaProgressbarOBSOLETA {
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

Function Write-LogAdvancedOBSOLETA {
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

Function Get-FileEncoding($Path) {
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

Function Out-FileUtf8NoBom {
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
