###################################################
#
#
# Script zur Steuerung der EAM-Gewichtsberechnung
#
# (c) 2018 Oliver Jung für MSG Mediaservice GmbH
# 09.01.2018 (Di)
#
# V1.0: 12.02.2018 (Mo)
###################################################

# Ermitteln des Script-Pfades
function Get-ScriptDirectory {
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
    Split-Path $Invocation.MyCommand.Path
}
$Pfad = Get-ScriptDirectory
$env:programmpfad = $Pfad+"\"
echo ==============================================

# Variablendeklaration
$logdate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
$datum = Get-Date
$ProdDatum = (Get-Date).AddDays(+1)
$env:EAMDatum = $ProdDatum.ToString("yyMMdd")
#$env:EAMDatum = '171104'
#$env:programmpfad ="C:\EAMtest\"
$input =$env:programmpfad+"EAMInput\*"
$output=$env:programmpfad+"EAMOutput\"
$putput=$env:programmpfad+"EAMPutput\"
$archiv=$env:programmpfad+"EAMArchiv\"
$Logfile=$env:programmpfad+"EAM_Gewichtsberechnung_log_"+$datum.year+".txt"
$scripts=$env:programmpfad+"Scripts\"
$finish = $output+"Finish.ok"
$javaprogramm = $scripts+'EAM_Gewichtsberechnung.bat '+$env:programmpfad
$scpdownload =$scripts+"EAM_GET.bat " +$env:programmpfad +" "+$env:EAMDatum
$scpupload_success = $scripts+'EAM_put.bat '+$env:programmpfad
$scpupload_fallback = $scripts+'EAM_fallback.bat '+$env:programmpfad
$UmfangslistePfad='\\ppi-classicserver01.server\pcxchange\SeitenProETag'
$BeilagenlistePfad='\\svl-fil02\msg$\Msg allgemein\Vertrieb\EAM_Beilagengewichte.txt'
$new=$putput+"Umfangsliste.txt"
###########################################
$SendmailPath = "C:\Sendmail\Sendmail.bat"
$From="-from ""eam@msgmediaservice.de"""
$to="-to ""o.jung@msgmediaservice.de"""
$cc="-cc ' '"
$bcc="-bcc ' '" 
$Body="-body """""
$Subject="-subject """""
#########################################

#Umfangliste von PPI-Server abholen

$ListenDatum=$ProdDatum.ToString("yyyyMMdd")
foreach($file in Get-ChildItem $UmfangslistePfad)
{
    if ($file.Name -match $ListenDatum)
    {
        Copy-Item $file.FullName -Destination $putput
        $dest=$putput + $file.Name
        Copy-Item $dest $new	
    }
}


#Beilagenliste von Filserver holen

if ((Get-Item $BeilagenlistePfad).CreationTime.ToString("yyyyMMdd")  -match $ListenDatum )
{
	Copy-Item -Path $BeilagenlistePfad -Destination $putput
}


#Download der EAM-Daten von Vi&Va mit Logline nur im Fehlerfall
Invoke-Expression $scpdownload
$return = $LASTEXITCODE
if ($return -eq 0)
{
    # Prüfen, ob Dateien heruntergeladen wurden
    if (Test-Path $input)
    {
    # Logifle schreiben
    echo ($logdate+": PS Script gestartet")
    Add-Content $Logfile("##########################################################################")
    Add-Content $Logfile ($logdate+": PS Script gestartet - Java Gewichtsberechnung wird aufgerufen")
	
	#EAM-Dateien ins Archiv kopieren
	foreach ($files in Get-ChildItem $input)
	{
		Copy-Item $files.FullName -Destination $archiv
	}
    # Aufruf Programm zur Berechnung
    #Add-Content $Logfile( Invoke-Expression $javaprogramm)

    # Check, ob finale OK-Datei erzeugt wurde
    <#if(Test-Path $finish) 
    {
        $logdate = Get-Date -Format "dd.MM.yyyy HH:mm:s"
        Add-Content $Logfile ($logdate+": Finishdatei existiert - EAM-Daten werden kopiert - Starte WINSCP")
        Add-Content $Logfile ("==========================================================================")
        Remove-Item $finish

        #BAT-Ausfürhen, welches internes Script über WINSCP ausführt und die Daten an EAM-Server lädt
        Add-Content $Logfile( Invoke-Expression $scpupload_success)
        $return = $LASTEXITCODE
        Add-Content $Logfile ("==========================================================================")
        $logdate = Get-Date -Format "dd.MM.yyyy HH:mm:s"
        if ($return -eq 0)
        {
            Add-Content $Logfile ($logdate+": Postupload erfolgreich")
            
        }
        else
        {
            Add-Content $Logfile ($logdate+": Postupload Fehlerhaft - bitte prüfen") 
			$Subject="-subject ""EAM-Postupload Fehlerhaft - bitte prüfen"""
			$env:message = $From +" "+$to +" "+$cc +" "+$bcc +" "+$Subject +" "+$Body
			$SendMailMessage =$SendmailPath+" "+$env:message
			Invoke-Expression $SendMailMessage 
			
        }
        
		#Verzeichnisse leeren		
        Remove-Item ($output+"\*")
        Remove-Item ($input)
    } #>

    # Wenn keine finale OK-Datei erzeugt wurde, gab es ein Problem bei der Gewichtsberechnung und es werden 
    # die alternativen Vi&Va-Dateien hochgeladen
    #else
    #{
        $logdate = Get-Date -Format "dd.MM.yyyy HH:mm:s"
        #Add-Content $Logfile ($logdate+": Problem bei Gewichtsberechnung - EAM-Fallback aus Vi&Va wird kopiert - Starte WINSCP")
		#$Subject="-subject ""Problem bei Gewichtsberechnung - EAM-Fallback aus Vi&Va wird kopiert"""
		#$Body="-body ""Details siehe Fehlermail von EAM-Berechnungsprogramm"""
	    #$env:message = $From +" "+$to +" "+$cc +" "+$bcc +" "+$Subject +" "+$Body
	    #$SendMailMessage =$SendmailPath+" "+$env:message
	   # Invoke-Expression $SendMailMessage
        Add-Content $Logfile ("==========================================================================")
        
        #BAT-Ausfürhen, welches internes Script über WINSCP ausführt und die Daten an EAM-Server lädt
        Add-Content $Logfile( Invoke-Expression $scpupload_fallback)
        $return = $LASTEXITCODE
        Add-Content $Logfile ("==========================================================================")
        $logdate = Get-Date -Format "dd.MM.yyyy HH:mm:s"
        if ($return -eq 0)
        {
            Add-Content $Logfile ($logdate+": Fallback Postupload erfolgreich")
            
        }
        else
        {
            Add-Content $Logfile ($logdate+": Fallback Postupload Fehlerhaft - bitte prüfen!")
			$Subject="-subject ""Fallback Postupload Fehlerhaft - bitte prüfen"""
			$Body="-body """""
			$env:message = $From +" "+$to +" "+$cc +" "+$bcc +" "+$Subject +" "+$Body
			$SendMailMessage =$SendmailPath+" "+$env:message
			Invoke-Expression $SendMailMessage
        }
        Remove-Item ($output+"\*")
        Remove-Item ($input)
        
    #}

    Add-Content $Logfile ($logdate+": PS Script beendet")
    Add-Content $Logfile("##########################################################################")
    }
    else
    {
        echo "Keine EAM-Daten vorhanden"
    }
}


#Falls Download Fehler geworfen hat, wird Problem-Mail ausgegeben und weitere Verarbeitung gestoppt
else
{
    Add-Content $Logfile ($logdate+": !!! Problem mit WINSCP-Download von Vi&Va !!!")
	$Subject="-subject ""EAM-Gewichtsberechnung kann EAM-Daten nicht downloaden"""
	$Body="-body ""Problem mit WINSCP-Download von ViVa. Es wurden keine EAM-Daten an die Post gesendet. Bitte unbedingt Logdaten und EAM-Daten prüfen"""
    $env:message = $From +" "+$to +" "+$cc +" "+$bcc +" "+$Subject +" "+$Body
    $SendMailMessage =$SendmailPath+" "+$env:message
    Invoke-Expression $SendMailMessage
}


#Umfangliste löschen
if (Test-Path $new) {
 Remove-Item -Path $new 
}

#Beilagenliste löschen
$BeilagenPfad = $putput+'EAM_Beilagengewichte.txt'
if (Test-Path $BeilagenPfad) {
 Remove-Item -Path $BeilagenPfad
}

#Alte Archivdateien löschen
$ArchivDatum = (Get-Date).AddDays(-30) 
Get-ChildItem $archiv | Where-Object { $_.CreationTime -lt $ArchivDatum } | Remove-Item
<#foreach ($archivFiles in Get-ChildItem $archiv)
{
	if ($archivFiles.CreationTime -lt $ArchivDatum)
	{
		Remove-Item $archivFiles.FullName
	}
} #>

