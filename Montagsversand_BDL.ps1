###################################################
#
#
# Umbau der Bereitstellung und Steuerung des 
# Montagsversands im Zuge der Umstellung auf 
# Dialog CRM von alten Batch-Skripten 
# auf PowerShell
#
#
# (c) 17.01.2018 Oliver Jung
#
#
###################################################

# Ermitteln des Script-Pfades
function Get-ScriptDirectory {
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
    Split-Path $Invocation.MyCommand.Path
}
#$Pfad = Get-ScriptDirectory
$Pfad ="\\svl-fil02\msg$\Msg allgemein\Vertrieb"
$env:programmpfad = $Pfad+"\"

echo ==============================================


#Variablendeklaration
$Datum = Get-Date
$KWString = Get-Date -UFormat %V
$KW = "{0:D2}" -f [int]$KWString
$dir = "\\svl-fil01\homeb$\MARKET\Versand\"+$Datum.year+"\"+$Datum.year+"_Zugaenge\"+$Datum.year+"_"+$KW
$logdate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
$Logfile=$env:programmpfad+"Montagszugeange_log"+$datum.year+".txt"
$AF_Upload = "& '"+$env:programmpfad+"Arbeitsdaten\Montagszugaenge\Adressfactory_PUT.bat"+"'"
$AF_Download = "& '"+$env:programmpfad+"Arbeitsdaten\Montagszugaenge\Adressfactory_GET.bat"+"'"
$unzip ="& '"+$env:programmpfad+"Arbeitsdaten\Montagszugaenge\Unzip_Adressfactory.bat"+"'"
$mailsend ="& '"+$env:programmpfad+"Arbeitsdaten\Montagszugaenge\Mailsend.bat"+"'"
$crmpath= "\\svl-bdldb01\dialogcrm"
$files = $crmpath+"\Exporte\Adressfactory"

Add-Content $Logfile("##########################################################################")
Add-Content $Logfile ($logdate+": PS Script gestartet - Wochenverzeichnis wird erstellt")
echo $logdate + ": PS Script gestartet - Wochenverzeichnis wird erstellt"

#Wochenverzeichnis bei Vertrieb wird angelegt
if (!(Test-Path $dir))
{
    new-item $dir -itemtype directory 
    echo "Verzeichnis angelegt"
}

#Montagsversanddatei wird vom CRM-Datenbankserver heruntergeladen und in alte "1_Abos_fertig_.csv" umbenannt
Add-Content $Logfile ($logdate+": Download Daten von CRM-Server")
echo "Download Daten von CRM-Server"
foreach($file in Get-ChildItem $files)
{
 $dest=$env:programmpfad+"\temp_daten\"
 $path=$dest+$file.Name
 Copy-Item $file.FullName -Destination $dest
 Rename-Item -Path $path -NewName "1_Abos_fertig_.csv"
 Remove-Item ($file.FullName)
}

$dest=$env:programmpfad+"\temp_daten\"

#Datei von DCRM wird an Adressfactory geschickt
$logdate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
Add-Content $Logfile ($logdate+": Adressfactory-Upload")
echo "Adressfactory-Upload"
Add-Content $Logfile ("==================================")
Add-Content $Logfile( Invoke-Expression $AF_Upload )
Add-Content $Logfile ("==================================")

#Warten während Adressfactory Automatic Daten prüft
echo "slepp -s 120"
start-sleep -s 120

#Datei von Adressfactory downloaden
$logdate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
Add-Content $Logfile ($logdate+": Adressfactory-Download")
Add-Content $Logfile ("==================================")
Add-Content $Logfile( Invoke-Expression $AF_Download)
Add-Content $Logfile ("==================================")

#Datei entpacken
$logdate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
Add-Content $Logfile ($logdate+": Datei entpacken")
echo "Datei entpacken"
Add-Content $Logfile ("==================================")
Add-Content $Logfile( Invoke-Expression $unzip)
Add-Content $Logfile ("==================================")
start-sleep -s 5


#CSV-Convert und Rückkopie auf CRM-Server
$logdate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
$filestamp = Get-Date -Format "yyyyMMdd"
Add-Content $Logfile ($logdate+": Rückkopie auf CRM-Server")
echo "CSV-Convert und Rückkopie auf CRM-Server"
$targetFile = $dest+'1_Abos_rueck.csv'
$AfterAF=$crmpath+"\Schnittstellen\Adressfactory_Import\ScanFolder\1_Abos_rueck_"+$filestamp+'.csv'
Copy-Item $targetFile $AfterAF
Copy-Item $targetFile -Destination $dir

#Warten bis CRM Daten reimportiert und aufbereitet und auf Agenturen aufgeteilt hat
start-sleep -s 120


#Ergebnisdateien archivieren und auf Cloud bereitstellen
$logdate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
Add-Content $Logfile ($logdate+": Ergebnisdateien archivieren und auf Cloud bereitstellen")
echo "Ergebnisdateien archivieren und auf Cloud bereitstellen"
$ZugaengeRsults=$crmpath+"\Exporte\Montagsversand"
foreach($file in Get-ChildItem $ZugaengeRsults)
{
    if ($file.Name -match "CityDialog")
    {
        Copy-Item $file.FullName -Destination "\\svl-fil01\homeb$\MARKET\Versand\CityDialog\Zugaenge_2018\" 
        Copy-Item $file.FullName -Destination "\\svl-operating01\C$\Agenturportal_Vertrieb\Agenturportal\CityDialog\fromSV"
		Copy-Item $file.FullName -Destination "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten" 
		$dest = "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten\" + $file.Name
		$newpath = "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten\Montagsversand_CityDialog.csv"
		Copy-Item $dest $newpath
		Remove-Item $file.FullName
		
    } 
    
    if ($file.Name -match "Nikolaidis")
    {
        Copy-Item $file.FullName -Destination "\\svl-fil01\homeb$\MARKET\Versand\Nikolaidis\Zugaenge_2018\" 
        Copy-Item $file.FullName -Destination "\\svl-operating01\C$\Agenturportal_Vertrieb\Agenturportal\Nikolaidis\fromSV"
		Copy-Item $file.FullName -Destination "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten" 
		$dest = "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten\" + $file.Name
		$newpath = "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten\Montagsversand_Nikolaidis.csv"
		Copy-Item $dest $newpath
		Remove-Item $file.FullName
    }     
    
    if ($file.Name -match "Ruggiero")
    {

		
        Copy-Item $file.FullName -Destination "\\svl-fil01\homeb$\MARKET\Versand\Ruggiero\Zugaenge_2018\" 
        Copy-Item $file.FullName -Destination "\\svl-operating01\C$\Agenturportal_Vertrieb\Agenturportal\Ruggiero\"
		Copy-Item $file.FullName -Destination "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten" 
		$dest = "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten\" + $file.Name
		$newpath = "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten\Montagsversand_Ruggiero.csv"		
		if ($file.Name -match "Ruggiero_Teilabo")
		{
			$newpath = "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten\Montagsversand_Ruggiero_Teilabo.csv"	
		}
		Copy-Item $dest $newpath
		Remove-Item $file.FullName
    }     
    
    if ($file.Name -match "nicht gefunden")
    {
        Copy-Item $file.FullName -Destination $dir
		Copy-Item $file.FullName -Destination "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten" 
        Remove-Item $file.FullName
    }     
	
    Remove-Item "\\svl-fil01\homeb$\MARKET\Versand\OT_Daten\V-Beauftragte.csv"
    if ($file.Name -match "VADler")
    {
        Copy-Item $file.FullName -Destination "\\svl-fil01\homeb$\MARKET\Versand\OT_Daten\"
        $path = "\\svl-fil01\homeb$\MARKET\Versand\OT_Daten\"+$file.Name
		Copy-Item $file.FullName -Destination "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten" 
		$dest = "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten\" + $file.Name
		$newpath="\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten\V-Beauftragte.csv"
		Copy-Item $dest $newpath			
        Copy-Item $newpath -Destination "\\svl-fil01\homeb$\MARKET\Versand\OT_Daten\"	
        Remove-Item $file.FullName
    }   
	
	Remove-Item "\\svl-fil01\homeb$\MARKET\Versand\OT_Daten\OT_Daten.csv"
	if ($file.Name -match "BriefOT")
    {
        $path = "\\svl-fil01\homeb$\MARKET\Versand\OT_Daten\"+$file.Name
		Copy-Item $file.FullName -Destination "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten" 
		$dest = "\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten\" + $file.Name
		$newpath="\\svl-fil02\msg$\Msg allgemein\Vertrieb\temp_daten\OT_Daten.csv"	
		Copy-Item $dest $newpath
        Remove-Item $file.FullName
    }
}

# Mailversand getriggert über Regel in Account von Oli Jung

$logdate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
Add-Content $Logfile ($logdate+": Mail versenden")
echo "Mail versenden"
Add-Content $Logfile ("==================================")
Add-Content $Logfile( Invoke-Expression $mailsend)
Add-Content $Logfile ("==================================")


$logdate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
Add-Content $Logfile ($logdate+": Bereitstellung beendet")
Add-Content $Logfile("##########################################################################")
echo "Bereitstellung beendet"
