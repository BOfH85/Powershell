###################################################
#
#
# Umbau der Bereitstellung und Steuerung des 
# KüRü-Abrufs im Zuge der Umstellung auf 
# Dialog CRM von alten Batch-Skripten mit Access
# auf PowerShell
#
#
# (c) 01.02.2018 Oliver Jung
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
$MonthString = Get-Date -UFormat %b
$MS = "{0:D2}" -f [int]$Datum.month
$dir = "\\svl-fil01\homeb$\MARKET\Versand\"+$Datum.year+"\"+$Datum.year+"_Kuendiger\"+$Datum.year+"_"+$MS+"_"+$MonthString+"\"
$dest=$env:programmpfad+"\KueRue\"
$destarchiv =$dest+$Datum.year+"_"+$MS+"_"+$MonthString+"\"
$dirKornmann="\\svl-fil01\homeb$\MARKET\Versand\"
$logdate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
$Logfile=$env:programmpfad+"KueRue_log"+$datum.year+".txt"
$mailsend ="& '"+$env:programmpfad+"Arbeitsdaten\KüRüMails.bat"+"'"
$crmpath= "\\svl-bdldb01\dialogcrm"
$files_kuerue = $crmpath+"\Exporte\KüRü"
$files_ehem = $crmpath+"\Exporte\KüRü-Ehemalige"

Add-Content $Logfile("##########################################################################")
Add-Content $Logfile ($logdate+": PS Script gestartet - Monatsverzeichnis wird erstellt")
echo $logdate + ": PS Script gestartet - Monatsverzeichnis wird erstellt"

#Monatsverzeichnis bei Vertrieb wird angelegt
if (!(Test-Path $dir))
{
    new-item $dir -itemtype directory 
    echo "Verzeichnis H-Laufwerk angelegt"
}

if (!(Test-Path $destarchiv))
{
    new-item $destarchiv -itemtype directory 
    echo "Verzeichnis MSG angelegt"
}

#KüRü-Datei wird vom CRM-Datenbankserver heruntergeladen und umbenannt
Add-Content $Logfile ($logdate+": KüRü-Datei wird vom CRM-Datenbankserver heruntergeladen und umbenannt")
echo "KüRü-Datei wird vom CRM-Datenbankserver heruntergeladen und umbenannt"
foreach($file_kuerue in Get-ChildItem $files_kuerue)
{

 $path=$dest+$file_kuerue.Name
 $path2=$dest+"KueRue.csv"
 Copy-Item $file_kuerue.FullName -Destination $dest
 if ($file_kuerue.Name -match "VADler")
{
	Copy-Item $path $path2
	Copy-Item $path2 -Destination $dir
	Copy-Item $path2 -Destination $dirKornmann
	Copy-Item $path2 -Destination $destarchiv
	Remove-Item $path2
}

 Copy-Item $path -Destination $dir
 Copy-Item $path -Destination $destarchiv
 Remove-Item $path
 Remove-Item ($file_kuerue.FullName)
}


#KüRü-Ehemamlige Datei wird vom CRM-Datenbankserver heruntergeladen und umbenannt
Add-Content $Logfile ($logdate+": KüRü-Ehemamlige Datei wird vom CRM-Datenbankserver heruntergeladen und umbenannt")
echo "KüRü-Ehemamlige Datei wird vom CRM-Datenbankserver heruntergeladen und umbenannt"
foreach($files_ehem in Get-ChildItem $files_ehem)
{
 $path=$dest+$files_ehem.Name
 $path2=$dest+"KueRue_Ehem.csv"
 Copy-Item $files_ehem.FullName -Destination $dest
 if ($files_ehem.Name -match "VADler")
{
	Copy-Item $path $path2
	Copy-Item $path2 -Destination $dir
	Copy-Item $path2 -Destination $dirKornmann
	Copy-Item $path2 -Destination $destarchiv
	Remove-Item $path2
}

 Copy-Item $path -Destination $dir
 Copy-Item $path -Destination $destarchiv
 Remove-Item $path
 Remove-Item ($files_ehem.FullName)
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