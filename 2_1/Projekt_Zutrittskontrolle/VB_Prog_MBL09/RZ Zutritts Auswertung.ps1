Function Using-Culture (
[System.Globalization.CultureInfo]$culture = (throw “USAGE: Using-Culture -Culture culture -Script {scriptblock}”),
[ScriptBlock]$script= (throw “USAGE: Using-Culture -Culture culture -Script {scriptblock}”))
{
    $OldCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
    trap 
    {
        [System.Threading.Thread]::CurrentThread.CurrentCulture = $OldCulture
    }
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
    Invoke-Command $script
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $OldCulture
}

Using-Culture de-IQ {get-date}
Get-Culture

Write-Host("Taste drücken zum fortfahren!")
Read-Host
#### Folder picker
Add-Type -AssemblyName System.Windows.Forms
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
    SelectedPath = $PSScriptRoot
    Description = "Ordner mit den Zutrittslisten auswählen"
}
[void]$FolderBrowser.ShowDialog()
$Basefolder = $FolderBrowser.SelectedPath

#Filepaths
#####TAG related files
$TagExcelFile = Get-ChildItem -Path $PSScriptRoot -Include prog.xlsm -Recurse
$PdfFiles = Get-ChildItem -Path $Basefolder -Include TAG*.pdf -Recurse
$baseCSV = $Basefolder+ "\TAG.csv"
$TagExportCSV = $Basefolder + "\TAGPrepared.csv"

#####ZAG related files
$ZAGFiles = Get-ChildItem -Path $Basefolder -Include ZAG*.lst -Recurse
$ZagExportCSV = $Basefolder +"\ZAGPrepared.csv"

#####QRZ related files
$QrzExcelFiles = Get-ChildItem -Path $Basefolder -Include QRZ*.xlsx -Recurse
$QrzExportCSV = $Basefolder + "\QRZprepared.csv"

#####Result related files
$SharepointExcelFile = Get-ChildItem -Path $Basefolder -Include Anmeldungen_Sharepoint.xlsx -Recurse
$outputCSV = $PSScriptRoot + "\Ausstehende Anmeldungen.csv"

#Generel Temp file
$tempCSV = $Basefolder +"\temp.csv"

##############################################################################################
#####################	BEGIN USER INPUT #####################################################
##############################################################################################

#date range input
Write-Host("Zeitraum eingeben")
$fromDate = Read-Host "Von (MM.YYYY) "
$toDate = Read-Host "Bis (MM.YYYY) "

##############################################################################################
#####################	END USER INPUT #######################################################
##############################################################################################


##############################################################################################
#####################	BEGIN TAG PROCESSING #################################################
##############################################################################################


Write-Host("Start verarbeitung TAG...")

############# Begin PDF to csv #################

#Create an excel object to export the PDFS to excel/csv
$excelApp = new-object -comobject excel.application
$workbook = $excelApp.workbooks.open($TagExcelFile.fullname)

Foreach($file in $PdfFiles)
{
   $worksheet = $workbook.worksheets.item(1)
   #PDF to excel Function
   $excelApp.Run("pdf_To_Excel_Word_Early_Binding",$file)
}

$excelApp.Run("saveSheetToCSV",$Basefolder)
$excelApp.Run("ClearWorksheet")

$workbook.save()
$workbook.close()

$excelApp.quit()

#Cleanup
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp)
[GC]::Collect()

############# END PDF to csv #################

############# Processing #################

Get-Content $baseCSV -Encoding:string | Select-String -Encoding utf8 `
-Pattern '^[0-9]+\.[a-zA-Z]*\.[0-9]+;[0-9:,]+;[0-9a-zA-Z-]+;[a-zA-Z0-9-üäöÜÄÖ. ]+;[a-zA-Z0-9._üäöÜÄÖ ]+;[0-9a-zA-ZüäöÜÄÖ_. ]+$' `
-AllMatches |  `
foreach {
    Add-Content $tempCSV $_.Matches.Value -Encoding UTF8
}

#Select the wanted columns and format the datetime
($rawCSV = import-Csv -Path $tempCSV -Encoding UTF7 -Delimiter ";" -Header "Datum","Zeit","Kartennr.","Inhaber","Ortstext","Firma" | `
    Select "Datum","Inhaber","firma" |  `
    sort Inhaber,Datum -Unique )| `
foreach {
    $_.Datum = [datetime]::parseexact($_.Datum, "dd.MMM.yy",[Globalization.CultureInfo]::CreateSpecificCulture('de-DE')).ToString('dd.MM.yyyy')
    $_.Inhaber = $_.Inhaber.Trim()
}

#adding additional csv field
$rawCSV = $rawCSV | Select-Object *,@{Name='RZ';Expression={'TAG'}}

#create formated csv
$rawCSV | Export-Csv $TagExportCSV -Delimiter ';' -NoTypeInformation -Encoding UTF8

############# Processing END #################

#cleanup
Remove-Item $tempCSV
Remove-Item $baseCSV

Clear-variable -Name "rawCSV"

##############################################################################################
#####################	END TAG PROCESSING ###################################################
##############################################################################################

##############################################################################################
#####################	BEGIN ZAG PROCESSING #################################################
##############################################################################################

Write-Host("Start verarbeitung ZAG...")

foreach ($file in $ZAGFiles) {

$date = ""

    Get-Content $file -Encoding:string | Select-String -Encoding utf8 `
    -Pattern '^ ?([0-9. ]+)? +[0-9:.]+ +[0-9]+ +[0-9a-zA-Z-]+ [0-9]+ +[a-zA-Z0-9.üöäÜÖÄ]+ [A-Z] +[0-9]+ +([a-zA-ZüöäÜÖÄ]+ [a-zA-ZüöäÜÖÄ]+) (?:[a-zA-Z0-9.üöäÜÖÄ ]*)? +0 +(?=Zutritt positiv +$)' |  `
    foreach {
			#date formating
        if (-not ([string]::IsNullOrEmpty(($_.Matches.Groups[1].ToString() -replace '\s','')))){
           $date = Get-Date ($_.Matches.Groups[1].ToString() -replace '\s','') -Format "dd.MM.yyyy"
        }
			
        $line = """$date"";""$($_.Matches.Groups[2])"";""ZAG"""

        Add-Content $tempCSV $line -Encoding UTF8 
    }
}

#Filter duplicates
$rawCSV = import-Csv -Path $tempCSV -Encoding UTF8 -Delimiter ";" -Header "Datum","Inhaber","RZ"| `
sort Inhaber,Datum,RZ -Unique 

$rawCSV | Export-Csv $ZagExportCSV -Delimiter ';' -NoTypeInformation -Encoding UTF8

#Cleanup
Remove-Item $tempCSV

Clear-variable -Name "rawCSV"

##############################################################################################
#####################	END ZAG PROCESSING ###################################################
##############################################################################################

##############################################################################################
#####################	BEGIN QRZ PROCESSING #################################################
##############################################################################################

Write-Host("Start verarbeitung QRZ...")

#Save excel to CSV
$E = New-Object -ComObject Excel.Application
$E.Visible = $false
$E.DisplayAlerts = $false
#opening all excels and saving them to own csv
$counter=1
foreach ($file in $QrzExcelFiles) {
    $wb = $E.Workbooks.Open($file)
    foreach ($ws in $wb.Worksheets)
    {
        $ws.SaveAs($Basefolder+"\workqrz"+$counter+".csv",6)
    }
    $E.Quit()
    $counter=$counter+1 
}
#merging multiple csvs into 1 for further processing
Get-ChildItem -Path $Basefolder -Filter workqrz*.csv | Select-Object -ExpandProperty FullName | Import-Csv | Export-Csv $tempCSV -NoTypeInformation -Append

($rawCSV = import-Csv -Path $tempCSV -Encoding UTF7 -Delimiter ","|  `
    where {$_.Ereignis -eq "Zutrittsbuchung erfolgreich (0)"} | `
    Select Zeitpunkt,Person  |  `
    sort Person,Zeitpunkt -Unique) | `
    foreach {
        $_.Person = $_.Person -replace '[,]',''
        $_.Zeitpunkt = [datetime]::parseexact($_.Zeitpunkt, "dd.MM.yyyy HH:mm:ss",[Globalization.CultureInfo]::CreateSpecificCulture('de-DE')).ToString('dd.MM.yyyy')
    }

#adding additional csv field
$rawCSV = $rawCSV | Select-Object *,@{Name='RZ';Expression={'QRZ'}}

#renaming headers
$rawCSV = $rawCSV | Select-Object @{ expression={$_.Zeitpunkt}; label='Datum' },@{ expression={$_.Person}; label='Inhaber' },RZ

##create formated csv
$rawCSV | sort Inhaber,Datum,RZ -Unique | Export-Csv $QrzExportCSV -Delimiter ';' -NoTypeInformation -Encoding UTF8


#cleanup
Remove-Item $tempCSV
For ($i=1;$i -lt $counter; $i++){
    $Path=$Basefolder+"\workqrz"+$i+".csv"
    Remove-Item $Path
}
Clear-variable -Name "rawCSV"


[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($E)
[GC]::Collect()


##############################################################################################
#####################	END QRZ PROCESSING ###################################################
##############################################################################################


##############################################################################################
#####################	BEGIN RESULT PROCESSING ##############################################
##############################################################################################

Write-Host("Erstelle Ergebnis csv...")


#Delete outputCsv isf present, since the programm only appends to the file
if (Test-Path $outputCSV) 
{
  Remove-Item $outputCSV
}

#Test date
#$fromDate = "8.2018"
#$toDate = "12.2018"

#Save excel to CSV
$E = New-Object -ComObject Excel.Application
$E.Visible = $false
$E.DisplayAlerts = $false
$wb = $E.Workbooks.Open($SharepointExcelFile)
foreach ($ws in $wb.Worksheets)
{
    $ws.SaveAs($tempCSV,6)
}
$E.Quit()

#Date conversion to dd.mm.yyyy format
($rawCSV = Import-Csv $tempCSV -Encoding UTF7 ) | `
    foreach { 
        $date = [datetime]::parseexact($_.Datum.ToString(),"m/d/yyyy",[Globalization.CultureInfo]::CreateSpecificCulture('de-DE'))
        $_.Datum = Get-date $date -Format "dd.mm.yyyy"
    }   

#applying date filter to only get the logs from the selected dates
$AnnouncedDCEntries = $rawCSV | ?{ ` #from-date block
       ( ( (Get-Date $_.Datum).Month -ge $fromDate.split('.')[0] ) -and ( (Get-Date $_.Datum).Year -ge $fromDate.split('.')[1] ) ) `
       -and #to-date block
       ( ( (Get-Date $_.Datum).Month -le $toDate.split('.')[0] ) -and ( (Get-Date $_.Datum).Year -le $toDate.split('.')[1] ) ) `
       }

#compares the exported sharepoint csv to the prepared CSV files
function compareCSV($CSVPath){
   $ActualDCEntries = Import-Csv $CSVPath -Delimiter ';' -Encoding UTF7

   $ActualDCEntries = $ActualDCEntries | ?{ ` #from-date block
       ( ( (Get-Date $_.Datum).Month -ge $fromDate.split('.')[0] ) -and ( (Get-Date $_.Datum).Year -ge $fromDate.split('.')[1] ) ) `
       -and #to-date block
       ( ( (Get-Date $_.Datum).Month -le $toDate.split('.')[0] ) -and ( (Get-Date $_.Datum).Year -le $toDate.split('.')[1] ) ) `
       }

   $ActualDCEntries |
        ForEach-Object {
            $ActualEntryRow = $_

            $ismatch = $false

            #True If Date matches AND if RZ matches AND if either the Person Matches or it is an A1 Employee
            $AnnouncedDCEntries |
                Foreach-Object { 
                If (($_.Datum -eq $ActualEntryRow.Datum ) -and 
                               ( $_.Rechenzentrum -eq $ActualEntryRow.RZ) -and 
                                   ( ( ($_."Name des Besuchers" -eq $ActualEntryRow.Inhaber) -or `
                                     (($_."Begleitende Personen inkl. Firmenname".ToLower() -like "*$($ActualEntryRow.Inhaber.split(' ')[0].ToLower())*") -and ($_."Begleitende Personen inkl. Firmenname".ToLower() -like "*$($ActualEntryRow.Inhaber.split(' ')[1].ToLower())*")) -or `
                                     (($_."2. Begleitende Person inkl. Firmenname".ToLower() -like "*$($ActualEntryRow.Inhaber.split(' ')[0].ToLower())*") -and ($_."3. Begleitende Person inkl. Firmenname".ToLower() -like "*$($ActualEntryRow.Inhaber.split(' ')[1].ToLower())*")) -or `
                                     (($_."3. Begleitende Person inkl. Firmenname".ToLower() -like "*$($ActualEntryRow.Inhaber.split(' ')[0].ToLower())*") -and ($_."3. Begleitende Person inkl. Firmenname".ToLower() -like "*$($ActualEntryRow.Inhaber.split(' ')[1].ToLower())*")) -or `
                                     (($_."4. Begleitende Person inkl. Firmenname".ToLower() -like "*$($ActualEntryRow.Inhaber.split(' ')[0].ToLower())*") -and ($_."4. Begleitende Person inkl. Firmenname".ToLower() -like "*$($ActualEntryRow.Inhaber.split(' ')[1].ToLower())*")) -or `
                                     (($_."5. Begleitende Person inkl. Firmenname".ToLower() -like "*$($ActualEntryRow.Inhaber.split(' ')[0].ToLower())*") -and ($_."5. Begleitende Person inkl. Firmenname".ToLower() -like "*$($ActualEntryRow.Inhaber.split(' ')[1].ToLower())*"))
                                     ) -or (($ActualEntryRow.Inhaber -eq "A1-Mitarbeiter") -and ($_."Name der Firma" -eq "A1TA")) #This line checks if the entry was made by A1
                                   )
                               
                               ){
                   #If person announced the entry set to true
                   #If the person used the "RZ LEIHKARTE" then filter it out for manuel confirmation
                   if ( $ActualEntryRow.Inhaber -ne "RZ LEIHKARTE" ){
                     $ismatch= $true
                   }


                   
                }
            }
                if (!$ismatch){
                    #Add Not announced entries to CSV
                    $ActualEntryRow | Export-Csv -Path $outputCSV -Append -Delimiter ";" -NoTypeInformation -Encoding Default
            }
            
   }
}

#Prepare CSV
Add-Content $outputCSV '"Datum";"Inhaber";"RZ"'

#Checking for not Announced entries
compareCSV $TagExportCSV
compareCSV $QrzExportCSV
compareCSV $ZagExportCSV

#Clean up
Remove-Item $tempCSV
Remove-Item $QrzExportCSV
Remove-Item $TagExportCSV
Remove-Item $ZagExportCSV

Add-Content $outputCSV "Erstellt von $env:UserName, Am $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')"

Write-Host("Verarbeitung Fertig!")
Write-Host("Das Ergebnis steht im File 'Ausstehende Anmeldungen.csv'")
Read-Host -Prompt "Taste Drücken zum Beenden."

##############################################################################################
#####################	END RESULT PROCESSING ################################################
##############################################################################################