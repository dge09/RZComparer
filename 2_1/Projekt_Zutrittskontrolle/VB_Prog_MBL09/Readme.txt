Diese Programm ben�tigt besteht aus 2 teilen.
1. Das Powershell Script "RZ Zutritts Auswertung"
2. Das excel mit macro "prog"

Die zu verarbeitenden Dateien k�nnen in einem Beliebigen Ordner sein.
Wichtig ist das alle Dateien die Verarbeitet werden sollen im GLEICHEN Ordner sind.

Ben�tigte Dateien:

1. Sharepoint liste (mit dem namen Anmeldung_Sharepoint.xlsx
2. TAG zutrittslisten. Die dateien m�ssen mit TAG prefixed sein und mit .pdf enden
3. ZAG zutrittslisten. Die dateien m�ssen mit ZAG prefixed sein und mit .lst enden
4. QRZ zutrittslisten. Die dateien m�ssen mit QRZ prefixed sein und mit .xlsx enden

USAGE:

Zum ausf�hren einfach das Powershell script starten. --> Windows Startmenue --> "Powershell" ausw�hlen --> �ffnet CMD-Window
Anschlie�end wird man aufgefordert einen Ordner zu w�hlen, dieser soll die Zutrittslisten + sharepoint excel beinhalten.
Danach muss nur noch der Zeitraum im Format "Monatszahl.Jahreszahl" 

WICHTIG!!!
Beim Monat darf keine 0 vorne dran sein, zb. kein 09 f�r den September

Nach dem Enter dr�cken startet die Verarbeitung. 
Beim ersten mal starten kann es sein das Word eine Meldung bringt die mit "ok" zu best�tigen ist.
Wenn die Verarbeitung abgeschlossen ist wird die csv Datei "Ausstehende Anmeldungen" im Ordner wo das Script liegt generiert.

Known problems:

Das Script muss auf einer Deutschen windows installation gestartet werden. Da sonst die CultureInfo nicht passt und es bei der Formatierung des Datums probleme gibt!

Hin und wieder kommt es vor das es beim TAG verarbeitungsschritt zu eine Problem kommt.
Wenn Eine Meldung von Excel kommt wo man nur "beenden"(oder so) oder "Debuggen" klicken kann, 
muss man das Script anhalten und �ber den Task Manager "Winword" und "Excel" Schlie�en.
anschlie�end sollte es ohne Probleme laufen.

Das obige Problem tritt auf, wenn das Script �fters ausgef�hrt wird. 
Da manchmal Excel nicht sauber geschlossen wird und somit noch ein File Offen hat das gebraucht wird.
