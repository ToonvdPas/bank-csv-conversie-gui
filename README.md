# Grafische gebruikersinterface voor de programma's rabo-csv.py en ing-csv.py
Dit programma biedt een grafische schil over de programma's ```rabo-csv.py``` en ```ing-csv.py```<br>
Om dit programma te kunnen gebruiken dien je dus ook één van die twee programma's te installeren.<br>
Raadpleeg daarvoor de betreffende repositories.

Het prgramma werkt op Linux en Windows.
![Screenshot van het openingswindow](Files/bank-csv-conversie-window.png)

Om het programma te gebruiken dien je eenmalig de configuratie in te voeren en op te slaan.  Bij volgende keren volstaat het dan om die configuratie te laden.  De configuratie-items zijn: ```script name```, ```infile```, ```matchfile```, ```outdir```, ```logfile``` en ```verbosity```.  Voor uitleg van deze parameters verwijs ik je naar het repository van ```rabo-csv.py``` of ```ing-csv.py```.

> [!NOTE]
> Het veld ```Script name``` dient alléén de bestandsnaam van het script te bevatten, dus **zonder** pad ervoor!  Dat script dient in dezelfde directory te staan als dit programma (```bank-csv-conversie-gui.py```).

Wanneer de configuratie geladen is kan de conversie worden gestart via de knop ```Run CSV Conversion```.  Het script ```Script name``` wordt nu uitgevoerd, waarbij per eigen bankrekening een apart uitvoerbestand in CSV-formaat wordt aangemaakt.  Een eventueel bestaande logfile wordt overschreven.

Na voltooiing scant dit programma de logfile en produceert een handig overzicht van mutaties die niet gematched zijn, en dus niet gerubriceerd konden worden.  Aan de hand van dit overzicht kun je makkelijk de ontbrekende rules toevoegen aan de ```matching-rules.csv```.

# Commandline-argumenten
```
options:<br>
  -h, --help             show this help message and exit<br>
  --cfgfile CFGFILE      initial INI-file, handy when starting up from a BAT-file<br>
```
