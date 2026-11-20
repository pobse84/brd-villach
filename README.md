# brd-villach

Zum automatisierten Auslesen der Beteiligung/Mitarbeit in der OS:

Für jedes Jahr einen Ordner erstellen und dann jede Kategorie einzeln für jedes Jahr exportieren.

1) **Einsätze**: BRIS: Berichte: Einsatzbeteiligung: Schnellauswahl: Jahr (z.B. 2025) auswählen. Achten, dass die richtige Ortsstelle ausgewählt ist. Dann exportieren (Excel, alle Spalten) auswählen und als "Einsätze.xlsx" speichern.
2) **Übungen**: BRIS: Berichte: Veranstaltungen: Schnellauswahl: Jahr (z.B. 2025) auswählen. Dann nach Klasse "Übung" filtern. Dann "exportieren (Excel, alle Spalten)" und als "Übungen.xlsx" speichern.
3) **Dienste**: Sollte per "Dienst" filter gehen aber die Daten sind unvollständig und man sollte die Wintersaison (Oktober-Juni) filtern und nicht das Kalenderjahr...
4) **Übungstour**: Übungen: BRIS: Berichte: Veranstaltungen: Schnellauswahl: Jahr (z.B. 2025) auswählen. Dann nach Klasse "Übungstour" filtern. Dann "exportieren (Excel, alle Spalten)" und als "Übungstour.xlsx" speichern.

Im Python Skript den Pfad zum Ordner angeben.
Auswertung starten mit Python 3.13