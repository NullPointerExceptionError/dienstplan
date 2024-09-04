# Dienstplan-Generator
## Zusammenfassung
Dieses Programm ist ein Dienstplan-Generator, speziell ein Messdienerplan-Generator.

## Installation
Neueste Version unter [Releases](https://github.com/NullPointerExceptionError/dienstplan/releases/) herunterladen (`dienstplan_generator_vX.Y.zip`) und Ordner entpacken

## Hinweise zur Benutzung
### Benutzung
1. Einzuteilende Personen in `personenliste.txt` schreiben (nur bei erster Benutzung oder Änderungen)
2. Termine (Daten & Uhrzeiten), Anlass und Anzahl einzuteilender Personen für den Plan in `daten.txt` schreiben
3. Termine, an denen Personen nicht können in `terminausnahmen.txt` schreiben
4. `dienstplan.exe` ausführen (beim ersten Starten unter Windows kommt eine blaue Sicherheitswarnung. Dort auf "Weitere Informationen" und dann auf "Trotzdem ausführen" klicken)
5. Falls ein Plan generiert werden konnte, werden die Häufigkeiten der Einteilungen je Nachname angezeigt. Hinweise und Fehlermeldungen werden ebenfalls angezeigt
6. Programm kann mit `Strg`+`C` oder Schließen des Fensters beendet werden

### Formatierungen
#### `personenliste.txt`
```
Vorname1 Nachname1
Vorname2 Nachname2
...
```
- mit Leerzeichen getrennt
- bei Familienmitgliedern ist gleicher Nachname erlaubt -> werden immer zusammen eingeteilt
- maximal 1 Leerzeile am Ende der Datei
#### `daten.txt`
```
03.03.2024	11.00  Hl. Messe  4
10.03.2024	11.00  Hl. Messe	4
17.03.2024	11.00  Hl. Messe	4
24.03.2024	11.00  Hl. Messe/\nPalmsonntag   3
30.05.2024  11.00  Hl. Messe/\nFronleichnam  0
```
- `Datum  Uhrzeit  Anlass  AnzahlEinzuteilenderPersonen`
- mit Tab getrennt (Leerzeichen bei Anlass erlaubt)
- in Reihenfolge aufgelistet, in der sie später im Plan sein soll
- Daten können auch in anderem Format (z.B. 03.03.24) geschrieben werden, solange sie exakt mit denen in `terminausnahmen.txt` übereinstimmen
- \n für Zeilenumbruch in Anlass
- 0 für Anzahl einzuteilender Personen, falls später im Dokument "freiwillig" oder "alle" manuell hinzugefügt wird
#### `terminausnahmen.txt`
```
NAchname1  17.03.2024, 05.05.2024
Nachname2  03.03.2024, 10.03.2024, 02.06.2024
Nachname3  10.03.2024
...
```
- muss nicht alphabetisch sortiert sein
- Daten müssen exakt mit denen in `daten.txt` übereinstimmen (xx.xx.24 ist nciht mit xx.xx.2024 kompatibel)
- Nachname von Daten mit Tab getrennt, Daten selbst mit `, ` (Komma und Leerzeichen) getrennt
- Personen, die immer können nicht in Datei schreiben

### sonstige Hinweise
- `dienstplan.exe` im Ordner mit den drei .txt-Dateien ausführen
- vor Ausführen der `dienstplan.exe` das Word-Dokument schließen
- bei Fehlermeldungen korrekte Formatierungen überprüfen

## Funktionsweise
Zuerst werden alle Personen eingeteilt, die nur an sehr wenigen Terminen können.
Dann wird für jeden Termin ein Topf aus allen Personen erzeugt, die an dem Termin können.
Dieser Topf wird wie folgt so lange gefiltert, bis nur noch eine Person (bzw. mehrere Personen einer Familie) übrigbleibt:
- Personen, die an dem Termin können (und nicht zu den bereits verteilten Sonderfällen mit sehr wenigen Terminen gehören)
- davon: Im Falle von Geschwistern werden sie aussortiert, falls sie mehr sind als Slots übrig
- davon: Personen, die bisher am wenigsten eingeteilt waren
- davon: Personen, die am längsten nicht mehr dran waren
- davon: Personen, die an den wenigsten Terminen können
- davon: zufällige Person
Das wird so oft durchgeführt, bis alle Slots eines Termins belegt sind.
