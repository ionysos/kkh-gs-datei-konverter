# Datei Konverter
Datei Konverter als Excel Add-In zur Nutzung im KKH GS

## Installation
### Windows
Kopiere die Excel Add-In Datei `DateiKonverter.xlam` aus dem Unterverzeichnis `ribbon` nach `C:\Users\<UserName>\AppData\Roaming\Microsoft\AddIns` (`<UserName>` im Pfad muss entsprechend angepasst werden).
Starte danach Excel (neu) und wähle über den Tab `Entwicklertools` in der Gruppe `Add-Ins` den Button `Excel-Add-Ins`.
Aktiviere anschließend dort den `Datei Konverter` und bestätige dies mit betätigen des `OK`-Buttons.
Nun gibt es ein zusätzliches Excel Tab für den `Datei Konverter`.

> Der Excel Tab `Entwicklertools` kann über `Datei` > `Optionen` > `Menüband anpassen` durch aktivieren bzw. deaktivieren der `Entwicklertools`, im rechten Auswahlfeld des Optionen-Fensters, ein- bzw. ausgeblendet werden. Der `OK`-Button bestätigt auch hier die Einstellung.

## Deinstallation
### Windows
Lösche die Excel Add-In Datei `DateiKonverter.xlam` aus dem Verzeichnis `C:\Users\<UserName>\AppData\Roaming\Microsoft\AddIns` (`<UserName>` im Pfad muss entsprechend angepasst werden).
Starte danach Excel (neu), bestätige (eventuell) die Dialog-Box "Wir konnten '`C:\Users\<UserName>\AppData\Roaming\Microsoft\AddIns\DateiKonverter.xlam`' nicht finden. Wurde das Objekt vielleicht verschoben, umbenannt oder gelöscht?" (`<UserName>` im Pfad variiert entsprechend) mit `OK` und wähle über den Tab `Entwicklertools` in der Gruppe `Add-Ins` den Button `Excel-Add-Ins`.
Deaktiviere anschließend dort den `Dateikonverter`, bestätige die Dialog-Box "Kann den Add-In '`C:\Users\<UserName>\AppData\Roaming\Microsoft\AddIns\DateiKonverter.xlam`' nicht finden. Aus der Liste löschen?" (`<UserName>` im Pfad variiert entsprechend) mit `Ja` und bestätige dies mit betätigen des `OK`-Buttons.
Nun gibt es kein zusätzliches Excel Tab für den `Datei Konverter` (mehr).

> Der Excel Tab `Entwicklertools` kann über `Datei` > `Optionen` > `Menüband anpassen` durch aktivieren bzw. deaktivieren der `Entwicklertools`, im rechten Auswahlfeld des Optionen-Fensters, ein- bzw. ausgeblendet werden. Der `OK`-Button bestätigt auch hier die Einstellung.

## System-Voraussetzungen
Typ | Wert
--- | ---
Betriebssystem: | Windows oder macOS
Software: | Excel 2007 oder höher

### Links
- https://www.vba-tutorial.de/
- https://trumpexcel.com/excel-add-in/
- https://docs.microsoft.com/de-de/office/vba/library-reference/concepts/customize-the-office-fluent-ribbon-by-using-an-open-xml-formats-file
