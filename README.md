# Trauerfeier Word-Dokument Generator

Dieses Python-Skript automatisiert das Erstellen von personalisierten Programmheften für Trauerfeiern aus einer Word-Vorlage (`.docx`). Es ersetzt Platzhalter für den Namen und geschlechtsspezifische Pronomen und Begriffe, um schnell ein korrektes Dokument zu generieren.

---

## Funktionsweise

Das Skript liest eine Word-Datei namens `Trauerfeier_Roh.docx`, die spezielle Platzhalter enthält. Anhand von Kommandozeilen-Argumenten für den **Namen** und das **Geschlecht** der verstorbenen Person werden diese Platzhalter durch die passenden Texte ersetzt. Die neue, personalisierte Datei wird anschließend in einem separaten `Output`-Ordner gespeichert.

---

## Voraussetzungen

- **Python 3**: Das Skript ist für Python 3 geschrieben. Auf macOS und den meisten modernen Systemen ist dies bereits installiert.
- **`python-docx` Bibliothek**: Diese Bibliothek wird benötigt, um Word-Dateien zu bearbeiten.

Installiere die Bibliothek mit folgendem Befehl im Terminal:
```shell
pip3 install python-docx