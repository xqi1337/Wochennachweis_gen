# Wochennachweis_gen

Installieren der Pakete aus der requirements.txt (auf einem anderen Computer oder in einer anderen Umgebung):

bash
Kopieren
pip install -r requirements.txt



Wochennachweis Ausfüllen
Dieses Python-Projekt ermöglicht es, einen Wochennachweis (z. B. für ein Praktikum oder eine Ausbildung) automatisch auszufüllen, basierend auf den Daten, die der Benutzer über eine benutzerfreundliche GUI eingibt. Die Anwendung nutzt verschiedene Bibliotheken, um Excel-Tabellen zu bearbeiten, Web-Inhalte zu extrahieren und die Ausgabe in einer strukturierten Form zu speichern. Zusätzlich wird beim Abschluss der Anwendung ein MP3-Sound abgespielt.

Features:
Benutzeroberfläche (GUI): Eine einfache, grafische Benutzeroberfläche ermöglicht es dem Benutzer, Name, Vorname, Kalenderwoche und Links zu verschiedenen Tagen der Woche (Montag bis Freitag) einzugeben.
Dynamisches Ausfüllen von Excel: Die Anwendung verwendet eine Excel-Vorlage, füllt sie basierend auf den Benutzereingaben und generiert einen ausgefüllten Wochennachweis.
Web Scraping: Mit der BeautifulSoup-Bibliothek wird aus den eingegebenen URLs für jeden Tag die Überschrift (z. B. aus einem Inhaltsverzeichnis) extrahiert und als Links in das Excel-Dokument eingefügt.
Audio-Wiedergabe: Nachdem der Wochennachweis erstellt wurde, wird eine MP3-Datei abgespielt, um den erfolgreichen Abschluss des Prozesses anzuzeigen.
JSON-Speicherung: Die Daten, die aus den Links extrahiert werden, werden in einer headings.json-Datei gespeichert.

Abhängigkeiten:
requests – Zum Abrufen von Web-Inhalten über HTTP.
BeautifulSoup (bs4) – Für das Web Scraping von HTML-Daten.
ttkbootstrap – Für eine moderne und anpassbare GUI mit Tkinter.
openpyxl – Zum Bearbeiten von Excel-Dateien.

Installation:
Klone das Repository: 

git clone https://github.com/xqi1337/Wochennachweis_gen.git

    cd repository
    
2. Installiere die notwendigen Pakete:
    
bash
    
    pip install -r requirements.txt
    
3. Stelle sicher, dass du ein geeignetes Python-Environment verwendest.

## Verwendung:

1. Starte das Skript:

bash

    
    python main.py
    

3. Gib die erforderlichen Informationen über die GUI ein:
   - Name, Vorname, Kalenderwoche, KW von, KW bis
   - Links zu den Tagen der Woche (Montag bis Freitag)

4. Klicke auf "Wochennachweis erstellen", um die Excel-Datei zu generieren.

5. Der Wochennachweis wird als Excel-Datei gespeichert und kann bei Bedarf weiterverwendet werden.

## Anpassung der Vorlage:
- Stelle sicher, dass du eine Excel-Vorlage (Vorlage_Wochennachweis_9h.xlsx) auf deinem Desktop speicherst, um den Wochennachweis korrekt ausfüllen zu können.

## Hinweise:
- Das Skript verwendet die beautifulsoup4-Bibliothek für Web Scraping. Falls beautifulsoup4 nicht installiert ist, wird es automatisch mit pip installiert.
- Die MP3-Datei cute-uwu.mp3` muss auf dem Desktop vorhanden sein oder kann nach Bedarf angepasst werden.
