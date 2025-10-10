### **Produktdokument (PRD): DocxTpl Automatisierungstool v7.1**

**Version:** 1.1
**Datum:** 2. Juli 2024
**Autor:** Gemini (basierend auf dem Projekt von Jonas Beseler)

---

#### **1. Einleitung**

##### **1.1 Produktname**
DocxTpl Automatisierungstool

##### **1.2 Zweck**
Dieses Tool automatisiert die serienmäßige Erstellung von Microsoft Word-Dokumenten (`.docx`) basierend auf standardisierten Vorlagen und dynamischen Daten aus einer Microsoft Excel-Datei. Sein Hauptzweck ist die drastische Reduzierung des manuellen Aufwands und der Fehleranfälligkeit bei der Erstellung von umfangreichen, sich wiederholenden Dokumentensätzen für Projekte.

##### **1.3 Zielgruppe**
Die Zielgruppe sind Projektmanager, technische Administratoren, Ingenieure und Sachbearbeiter, die regelmäßig große Mengen an standardisierten Dokumenten erstellen müssen (z. B. für Windparks, Bauprojekte, Anlagenzertifizierungen). Die Benutzer sind mit der Strukturierung von Daten in Excel und der Arbeit mit Dateivorlagen vertraut, sind aber keine Softwareentwickler. Sie benötigen eine zuverlässige, einfach zu bedienende grafische Oberfläche, um den Generierungsprozess zu steuern.

---

#### **2. Ziele und Erfolgskriterien**

| Ziel | Erfolgskriterium |
| :--- | :--- |
| **Zeitersparnis** | Die Generierung eines kompletten Dokumentensatzes für ein Projekt mit >10 Einzelanlagen dauert weniger als 5 Minuten. |
| **Fehlerreduktion** | Copy-Paste-Fehler werden eliminiert. Konfigurationsfehler (fehlende Bilder/Platzhalter) werden vor der Generierung durch einen Trockenlauf aufgedeckt. |
| **Konsistenz** | Alle generierten Dokumente folgen exakt den vorgegebenen Vorlagen und Namenskonventionen. |
| **Benutzerfreundlichkeit** | Ein Anwender kann das Tool nach einer kurzen Einarbeitung ohne Entwicklerunterstützung bedienen. Die GUI gibt durchgängig visuelles Feedback. |

---

#### **3. Benutzer und Anwendungsfälle**

##### **3.1 Benutzer-Persona: Der Projektadministrator**
*   **Rolle:** Verantwortlich für die Erstellung und Verwaltung der gesamten Projektdokumentation.
*   **Aufgaben:** Sammelt Daten, pflegt die Master-Excel-Liste, stellt sicher, dass alle notwendigen Schilder, Betriebsanweisungen und Notfallpläne für jede einzelne Anlage eines Projekts korrekt erstellt werden.
*   **Motivation:** Möchte den Prozess so effizient und fehlerfrei wie möglich gestalten.

##### **3.2 Anwendungsfall 1: Konfiguration prüfen (Trockenlauf)**
1.  **Vorbereitung:** Der Benutzer hat die Pfade zur Excel-Datei und zum Vorlagen-Ordner konfiguriert.
2.  **Start der Prüfung:** Er klickt auf den Button "Konfiguration prüfen".
3.  **Verarbeitung:** Das Tool simuliert einen Generierungslauf. Es prüft, ob alle in den Word-Vorlagen verwendeten Platzhalter in der Excel-Datei existieren und ob alle in den Excel-Daten referenzierten Bilddateien im Bilder-Ordner vorhanden sind. Es werden keine Dokumente geschrieben.
4.  **Ergebnis:** Das Log-Fenster füllt sich mit den Ergebnissen der Prüfung. Der Benutzer erhält am Ende eine Zusammenfassung, ob die Konfiguration gültig ist oder ob Fehler gefunden wurden.

##### **3.3 Anwendungsfall 2: Generierung eines vollständigen Dokumentensatzes**
1.  **Vorbereitung:** Der Benutzer hat idealerweise einen erfolgreichen Trockenlauf durchgeführt. Alle Pfade sind konfiguriert und validiert (grün hinterlegt).
2.  **Start des Tools:** Er klickt auf den "Start"-Button.
3.  **Verarbeitung:** Das Tool arbeitet die Daten und Vorlagen asynchron ab. Die GUI bleibt bedienbar. Ein laufender Prozess kann über den "Abbrechen"-Button sicher gestoppt werden.
4.  **Abschluss:** Nach Abschluss erhält der Benutzer eine Erfolgsmeldung. Ein Button "Export-Ordner öffnen" erscheint, mit dem er direkt zum Ausgabeordner navigieren kann.

---

#### **4. Funktionsanforderungen (FR)**

##### **FR-01: Daten-Input (Excel)**
*   Das Tool muss eine `.xlsx`-Datei als Datenquelle einlesen.
*   **Blatt 1 (Datensätze):** Enthält die anlagenspezifischen Daten. Jede Zeile mit einer ausgefüllten `anlage_seriennummer` wird als ein separater Datensatz behandelt. Die Kopfzeile ist über die GUI konfigurierbar.
*   **Blatt 2 (Fallback-Marken):** Enthält allgemeine, projektweite Daten (z. B. Projektname). Wenn ein Wert in einem Datensatz aus Blatt 1 fehlt, wird der entsprechende Wert aus Blatt 2 als Fallback verwendet.

##### **FR-02: Vorlagen-Verarbeitung**
*   Das Tool durchsucht rekursiv einen angegebenen Vorlagenordner nach `.docx`-Dateien.
*   **Anlagenspezifische Vorlagen:** Vorlagen im Unterordner `Anlagen` werden für jeden einzelnen Datensatz aus der Excel-Datei generiert.
*   **Allgemeine Vorlagen:** Vorlagen im Unterordner `Allgemein` werden nur einmal pro Programmdurchlauf generiert, wobei die Daten aus dem ersten Datensatz und den Fallback-Marken verwendet werden.

##### **FR-03: Platzhalter-Ersetzung (docxtpl)**
*   Das Tool verwendet `docxtpl`, um Jinja2-Platzhalter (z. B. `{{ projekt_name }}`) zu ersetzen und behält die Formatierung bei.
*   **Sonderfall Zeitstempel:** Der Platzhalter `{{ datetime_utc }}` wird automatisch mit dem aktuellen UTC-Zeitstempel ersetzt. Das genaue Format dieses Zeitstempels ist über ein eigenes Feld in der GUI frei konfigurierbar.

##### **FR-04: Dynamische Bild- und QR-Code-Ersetzung**
*   **Bilder:** Ein Platzhalter mit dem Suffix `_img` (z. B. `{{ lageplan_img }}`) wird durch eine Bilddatei ersetzt. Der Dateiname des Bildes wird aus der Excel-Zelle gelesen.
*   **QR-Codes:** Platzhalter mit den Suffixen `_qr` oder `_link` (z. B. `{{ anmeldung_qr }}`) werden durch einen QR-Code ersetzt, der den Text aus der entsprechenden Excel-Zelle kodiert.
*   **Größensteuerung:** Die Breite der eingefügten Bilder und QR-Codes kann über einen zusätzlichen Platzhalter mit dem Suffix `_size` (z. B. `{{ lageplan_img_size }}`) in Zentimetern gesteuert werden.
*   **SVG-Unterstützung:** SVG-Bilder, die in der Excel-Datei referenziert werden, werden automatisch in PNG konvertiert, um sie in Word einbetten zu können.

##### **FR-05: Ausgabe-Management**
*   Für jeden Programmdurchlauf wird ein eindeutiger Haupt-Exportordner erstellt, benannt nach dem Schema `JJJJ-MM-TT_HHMMSS_Projektname`.
*   **Kategorisierung:** Innerhalb des Haupt-Exportordners werden Unterordner basierend auf den in der GUI konfigurierten Kategorien erstellt (z. B. "Beschilderung", "Betriebsanweisung"). Die Zuordnung erfolgt über Präfixe im Dateinamen der Vorlage (z. B. `B_` für Beschilderung).
*   **Dateibenennung:**
    *   Anlagenspezifische Dokumente: `{seriennummer}_{vorlagenname_ohne_präfix}.docx`
    *   Allgemeine Dokumente: `{vorlagenname_ohne_präfix}.docx`
*   **Direktes Öffnen:** Nach einer erfolgreichen Generierung wird ein "Export-Ordner öffnen"-Button sichtbar, der den Ausgabeordner im System-Explorer öffnet.

##### **FR-06: Grafische Benutzeroberfläche (GUI)**
*   **FR-6.1: Tab-Struktur:** Die GUI ist in die Tabs "Hauptsteuerung", "Kategorien" und "Logs" unterteilt.
*   **FR-6.2: Pfad- und Einstellungs-Inputs:**
    *   Bietet Eingabefelder und Buttons für alle Pfade.
    *   Bietet ein zusätzliches Eingabefeld zur Definition des Formats für `datetime_utc`.
*   **FR-6.3: Asynchrone Verarbeitung:** Der Generierungsprozess läuft in einem separaten Thread. Ein laufender Prozess kann über den "Abbrechen"-Button sicher gestoppt werden, ohne die Anwendung zu schließen.
*   **FR-6.4: Fortschrittsanzeige:** Ein Fortschrittsbalken und ein Textlabel zeigen den aktuellen Status der Verarbeitung an.
*   **FR-6.5: Live-Logs:**
    *   Alle Aktionen werden in Echtzeit im "Logs"-Tab angezeigt.
    *   Zur besseren Lesbarkeit werden Dateinamen **fett** und Platzhalter-Listen `lila` dargestellt. Vor `ERROR`- und `FATAL`-Meldungen wird ein Zeilenumbruch eingefügt, um sie optisch hervorzuheben.
*   **FR-6.6: Datenvorschau:** Die Daten aus dem ersten Blatt der ausgewählten Excel-Datei werden in einer Tabelle angezeigt.
*   **FR-6.7: Konfigurationspersistenz:** Alle Pfade, Einstellungen und Kategorien werden beim Schließen der Anwendung in einer `settings.json`-Datei gespeichert und beim nächsten Start automatisch geladen.
*   **FR-6.8: Anpassbare Themes:** Themes werden aus externen `.qss`-Dateien geladen.
*   **FR-6.9: Echtzeit-Pfadvalidierung:** Pfad-Eingabefelder ändern ihre Hintergrundfarbe (grün/rot), um sofortiges visuelles Feedback über die Gültigkeit des Pfades zu geben.
*   **FR-6.10: Trockenlauf-Funktion:** Ein "Konfiguration prüfen"-Button startet eine Validierung, die alle Vorlagen und Bilder prüft, ohne Dokumente zu schreiben, und die Ergebnisse im Log anzeigt.

---

#### **5. Nicht-funktionale Anforderungen (NFR)**

| Anforderung | Beschreibung |
| :--- | :--- |
| **Performance** | Die GUI bleibt jederzeit reaktionsschnell. SVG-Konvertierung wird gecacht. |
| **Zuverlässigkeit** | Robuste Fehlerbehandlung. Die Trockenlauf-Funktion ermöglicht das Abfangen von Konfigurationsfehlern vor dem Start. |
| **Kompatibilität** | Windows-fokussiert, erfordert Python. |
| **Wartbarkeit** | Code ist in Logik und GUI getrennt, gut dokumentiert. Stylesheets sind in `.qss`-Dateien ausgelagert, was die Wartung des Designs vereinfacht. |
| **Reproduzierbarkeit** | Das Projekt enthält eine `requirements.txt`-Datei, die eine einfache und konsistente Installation aller Abhängigkeiten ermöglicht. |

---

#### **6. Zukünftige Erweiterungen (Außerhalb des aktuellen Umfangs)**

*   **PDF-Export:** Eine Option, die generierten Word-Dokumente direkt als PDF-Dateien zu speichern.
*   **Unterstützung für weitere Datenquellen:** Einlesen von Daten aus CSV-Dateien oder direkt aus einer Datenbank.
*   **Web-Oberfläche:** Bereitstellung des Tools als Webanwendung für den plattformunabhängigen Zugriff.
*   **Validierungs-Tool:** Eine Funktion, um die Excel-Datei und die Vorlagen vor dem eigentlichen Lauf auf häufige Fehler zu überprüfen. 