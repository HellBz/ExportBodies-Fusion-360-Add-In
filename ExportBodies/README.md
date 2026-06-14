# 🔄 ExportBodies – Fusion 360 Add-In

**ExportBodies** ist ein Fusion 360 Add-In, mit dem du alle BRep-Bodies aus dem Root-Component (und optional aus Unterkomponenten) schnell und einfach in mehrere Dateiformate gleichzeitig exportieren kannst – strukturiert und wiederholbar.

---

## 📦 Features

- ✅ Export **aller sichtbaren Bodies** im Root-Component
- ✅ Optional: Export von Bodies aus **allen Unterkomponenten**
- ✅ Auswahl mehrerer Formate gleichzeitig: **STL, 3MF, STEP, IGES, SAT, SMT**
- ✅ **Versionierte Dateinamen** basierend auf dem Dokumentnamen (z. B. `Body_v2.stl`)
- ✅ Merkt sich die **zuletzt gewählten Exportformate**
- ✅ Automatischer Vorschlag des Exportordners – mit Möglichkeit zur manuellen Auswahl
- ✅ Erstellt automatisch einen **projektbezogenen Unterordner**
- ✅ **Mesh-Verfeinerung** für STL & 3MF wählbar: Low / Medium / High / Custom
- ✅ Im **Custom-Modus**: manuelle Eingabe von Surface Deviation, Normal Deviation, Max Edge Length & Aspect Ratio
- ✅ Unsichtbare Bodies werden für den Export **temporär eingeblendet** und danach wieder versteckt
- ✅ Öffnet den Exportordner nach dem Export optional automatisch
- ✅ Plattformunterstützung: **Windows & macOS**

---

## 🛠️ Installation

1. Repository klonen oder als ZIP herunterladen.

2. Den gesamten `ExportBodies`-Ordner in das Fusion 360 Add-Ins-Verzeichnis verschieben:

   - **Windows**  
     `C:\Users\<DeinBenutzername>\AppData\Roaming\Autodesk\Autodesk Fusion 360\API\AddIns\`

   - **macOS**  
     `~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/AddIns/`

3. Fusion 360 starten.

4. Den Dialog **Scripts and Add-Ins** öffnen:  
   `Tools` → `Add-Ins` → `Scripts and Add-Ins`

5. Auf den Tab **Add-Ins** wechseln und **ExportBodies** in der Liste finden.

6. Auf **Run** klicken – optional **Run on Startup** aktivieren.

---

## 🧰 Verwendung

1. Im **Solid**-Tab auf den **Export Bodies**-Button im Panel klicken.

2. Im Dialog die gewünschten **Exportformate** per Checkbox auswählen.

3. Optional:
   - **Only export visible bodies** – nur sichtbare Bodies exportieren
   - **Export bodies from components** – Bodies aus Unterkomponenten einschließen
   - **Mesh Refinement (STL & 3MF)** – Qualitätsstufe der Mesh-Vernetzung wählen: `Low`, `Medium`, `High` oder `Custom`
     - Bei `Custom` erscheinen zusätzliche Felder: `Surface Deviation`, `Normal Deviation`, `Max Edge Length`, `Aspect Ratio`

4. Den **Exportordner** bestätigen oder einen eigenen Ordner wählen.

5. Auf **OK** klicken – alle Bodies werden automatisch exportiert.

6. Nach dem Export: optional den **Ordner direkt öffnen**.

---

## 📁 Ausgabe-Ordnerstruktur

Exporte werden in einem `EXPORTED`-Ordner gespeichert, der sich **zwei Ebenen über** dem Add-In-Verzeichnis befindet.  
Pro Fusion-Dokument wird automatisch ein eigener Unterordner erstellt:

```
EXPORTED/
└── MeinProjekt/
    ├── Body1_v3.stl
    ├── Body1_v3.3mf
    ├── Body2_v3.stl
    └── ...
```

Bei Bodies aus Unterkomponenten wird der Dateiname aus **Komponentenname + Bodyname** zusammengesetzt:

```
EXPORTED/
└── MeinProjekt/
    ├── Root_Body1_v1.stl
    ├── Gehaeuse_Body1_v1.stl
    └── Deckel_Body2_v1.stl
```

---

## ⚙️ Unterstützte Formate

| Format | Beschreibung                        |
|--------|-------------------------------------|
| STL    | Mesh (hohe Auflösung)               |
| 3MF    | Mesh, Farben & Metadaten (hohe Auflösung) |
| STEP   | Parametrisches CAD-Format           |
| IGES   | Klassisches CAD-Austauschformat     |
| SAT    | ACIS-Solid-Format                   |
| SMT    | Autodesk Shape Manager Format       |

---

## 🗂️ Projektstruktur

```
ExportBodies/
├── commands/
│   └── Export/
│       ├── entry.py           # Hauptlogik: Dialog & Export
│       ├── resources/         # Icons für das Add-In
│       └── last_export_config.txt  # Gespeicherte Formatauswahl
├── lib/
│   └── fusionAddInUtils/      # Hilfsfunktionen (Event-Handling, Logging)
├── config.py                  # Konfiguration (Panel-Name, IDs, etc.)
├── ExportBodies.py            # Einstiegspunkt des Add-Ins
└── ExportBodies.manifest      # Add-In Manifest

```

---

## 📝 Lizenz

MIT License – frei verwendbar und anpassbar.
