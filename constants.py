SCHICHT_MODELS = {1: "Dauerschicht", 2: "Schicht Zyklus", 3: "individuelle Schicht"}
SCHICHT_MODELS_BESCHREIBUNG = ["Nur eine Schicht das ganze Jahr.", "Wiederholender Rhythmus.", "Bevorzugte Schichten nach Priorität sortiert."]

SCHICHT_RHYTHMUS = {1: "Früh", 2: "Spät", 3: "Nacht"}

# FILE_NAME = "Mitarbeiter_2124_Fertigung.json"
# FILE_NAME = "Mitarbeiter_212.4_Montage.json"
FILE_NAME = "Mitarbeiter.json"

EMPLOYEE_KEY = ["name", "schicht_model", "schicht_rhythmus", "bereiche", "urlaub_kw", "urlaub_tage", "link"]

EXCEL_SHEET_PW = 5678


# Variable
PROGRAM_RUNS = 100


SETTINGS_VARIABLES = {
    "WEEK_RANGE": 6,  # Excel view with how much days the week is showed
    "MAX_RUNS_IN_YEAR_FOR_MODEL_3_EMP_WITHOUT_SHIFT": 3,  # Max run in year for emp without shift in model 3
    "FACTOR_FOR_MODEL_3_EMP_WITHOUT_SHIFT": 5,  # Factor * shift_count / program thinks he did more of the not choose shift
    "USE_MODEL_2_EMP_WITH_ALL_3_SHIFTS_IF_AREAS_IN_OTHER_SHIFTS_UNDERSTAFFED": False,
    "USE_ALL_MODEL_2_EMP_IF_AREAS_IN_OTHER_SHIFTS_UNDERSTAFFED": False,
    "IGNORE_MODEL_2_START_WEEK_LOGIC": False,
}

SETTING_DESCRIPTION = {
    "WEEK_RANGE": "Bestimmt, für wie viele Tage in der Woche die Excel-Datei erstellt wird.",
    "MAX_RUNS_IN_YEAR_FOR_MODEL_3_EMP_WITHOUT_SHIFT": "Legt fest, wie viele Schichten ein Mitarbeiter im Schichtmodell 3 ohne die benötigte Schicht im Laufe des Jahres übernehmen kann, wenn niemand anderes im Modell 3 verfügbar ist.",
    "FACTOR_FOR_MODEL_3_EMP_WITHOUT_SHIFT": "Falls ein Mitarbeiter im Schichtmodell 3 gelegentlich Schichten ohne die gewünschte Schicht übernimmt, berechnet das Programm bei zukünftigen Schichtplanungen einen Faktor * die Anzahl der ungewollten Schichten.",
    "USE_MODEL_2_EMP_WITH_ALL_3_SHIFTS_IF_AREAS_IN_OTHER_SHIFTS_UNDERSTAFFED": "Wenn kein Mitarbeiter im Modell 3 verfügbar ist, um eine unterbesetzte Schicht zu übernehmen, wird, wenn möglich, ein Mitarbeiter im Modell 2 mit allen Schichten im Rhythmus eingesetzt.",
    "USE_ALL_MODEL_2_EMP_IF_AREAS_IN_OTHER_SHIFTS_UNDERSTAFFED": "Das Gleiche wie 'USE_MODEL_2_EMP_WITH_ALL_3_SHIFTS_IF_AREAS_IN_OTHER_SHIFTS_UNDERSTAFFED', jedoch können alle Mitarbeiter im Modell 2 verwendet werden.",
    "IGNORE_MODEL_2_START_WEEK_LOGIC": "Einschalten, wenn die Schichten für Model-2-Mitarbeiter in KW 1 genau so eingetragen werden sollen, wie sie bereits erfasst sind in der Datenbank.",
}

AREAS_DESCRIPTION = (
    "=> Um leere Reihen bei der Excel Tabelle hinzuzufügen, erstelle einen Bereich mit min=0. Max gibt wieder wie viele Zeilen für den Bereich hinzugefügt werden.\n"
    + "=> Überhang bekommt max value + 2 extra Reihen\n"
    + '=> Benutze "-b" als name um einen Bereichstrennung zu erzeugen(min:0,max:0). Hier wird ein MA Zähler hinzugefügt um Bereiche zu trennen. Daten können später hinzugefügt werden.\n   Der Schichtplan benutzt nur bereich welche auch die MA als Fähigkeit hat, nur die Excel Datei wird dadurch angepasst.\n   Ohne -b wird zum Ende in der Excel Datei automatisch ein Zähler hinzugefügt.'
)

PROGRAMM_INFOS = """
Layout Information:
- Links können Mitarbeiter erstellt und vorhandene Mitarbeiter bearbeitet werden.
- Rechts sind die Informationen zu Mitarbeitern, die bereits in der Datenbank vorhanden sind.
- Der "Mitarbeiter Linken" ist für Mitarbeiter gedacht, die beispielsweise eine Fahrgemeinschaft haben (Voraussetzung: gleicher Rhythmus).

Bereiche bearbeiten Button:
- Die Sortierung der Bereiche entspricht der Sortierung in der Excel-Datei.
- Das "-b" wird genutzt, um eine Trennung in der Excel-Datei zu erstellen, um eventuell Überbereiche aufzuteilen.
- Bereiche mit einem minimalen Wert von 0 werden im Algorithmus nicht beachtet, aber die Excel-Datei erstellt für den Bereich zusätzliche Zeilen entsprechend des maximalen Werts.
- Die "Prio" bestimmt die Wichtigkeit des Bereichs. Zum Beispiel, wenn zwei Bereiche jeweils einen Mitarbeiter benötigen und nur ein Mitarbeiter verfügbar ist, welcher Bereich ist wichtiger? 1 ist am besten.

Erstelle Schichtplan Button:
- Schichtplan nach Wünschen erstellen.
- Einstellungen sind hier verfügbar und mit einer Beschreibung versehen.
- Je häufiger der Schichtplan läuft, desto höher ist die Wahrscheinlichkeit, einen besseren Schichtplan zu erhalten.
- Excel muss installiert sein, und Makros für Excel müssen erlaubt sein (Excel Trust Center).
- Eine Internetverbindung wird benötigt, um Schulferien und Feiertage in die Excel-Datei aufzunehmen.
- Ein Schichtplan kann auch erstellt werden ohne Mitarbeiterinfos. Der Schichtplan wird nach den Bereichsvorgaben erstellt und die Formeln/Statistik angepasst. Wenn Mitarbeiter später hinzugefügt werden, ändern sich die Statistiken entsprechend interaktiv.

Algorithmus:
- Bereiche, die keinen Mitarbeiter haben, werden ignoriert.
- "Dauerschicht" sind Mitarbeiter, die nur eine Schicht haben.
- "Schichtzyklus" sind Mitarbeiter, die einen wiederholenden Rhythmus haben, wie z.B. Früh, Spät, Nacht in einer 3-Schicht-Rotation. Hier kann auch bestimmt werden, wann der Rhythmus starten soll.
- "Individuelle Mitarbeiter" bevorzugen bestimmte Schichten, sind jedoch flexibel einsetzbar (z.B. als Springer). Hier wird darauf geachtet, dass ein fairer Ausgleich für alle Mitarbeiter mit derselben Lieblingsschicht gewährleistet wird.

- 'Dauerschicht'-Mitarbeiter werden zu 100% in ihrer Schicht eingesetzt.
- 'Schichtzyklus'-Mitarbeiter werden in ihrer Schicht entsprechend der Kalenderwoche (KW) eingesetzt. In den Einstellungen kann festgelegt werden, dass sie verschoben werden, um Unterbesetzte Schichten zu besetzen, wenn sie sich in einer überbesetzten Schicht befinden und keine individuellen Mitarbeiter verfügbar sind.
- 'Individuelle Mitarbeiter' werden nach Bedarf im Schichtplan eingesetzt, unter Berücksichtigung ihrer priorisierten Schicht.

Excel Datei:
- Die Datei hat eine Seite für jede Kalenderwoche.
- Statistiken am Ende der Datei sind interaktiv und beinhalten ein Log-System, das zeigt, wer, wann, wo und was geändert hat.
- Beim ersten Speichern wird eine Abfrage gestellt, ob eine Kopie für alle Mitarbeiter erstellt werden soll. Dabei kann ein Pfad für die Speicherung festgelegt werden.
- In diesem Modus ist die Datei schreibgeschützt. Wenn die Hauptdatei gespeichert wird, wird die Kopie automatisch aktualisiert.
- Wenn dies nicht gewünscht ist, kann die Frage beim ersten Speichern verneint werden.
- Wenn die Funktion später geändert werden sollen, muss die Zelle A:1 auf der Seite mit den Bereichsinformationen geleert werden. Diese Zelle ist versteckt und geschützt.
"""
