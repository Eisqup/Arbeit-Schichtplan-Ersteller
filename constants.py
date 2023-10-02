SCHICHT_MODELS = {1: "Dauerschicht", 2: "Schicht Zyklus", 3: "individuelle Schicht"}

SCHICHT_RHYTHMUS = {1: "Früh", 2: "Spät", 3: "Nacht"}

# FILE_NAME = "Mitarbeiter_2124_Fertigung.json"
# FILE_NAME = "Mitarbeiter_212.4_Montage.json"
FILE_NAME = "Mitarbeiter.json"

EMPLOYEE_KEY = ["name", "schicht_model", "schicht_rhythmus", "bereiche", "urlaub_kw", "urlaub_tage", "link"]


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
    + "=>Überhang bekommt max value + 2 extra Reihen\n"
    + '=>Benutze "-b" als name um einen Bereichstrennung zu erzeugen. Hier wird ein MA Zähler hinzugefügt'
)
