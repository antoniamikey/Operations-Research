"""
Abschlussprojekt
Daten-Datei
@author: Antonia Schikora, Cedric Schuster, Jonas Stassiuk, Pascal Moroschan
"""

import useful_functions as use

# Öffnet das erste Tabellenblatt (Index 0) der Excel-Datei zum Lesen und Schreiben
workbook, sheet = use.open_sheet("ProjektLastgangUpgrade.xlsx", 0) 

################################
### PARAMETER, MENGEN        ###
################################

### Skalarparameter ###########

# Lastgangpuffer, wie viel Spielraum ist nach oben und nach unten gegeben
puffer_oben = use.read_scalar(sheet, "L14") 
puffer_unten = use.read_scalar(sheet, "L15")

### Mengen #####################

# Wochentag
TAG = use.read_index(sheet, "B2", "horizontal", 7)

# Tageszeit
STUNDE = use.read_index(sheet, "A3", "vertical", 24)

# Generatortypen -> hier Zahl ändern, wenn mehr Generatortypen verwendet werden sollen
GEN = use.read_index(sheet, "K3", "vertical", 3)

### Vektorparameter ############

# Anlasskosten (EUR pro Anlassung)
anlassko = use.read_table(sheet, "L3", GEN)

# Fixkosten auf Minimallast (EUR pro Stunde je Generator)
minlastko = use.read_table(sheet, "O3", GEN)

# Variable Kosten oberhalb Minimallast (EUR pro MW und Stunde)
kouebermin = use.read_table(sheet, "R3", GEN)

# Minimallast (MW)
minlast = use.read_table(sheet, "U3", GEN)

# Maximallast (MW)
maxlast = use.read_table(sheet, "X3", GEN)

# Anzahl verfügbarer Generatoren
verfgen = use.read_table(sheet, "AA3", GEN)

# Lastgang-Bedarf (MW)
bedarf = use.read_table(sheet, "B3", STUNDE, TAG)