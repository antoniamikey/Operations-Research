"""
Abschlussprojekt
Modell-Datei
@author: Antonia Schikora, Cedric Schuster, Jonas Stassiuk, Pascal Moroschan 
"""

import xpress as xp
import useful_functions as use

# Öffne Datei 
workbook, sheet = use.open_sheet("ProjektLastgang.xlsx", 0) 

# Initialisiere Instanz
lastgang = xp.problem()


################################
### PARAMETER, MENGEN        ###
################################

### Skalarparameter ###########

# Anlasskosten
ako1 = use.read_scalar(sheet, "N2") # Anlasskosten für Generator Typ 1
ako2 = use.read_scalar(sheet, "O2") # Anlasskosten für Generator Typ 2 

# Kosten auf Minimallast pro Stunde
minlastko1 = use.read_scalar(sheet, "N3") # Anlasskosten für Generator Typ 1
minlastko2 = use.read_scalar(sheet, "O3") # Anlasskosten für Generator Typ 2 

# Kostenoberhalb der Minimallast pro MW pro Stunde
kouebermin1 = use.read_scalar(sheet, "N4") # Anlasskosten für Generator Typ 1
kouebermin2 = use.read_scalar(sheet, "O4") # Anlasskosten für Generator Typ 2 

# Minimallast in MW
minlast1 = use.read_scalar(sheet, "N5") # Anlasskosten für Generator Typ 1
minlast2 = use.read_scalar(sheet, "O5") # Anlasskosten für Generator Typ 2 

# Maximallast in MW
maxlast1 = use.read_scalar(sheet, "N6") # Anlasskosten für Generator Typ 1
maxlast2 = use.read_scalar(sheet, "O6") # Anlasskosten für Generator Typ 2 

# Anzahl verfügbarer Generatoren
vergen1 = use.read_scalar(sheet, "N7") # Anlasskosten für Generator Typ 1
vergen2 = use.read_scalar(sheet, "O7") # Anlasskosten für Generator Typ 2 

### Mengen #####################

# Tageszeit
ZEIT = use.read_index(sheet, "A3", "vertical", 24)

# Wochentag
TAG = use.read_index(sheet, "B2", "horizontal", 7)

### Vektorparameter ############

# Transportkosten
bedarf = use.read_table(sheet, "B3", TAG, ZEIT)


################################
### VARIABLEN                ###
################################

# Produktionsmenge pro Periode -> AUS VORHERIGEM PROJEKT; MUSS NOCH GEÄNDERT WERDEN
x = lastgang.addVariables(PERIODE, name="x", 
                     lb=0, vartype = xp.continuous)




################################
### ZIELFUNKTION             ###
################################

# Minimierung der Produktionskosten über alle Perioden-> AUS VORHERIGEM PROJEKT; MUSS NOCH GEÄNDERT WERDEN
produktionskosten = xp.Sum(x[p] * pk +
                    l[p] * lk +
                    erh[p] * ek +
                    verm[p] * vk
                    for p in PERIODE)

# Hinzufügen der Zielfunktion zum Modell 
lastgang.setObjective(kosten, sense=xp.minimize)


################################
### NEBENBEDINGUNGEN         ###
################################

######-> AUS VORHERIGEM PROJEKT; MUSS NOCH GEÄNDERT WERDEN

# Anpassung der Produktionsmenge an den Bedarf -> Änderungsmenge für die erste Periode wird nicht betrachtet, da diese immer 0 ist
produktionsmenge = [x[PERIODE[i]] + l[PERIODE[i-1]] >= bedarf[PERIODE[i]] 
                   for i in range(1,I)] 

# Veränderungen in der Produktionsmenge, welche ebenfalls Kosten verursachen 
bilanz = [erh[PERIODE[i]] - verm[PERIODE[i]] == x[PERIODE[i]] - x[PERIODE[i-1]]
             for i in range(1, I)] 

# Berechnung des Lagerbestands 
lager_start = l0 + x[PERIODE[0]] - bedarf[PERIODE[0]]==l[PERIODE[0]]  # Lagerbestand wird für die erste Periode gegeben

lager = [x[PERIODE[i]] + l[PERIODE[i-1]] - bedarf[PERIODE[i]] == l[PERIODE[i]]
         for i in range(1, I)] # Lagerbestand für die weiteren Perioden

# Nebenbedingungen zur Instanz hinzufügen
lastgang.addConstraint()


################################
### INSTANZ LOESEN & AUSGABE ###
################################

# Instanz lösen
lastgang.lpOptimize()

# Lösung anfordern 

# Ausgabe in der Konsole

# Ausgabe im Excel Sheet

# Datei speichern und schließen -> AUS VORHERIGEM PROJEKT; MUSS NOCH GEÄNDERT WERDEN
use.save_sheet(workbook, "U22 Produktionsplan.xlsx")


#########################################
### Dualvariablen & reduzierte Kosten ###
#########################################

