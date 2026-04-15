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

# Lastgangpuffer, wie viel Spielraum ist nach oben und nach unten gegeben
puffer_oben = use.read_scalar(sheet, "J2") 
puffer_unten = use.read_scalar(sheet, "K2")

### Mengen #####################

# Wochentag
TAG = use.read_index(sheet, "B2", "horizontal", 7)

# Tageszeit
STUNDE = use.read_index(sheet, "A3", "vertical", 24)

# Generatortypen -> hier Zahl ändern, wenn mehr Generatortypen verwendet werden sollen
GEN = use.read_index(sheet, "N1", "horizontal", 2)

### Vektorparameter ############

# Anlasskosten (EUR pro Anlassung)
anlassko = use.read_table(sheet, "N2", GEN)

# Fixkosten auf Minimallast (EUR pro Stunde je Generator)
minlastko = use.read_table(sheet, "N3", GEN)

# Variable Kosten oberhalb Minimallast (EUR pro MW und Stunde)
kouebermin = use.read_table(sheet, "N4", GEN)

# Minimallast (MW)
minlast = use.read_table(sheet, "N5", GEN)

# Maximallast (MW)
maxlast = use.read_table(sheet, "N6", GEN)

# Anzahl verfügbarer Generatoren
verfgen = use.read_table(sheet, "N7", GEN)

# Lastgang-Bedarf (MW)
bedarf = use.read_table(sheet, "B3", STUNDE, TAG)


################################
### VARIABLEN                ###
################################

# Gesamtproduktion aller Generatoren von Typ g in Stunde h an Tag t (in MW)
ProdV = lastgang.addVariables(GEN, STUNDE, TAG, name="ProdV", 
                     lb=0, vartype = xp.continuous)

# Anzahl laufender Generatoren von Typ g in Stunde h an Tag t (ganzzahlig)
LaufGV = lastgang.addVariables(GEN, STUNDE, TAG, name="LaufGV", 
                     lb=0, vartype = xp.integer)

# Anzahl neu gestarteter Generatoren von Typ g in Stunde h an Tag t (ganzzahlig)
NeuGV = lastgang.addVariables(GEN, STUNDE, TAG, name="NeuGV", 
                     lb=0, vartype = xp.integer)

# Netzwerkstabilität (NetzV[STUNDE, TAG] = 1 wenn mehr als 3 Generatoren vom Typ 1 verwendet werden)
NetzV = lastgang.addVariables(STUNDE, TAG, name="NetzV", 
                          lb=0, vartype=xp.binary)


################################
### ZIELFUNKTION             ###
################################

# Minimierung der Gesamtkosten
gesamtkosten = (
    
    # Anlasskosten
    xp.Sum(anlassko[i] * NeuGV[i, h, t] for i in GEN for h in STUNDE for t in TAG)

    # + Kosten auf Minimallast
    + xp.Sum(minlastko[i] * LaufGV[i, h, t] for i in GEN for h in STUNDE for t in TAG)

    # + Kosten über Minimallast 
    + xp.Sum(kouebermin[i] * (ProdV[i, h, t] - minlast[i] * LaufGV[i, h, t]) for i in GEN for h in STUNDE for t in TAG)
)

# Hinzufügen der Zielfunktion zum Modell 
lastgang.setObjective(gesamtkosten, sense=xp.minimize)


################################
### NEBENBEDINGUNGEN         ###
################################

# Lastgang muss mindestens dem Bedarf entsprechen
bedarfsdeckung = [xp.Sum(ProdV[i, h, t] for i in GEN) >= bedarf[h, t] 
                   for h in STUNDE for t in TAG] 

# Lastgang muss die Puffer einhalten
bilanz = [erh[PERIODE[i]] - verm[PERIODE[i]] == x[PERIODE[i]] - x[PERIODE[i-1]]
             for i in range(1, I)] 

# es muss Netzwerkstabilität gewährleistet werden -> es ist nicht möglich, mehr als 3 Generatoren vom Typ 2 zu verwenden, wenn zu demselben Zeitpunkt mehr als 3 Generatoren vom Typ 1 verwendet werden.
lager_start = l0 + x[PERIODE[0]] - bedarf[PERIODE[0]]==l[PERIODE[0]]  

# Lastgang zum Ende der Woche muss dem Lastgang zu Beginn der nächsten Periode entsprechen 
lager = [x[PERIODE[i]] + l[PERIODE[i-1]] - bedarf[PERIODE[i]] == l[PERIODE[i]]
         for i in range(1, I)] # Lagerbestand für die weiteren Perioden

# Kostensachen

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

