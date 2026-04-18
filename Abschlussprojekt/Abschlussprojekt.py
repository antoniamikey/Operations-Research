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
puffer_oben = use.read_scalar(sheet, "L14") 
puffer_unten = use.read_scalar(sheet, "L15")

### Mengen #####################

# Wochentag
TAG = use.read_index(sheet, "B2", "horizontal", 7)

# Tageszeit
STUNDE = use.read_index(sheet, "A3", "vertical", 24)

# Generatortypen -> hier Zahl ändern, wenn mehr Generatortypen verwendet werden sollen
GEN = use.read_index(sheet, "K3", "vertical", 2)

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
gesamtkosten = xp.Sum(anlassko[g] * NeuGV[g, s, t] #Anlasskosten
    + minlastko[g] * LaufGV[g, s, t] # Kosten auf Minimallast
    + kouebermin[g] * (ProdV[g, s, t] - minlast[g] * LaufGV[g, s, t]) #Kosten über Minimallast 
    for g in GEN for s in STUNDE for t in TAG)

# Hinzufügen der Zielfunktion zum Modell 
lastgang.setObjective(gesamtkosten, sense=xp.minimize)


################################
### NEBENBEDINGUNGEN         ###
################################

# Lastgang muss mindestens dem Bedarf entsprechen
bedarfsdeckung = [xp.Sum(ProdV[i, h, t] for i in GEN) >= bedarf[h, t] 
                   for h in STUNDE for t in TAG] 

# Minimalproduktion (bestimmt durch Minimallast)
produktionsminimum = [ProdV[i, h, t] >= minlast[i] * LaufGV[i, h, t]
                      for i in GEN for h in STUNDE for t in TAG]

# Maximalproduktion (bestimmt durch Maximallast)
produktionsmaximum = [ProdV[i, h, t] <= maxlast[i] * LaufGV[i, h, t]
                      for i in GEN for h in STUNDE for t in TAG]


# Lastgang muss die Puffer einhalten
# oberer Puffer -> Generatoren müssen den Bedarf plus den Erhöhungspuffer abdecken können ohne einen neuen Generator anlassen zu müssen
obererPuffer = [xp.Sum(maxlast[i] * LaufGV[i, h, t] for i in GEN) >= (1 + puffer_oben) * bedarf[h, t]
                for h in STUNDE for t in TAG] 

# unterer Puffer -> Generatoren müssen den Bedarf abzüglich des Verminderungspuffers abdecken können ohne einen Generator unter seine Minimallast zu bringen
untererPuffer = [xp.Sum(minlast[i] * LaufGV[i, h, t] for i in GEN) <= (1 - puffer_unten) * bedarf[h, t]
                 for h in STUNDE for t in TAG]


# es muss Netzwerkstabilität gewährleistet werden -> es ist nicht möglich, mehr als 3 Generatoren vom Typ 2 zu verwenden, wenn zu demselben Zeitpunkt mehr als 3 Generatoren vom Typ 1 verwendet werden.

M1 = verfgen[GEN[0]]-3
M2 = verfgen[GEN[1]]-3

netzwerkstabilität1 = [LaufGV[GEN[0], h, t] <= 3 + M1 * NetzV[h, t]
                       for h in STUNDE for t in TAG] # y wird auf 1 gesetzt, wenn von Typ 1 mehr als 3 verwendet werden

netzwerkstabilität2 = [LaufGV[GEN[1], h, t] <= 3 + M2 * (1 - NetzV[h, t])
                      for h in STUNDE for t in TAG]  


# Lastgang zum Ende der Woche muss dem Lastgang zu Beginn der nächsten Periode entsprechen / entstehende Anlasskosten
neugestarteteGeneratoren_start = [NeuGV[g, STUNDE[0], TAG[0]] >= LaufGV[g, STUNDE[0], TAG[0]]
                                   for g in GEN]  # Startzeitpunkt: Vorgänger = 0
neugestarteteGeneratoren_stunde = [NeuGV[g, STUNDE[s], t] >= LaufGV[g, STUNDE[s], t] - LaufGV[g, STUNDE[s-1], t]
                                   for g in GEN for s in range(1, len(STUNDE)) for t in TAG ]
neugestarteteGeneratoren_tag = [NeuGV[g, STUNDE[0], TAG[t]] >= LaufGV[g, STUNDE[0], TAG[t]] - LaufGV[g, STUNDE[0], TAG[t-1]]
                                for g in GEN for t in range(1, len(TAG))]


# # Berechnung des Lagerbestands 
# lager_start = l0 + x[PERIODE[0]] - bedarf[PERIODE[0]]==l[PERIODE[0]]  # Lagerbestand wird für die erste Periode gegeben

# Nebenbedingungen zur Instanz hinzufügen
lastgang.addConstraint(bedarfsdeckung, produktionsminimum, produktionsmaximum, obererPuffer, untererPuffer, netzwerkstabilität1, netzwerkstabilität2, neugestarteteGeneratoren_start, neugestarteteGeneratoren_stunde, neugestarteteGeneratoren_tag)


################################
### INSTANZ LOESEN & AUSGABE ###
################################

# Instanz lösen
lastgang.mipOptimize()


# Lösung anfordern 

# Ausgabe in der Konsole

# Ausgabe im Excel Sheet

# Datei speichern und schließen -> AUS VORHERIGEM PROJEKT; MUSS NOCH GEÄNDERT WERDEN
# use.save_sheet(workbook, "U22 Produktionsplan.xlsx")


#########################################
### Dualvariablen & reduzierte Kosten ###
#########################################

