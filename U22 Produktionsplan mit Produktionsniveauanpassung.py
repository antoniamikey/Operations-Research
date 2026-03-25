"""
U22 Produktionsplan mit Produktionsniveauanpassung 
Modell-Datei
@author: Antonia Schikora
"""

import xpress as xp
import useful_functions as use

# Öffne Datei 
workbook, sheet = use.open_sheet("U22 Produktionsplan.xlsx", 0) 

# Initialisiere Instanz
U22 = xp.problem()


################################
### PARAMETER, MENGEN        ###
################################

### Skalarparameter ###########

# Kosten die durch die Veränderunge der Produktionsmenge entstehen 
ek = use.read_scalar(sheet, "C3") # Erhöhungskosten
vk = use.read_scalar(sheet, "D3") # Verminderungskosten 


# Produktionskosten 
pk = use.read_scalar(sheet, "A3")

# Lagerkosten 
lk = use.read_scalar(sheet, "B3")

# Lagerbestand zum Zeitpunkt t=0
l0 = use.read_scalar(sheet, "E8")

# Anzahl Perioden in denen produziert wird
I = use.read_scalar(sheet, "F3")

### Mengen #####################

# Perioden in denen produziert wird
PERIODE = use.read_index(sheet, "A9", "vertical", I)

### Vektorparameter ############

# Bedarf pro Periode
bedarf = use.read_table(sheet, "B9", PERIODE)


################################
### VARIABLEN                ###
################################

# Produktionsmenge pro Periode
x = U22.addVariables(PERIODE, name="x", 
                     lb=0, vartype = xp.continuous)

# Lagerbestand zum Ende jeder Periode
l = U22.addVariables(PERIODE, name="l",
                     lb=0, vartype = xp.continuous)

# Veränderungen in der Produktionsmenge
erh = U22.addVariables(PERIODE, name="erh",
                     lb=0, vartype = xp.continuous)

verm = U22.addVariables(PERIODE, name="verm",
                     lb=0, vartype = xp.continuous)


################################
### ZIELFUNKTION             ###
################################

# Minimierung der Produktionskosten über alle Perioden
produktionskosten = xp.Sum(x[p] * pk +
                    l[p] * lk +
                    erh[p] * ek +
                    verm[p] * vk
                    for p in PERIODE)

# Hinzufügen der Zielfunktion zum Modell
U22.setObjective(produktionskosten, sense=xp.minimize)


################################
### NEBENBEDINGUNGEN         ###
################################

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
U22.addConstraint(lager_start, lager, bilanz, produktionsmenge)


################################
### INSTANZ LOESEN & AUSGABE ###
################################

# Instanz lösen
U22.lpOptimize()

# Lösung anfordern 
produktionsmengex = U22.getSolution(x)
lagerbestand = U22.getSolution(l)
verminderungen = U22.getSolution(verm)
erhöhungen = U22.getSolution(erh)

# Berechnung der Kosten für die Veränderungen des Produktionsniveaus
kosten_produktionsniveau = sum(verminderungen[i] * vk + erhöhungen[i] * ek for i in PERIODE) 

# Berechnung der 
gesamtkosten = sum(produktionsmengex[i] * 10 + verminderungen[i] * vk + erhöhungen[i] * ek + lagerbestand[i] * lk for i in PERIODE)

# Ausgabe in der Konsole
print("Gesamtkosten: ", gesamtkosten)
print("Produktionsmenge pro Periode: ", produktionsmengex)
print("Lagerbestand pro Periode: ", lagerbestand)
print("Verminderungen pro Periode: ", verminderungen)
print("Erhöhungen pro Periode: ", erhöhungen)
print("Gesamtkosten für den Wechsel des Produktionsniveaus: ", kosten_produktionsniveau)

# Ausgabe im Excel Sheet
use.write_tbody(sheet, "D8", lagerbestand, PERIODE)
use.write_tbody(sheet, "G8", produktionsmengex, PERIODE)
use.write_scalar(sheet, "H2", kosten_produktionsniveau, "Gesamtkosten für die Änderung des Produktionsniveaus")

# Datei speichern und schließen 
use.save_sheet(workbook, "U22 Produktionsplan.xlsx")

dual_Lager_start = U22.getDuals(lager_start)
dual_lager = U22.getDuals(lager)

# Hier kommentieren
alle_duals = {}
alle_duals[PERIODE[0]] = dual_Lager_start  
for i in range(1, I):
    alle_duals[PERIODE[i]] = dual_lager[i-1] #hier negativ, da bei uns der
print("Dualwerte für die Erhöhung des Bedarfs: ", alle_duals)

min_periode = min(alle_duals, key=lambda p: alle_duals[p])
print("Günstigste Periode:", min_periode)
print("Geringste Kostenerhöhung:", alle_duals[min_periode])

# (b) Gib eine Interpretation der reduzierten Kosten derjenigen Variable an, die den Lagerbestand am Anfang von Periode 4 modelliert.
#rc = U22.getRedCosts(l)
#print("Reduzierte Kosten l[Periode 3]:", rc[PERIODE[2]])