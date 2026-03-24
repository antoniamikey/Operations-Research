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
print(pk)

# Lagerkosten 
lk = use.read_scalar(sheet, "B3")

# Lagerbestand zum Zeitpunkt t=0
l0 = use.read_scalar(sheet, "K3")

# Anzahl Perioden
I = use.read_scalar(sheet, "G3")
print(I)

### Mengen #####################
PERIODE = use.read_index(sheet, "I4", "vertical", I)
print(PERIODE[0])

#PERIODEPRDUKTION = use.read_index(sheet, "J10", "vertical", "I")

### Vektorparameter ############

# Bedarf pro Periode
bedarf = use.read_table(sheet, "J4", PERIODE)
print(bedarf)

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

produktionskosten = xp.Sum(x[p] * pk +
                    l[p] * lk +
                    erh[p] * ek +
                    verm[p] * vk
                    for p in PERIODE)

U22.setObjective(produktionskosten, sense=xp.minimize)

################################
### NEBENBEDINGUNGEN         ###
################################
produktionsmenge_start = x[PERIODE[0]] + l0 >= bedarf[PERIODE[0]]
produktionsmenge = [x[PERIODE[i]] + l[PERIODE[i-1]] >= bedarf[PERIODE[i]] for i in range(1,I)]

# Produktionsbilanz 
# erhb = [erh[PERIODE[i]] == x[PERIODE[i]] - x[PERIODE[i-1]]
#        for i in range(1,I)
#        if x[PERIODE[i]] - x[PERIODE[i-1]] >=0]

# vermb = [verm[PERIODE[i]] == -x[PERIODE[i]] + x[PERIODE[i-1]]
#        for i in range(1,I)
#        if x[PERIODE[i]] - x[PERIODE[i-1]] <=0]

bilanz = [erh[PERIODE[i]] - verm[PERIODE[i]]
             == x[PERIODE[i]] - x[PERIODE[i-1]]
             for i in range(1, I)] 


# Lagerbestand
lager_start = l[PERIODE[0]] == l0 + x[PERIODE[0]] - bedarf[PERIODE[0]]
lager = [l[PERIODE[i]] == x[PERIODE[i]] + l[PERIODE[i-1]] - bedarf[PERIODE[i]]
         for i in range(1, I)]

U22.addConstraint(produktionsmenge_start, produktionsmenge, lager_start, lager)

################################
### INSTANZ LOESEN & AUSGABE ###
################################

# Instanz loesen
U22.lpOptimize()
U22.getSolution(x)

print(U22.getSolution(x))
print(U22.getSolution(verm))
print(U22.getSolution(erh))