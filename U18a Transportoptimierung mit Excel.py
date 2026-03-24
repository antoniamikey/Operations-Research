"""
U18a Transportoptimierung mit Excel
Modell-Datei
@author: Kai Helge Becker
"""

import xpress as xp
import useful_functions as use

# Öffne Datei 
workbook, sheet = use.open_sheet("U18aTransportoptimierung.xlsx", 0) 

# Initialisiere Instanz
LP1 = xp.problem()

################################
### PARAMETER, MENGEN        ###
################################

### Skalarparameter ##########

# Anzahl Lieferknoten (Quellen)
L = use.read_scalar(sheet, "A2")

# Anzahl andere Knoten (Senken und Umladeknoten)
A = use.read_scalar(sheet, "B2")  

### Mengen #####################

# Lieferknoten
LKNOTEN = use.read_index(sheet, "J2", "vertical", L)
print(LKNOTEN)

# Weitere Knoten
WKNOTEN = use.read_index(sheet, "J6", "vertical", A) 

# Alle Knoten
KNOTEN = use.read_index(sheet, "B5", "horizontal", L+A)

### Vektorparameter ############

# Mindestbereitstellung
minbereit = use.read_table(sheet, "K2", LKNOTEN)

# Bedarf
bedarf = use.read_table(sheet, "K6", WKNOTEN)

# Transportkosten
kosten = use.read_table(sheet, "B6", KNOTEN, KNOTEN)

################################
### VARIABLEN                ###
################################

# Transportmenge
FlussV = LP1.addVariables(KNOTEN, KNOTEN, name="FlussV", 
                          lb=0, vartype=xp.continuous)

# Lieferung
LieferV = LP1.addVariables(LKNOTEN, name="LieferV", 
                           lb=0, vartype=xp.continuous)


################################
### ZIELFUNKTION             ###
################################

# minimiere Summe der Transportkosten auf allen Flüssen
transportkosten = xp.Sum(kosten[i, j] * FlussV[i, j] 
                         for i in KNOTEN for j in KNOTEN)
    
# Füge Zielfunktion zum Modell hinzu    
LP1.setObjective(transportkosten, sense=xp.minimize)

################################
### NEBENBEDINGUNGEN         ###
################################

# Knotenbilanzgleichung Lieferknoten
knotenBilanzL = [xp.Sum(FlussV[l, j] for j in KNOTEN)
               == xp.Sum(FlussV[i, l] for i in KNOTEN)
               +  LieferV[l]
               for l in LKNOTEN]

# Knotenbilanzgleichung andere Knoten
knotenBilanzA = [xp.Sum(FlussV[w, j] for j in KNOTEN)
               +  bedarf[w]  
               == xp.Sum(FlussV[i, w] for i in KNOTEN)
               for w in WKNOTEN]

# Mindestbereitstellung
minLieferung = [LieferV[l] >= minbereit[l] for l in LKNOTEN]

# Nebenbedingungen zur Instanz hinzufügen
LP1.addConstraint(knotenBilanzL, knotenBilanzA, minLieferung)


################################
### INSTANZ LOESEN & AUSGABE ###
################################

# Instanz loesen
LP1.lpOptimize('n')

# Loesung anfordern
transport_menge = LP1.getSolution(FlussV)
liefer_menge = LP1.getSolution(LieferV)
transportkosten = LP1.attributes.objval

# Ausgabe
use.write_thead(sheet, "A14", "Transportmenge", "Knoten", "Knoten")
use.write_tbody(sheet, "A14", transport_menge, KNOTEN, KNOTEN)
use.write_thead(sheet, "J15", "Liefermenge", "Lieferknoten")
use.write_tbody(sheet, "J15", liefer_menge, LKNOTEN)
use.write_scalar(sheet, "M15", transportkosten, "Gesamtkosten")

# Datei speichern und schließen
use.save_sheet(workbook, "U18aTransportoptimierung.xlsx")


