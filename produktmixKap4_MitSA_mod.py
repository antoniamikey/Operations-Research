"""
Produkt-Mix-Problem mit Lagerhaltung
Modell-Datei
@author: Kai Helge Becker
"""

import xpress as xp
import useful_functions as use

# Verknuepfe mit Datenfile
import produktmixKap4_dat as dat

# Initialisiere Instanz
LP1 = xp.problem()

################################
### PARAMETER, MENGEN        ###
################################

### Skalarparameter ##########

# Anzahl Zeitperioden
T = dat.T  
# Anzahl Produkte
P = dat.P 
# Anzahl Maschinen
M = dat.M


### Mengen #####################

# Zeitperioden
ZEIT  = dat.ZEIT 
ZEIT2 = dat.ZEIT2
# Produktionsstatus
STAT  = dat.STAT
# Produkte
PROD  = dat.PROD
# Maschinen
MASC  = dat.MASC

### Vektorparameter ############

# Lagerkapazitaet
lkap = dat.lkap     # PROD
# Lagerkosten
lkost = dat.lkost   # PROD
# Verkaufspreis
preis = dat.preis   # ZEIT, PROD
# Reduzierter Verkaufspreis
redp = dat.redp     # PROD
# Nachfrage
nach = dat.nach     # ZEIT
# Maschinenkapazität 
mkap = dat.mkap     # ZEIT, STAT, MASC
# Kapazitätsverbrauch
kapver = dat.kapver # ZEIT, STAT, PROD, MASC
# Produktionskosten
pkost = dat.pkost   # ZEIT, STAT, PROD, MASC
  

################################
### VARIABLEN                ###
################################

# Produktionsmenge
ProdV = LP1.addVariables(ZEIT, STAT, PROD, MASC, 
                         name="ProdV", lb=0, vartype=xp.continuous)
# Lagerbestand
LagerV = LP1.addVariables(ZEIT2, PROD, 
                          name="LagerV", lb=0, vartype=xp.continuous)
# Verkaufsmenge
VerkV = LP1.addVariables(ZEIT, PROD, 
                         name="VerkV", lb=0, vartype=xp.continuous)


################################
### ZIELFUNKTION             ###
################################

#    Deckungsbeitrag
# := Umsatz durch regulaeren Verkauf
#  + Umsatz durch Abschlussverkauf
#  - Produktionskosten
#  - Lagerkosten

deckungsbeitrag = (
    
    # Umsatz durch regulaeren Verkauf
    xp.Sum(preis[i, k] * VerkV[i, k] for i in ZEIT for k in PROD)
    
    # + Umsatz durch Abschlussverkauf
    + xp.Sum(redp[k] * LagerV[T, k] for k in PROD)
    
    # - Produktionskosten
    - xp.Sum(pkost[i, j, k, h] * ProdV[i, j, k, h]
             for i in ZEIT for j in STAT for k in PROD for h in MASC)
    
    # - Lagerkosten
    - xp.Sum(lkost[k] * xp.Sum(LagerV[i, k] for i in ZEIT if i < T)
             for k in PROD)
    
    )
    
LP1.setObjective(deckungsbeitrag, sense=xp.maximize)

################################
### NEBENBEDINGUNGEN         ###
################################

# Lagerbilanzgleichung sicherstellen
lagerBilanz = [xp.Sum(ProdV[i, j, k, h] for j in STAT for h in MASC)
               +  LagerV[i-1, k]
               == LagerV[i, k]
               +  VerkV[i, k]
               for i in ZEIT for k in PROD]

# Lageranfangsbestand festsetzen
lagerAnfang = [LagerV[0, k] == 0 for k in PROD]

# Lagerkapazitaet einhalten
lagerKap = [LagerV[i, k] <= lkap[k] for i in ZEIT for k in PROD]

# Maschinenkapazitaet einhalten
maschKap = [xp.Sum(kapver[i, j, k, h] * ProdV[i, j, k, h] for k in PROD)
            <= mkap[i, j, h]
            for i in ZEIT for j in STAT for h in MASC]

# Nachfrage erfuellen
nachfrage = [VerkV[i, k] == nach[i, k] for i in ZEIT for k in PROD]

# Nebenbedingungen zur Instanz hinzufuegen
LP1.addConstraint(lagerBilanz, lagerAnfang, lagerKap, maschKap, nachfrage)

################################
### INSTANZ LOESEN & AUSGABE ###
################################

# Instanz loesen
LP1.lpOptimize()

# Loesung anfordern
prod_menge = LP1.getSolution(ProdV)
lager_bestand = LP1.getSolution(LagerV)
verk_menge = LP1.getSolution(VerkV)
zielfunktion = LP1.attributes.objval

# Ausgabe der Lösung
use.output("Produktionsmenge", prod_menge, ZEIT, STAT, PROD, MASC)
use.output("Lagerbestand", lager_bestand, ZEIT, PROD)
use.output("Verkaufsmenge", verk_menge, ZEIT, PROD)
print()
print("Deckungsbeitrag = ", zielfunktion)

# Ausgabe der Schlupfe der NB zur Maschinenkapazität
use.out_sl(LP1, "maschKap", maschKap, ZEIT, STAT, MASC)

# Ausgabe der Dualwerte der NB zur Maschinenkapazität
use.out_dv(LP1, "maschKap", maschKap, ZEIT, STAT, MASC)

# Ausgabe der reduzierten Kosten der Lagerbestandsvariable
use.out_rc(LP1, "LagerV", LagerV, ZEIT2, PROD)

# Sensitivitätsanalyse für die Lagerbestandsvariable
use.saOFC(LP1, "LagerV", LagerV, ZEIT2, PROD)

# Sensitivitätsanalyse für die NB zur Maschinenkapazität
use.saRHS(LP1, "maschKap", maschKap, ZEIT, STAT, MASC)