"""
U12/U31 Produktionsplanung mit mehreren Standorten
Analyse der Loesung
"""
import xpress as xp

U21 = xp.problem()

###### Variablen ######

# Produkte von Standort 1
xgr1 = U21.addVariable() # Anzahl hergestellter grosser Produkte
xmi1 = U21.addVariable() # Anzahl hergestellter mittlerer Produkte
xkl1 = U21.addVariable() # Anzahl hergestellter kleiner Produkte

# Produkte von Standort 2
xgr2 = U21.addVariable() # Anzahl hergestellter grosser Produkte
xmi2 = U21.addVariable() # Anzahl hergestellter mittlerer Produkte
xkl2 = U21.addVariable() # Anzahl hergestellter kleiner Produkte

# Produkte von Standort 3
xgr3 = U21.addVariable() # Anzahl hergestellter grosser Produkte
xmi3 = U21.addVariable() # Anzahl hergestellter mittlerer Produkte
xkl3 = U21.addVariable() # Anzahl hergestellter kleiner Produkte

# Gleicher genutzter Prozentsatz der Kapazitaet
proz = U21.addVariable()


##### Nebenbedingungen #####

# Freie Kapazitaeten der Standorte nicht unterschreiten
standort1kap = xgr1 + xmi1 + xkl1 <= 500
standort2kap = xgr2 + xmi2 + xkl2 <= 600
standort3kap = xgr3 + xmi3 + xkl3 <= 300

# Lagerkapazitaeten der Standorte einhalten
standort1lager = 20*xgr1 + 15*xmi1 + 12*xkl1 <= 5000
standort2lager = 20*xgr2 + 15*xmi2 + 12*xkl2 <= 8000
standort3lager = 30*xgr3 + 15*xmi3 + 12*xkl3 <= 3500

# Mehr als Nachfrage kann nicht verkauft werden
nachgross  = xgr1 + xgr2 + xgr3 <= 600
nachmittel = xmi1 + xmi2 + xmi3 <= 800
nachklein  = xkl1 + xkl2 + xkl3 <= 500

# Prozentsatz genutzter freier Kapazitaet bei allen Standorten gleich
prozstandort1 = (xgr1 + xmi1 + xkl1)/500 == proz
prozstandort2 = (xgr2 + xmi2 + xkl2)/600 == proz
prozstandort3 = (xgr3 + xmi3 + xkl3)/300 == proz

U21.addConstraint(standort1kap, standort2kap,standort3kap,
                  standort1lager, standort2lager, standort3lager,
                  nachgross, nachmittel, nachklein,
                  prozstandort1, prozstandort2, prozstandort3) 
                
##### Zielfunktion: Maximiere Nettogewinn #####

nettogewinn = (   12*(xgr1 + xgr2 + xgr3)
               +  10*(xmi1 + xmi2 + xmi3)
               +  10.25*xkl1 + 9*xkl2 + 9*xkl3)

U21.setObjective(nettogewinn, sense=xp.maximize)

##### Loesen und Ausgabe #####
U21.lpOptimize()
ZFWert = U21.attributes.objval
ProdGr1 = U21.getSolution(xgr1)
ProdMi1 = U21.getSolution(xmi1) 
ProdKl1 = U21.getSolution(xkl1)
ProdGr2 = U21.getSolution(xgr2)
ProdMi2 = U21.getSolution(xmi2)
ProdKl2 = U21.getSolution(xkl2)
ProdGr3 = U21.getSolution(xgr3)
ProdMi3 = U21.getSolution(xmi3)
ProdKl3 = U21.getSolution(xkl3)
print("Produzierte Mengeneinheiten Standort 1:", ProdGr1, ProdMi1, ProdKl1)
print("Produzierte Mengeneinheiten Standort 2:", ProdGr2, ProdMi2, ProdKl2)
print("Produzierte Mengeneinheiten Standort 3:", ProdGr3, ProdMi3, ProdKl3)
print("Gesamtgewinn:", ZFWert) 
freieKapStandort1 = U21.getSolution(500 - xgr1 - xmi1 - xkl1) 
freieKapStandort2 = U21.getSolution(600 - xgr2 - xmi2 - xkl2)
freieKapStandort3 = U21.getSolution(300 - xgr3 - xmi3 - xkl3)
print("Freie Kapazitaet Standort 1:", freieKapStandort1)
print("Freie Kapazitaet Standort 2:", freieKapStandort2)
print("Freie Kapazitaet Standort 3:", freieKapStandort3)
anteilGenutzterKap = U21.getSolution(proz)
print("Anteil der genutzten Kapazitaet:", anteilGenutzterKap)
lagerGenutztStand1 = U21.getSolution(20*xgr1 + 15*xmi1 + 12*xkl1)
lagerGenutztStand2 = U21.getSolution(20*xgr2 + 15*xmi2 + 12*xkl2)
lagerGenutztStand3 = U21.getSolution(20*xgr3 + 15*xmi3 + 12*xkl3)
print("Quadratmeter Lager genutzt an Standort 1:", lagerGenutztStand1)
print("Quadratmeter Lager genutzt an Standort 2:", lagerGenutztStand2)
print("Quadratmeter Lager genutzt an Standort 3:", lagerGenutztStand3)

# ********************************************************************
# Beispiele für Schlupf, Werte der Dualvariablen und reduzierte Kosten
# ********************************************************************

# Man kann Schlupf, Dualwerte und reduzierte Kosten für einzelne Variablen
# bzw. NB abfragen, indem man die Namen der Variablen/NB eingibt. 

print("\nEinzelwerte")
schlupf = U21.getSlacks(standort3lager)
dualwerte = U21.getDuals(standort1lager)
redkosten = U21.getRedCosts(xgr1)
print("Schlupf für Lagerkapazität Standort 3:", schlupf)
print("Dualwert für Lagerkapazität Standort 1:", dualwerte)
print("Reduzierte Kosten für großes Produkt an Standort 1:", redkosten)

# Lässt man die Klammern leer, erhält man die Werte für alle Variablen/NB
# in der Reihenfolge, in der die Variablen in der Zielfunktion stehen
# bzw. in der die NB dem Problem hinzugefügt wurden.

print("\nAlle Werte")
schlupf = U21.getSlacks()
dualwerte = U21.getDuals()
redkosten = U21.getRedCosts()
print("Schlupf:", schlupf)
print("Dualwerte:", dualwerte)
print("Reduzierte Kosten:", redkosten)

# ************************************
# Beispiele zur Sensitivitaetsanalyse
# ************************************

# Herausfinden, in welchem Intervall der Gewinn, 
# also der Zielfunktionskoeffizient (ZFK), 
# des kleinen Produkts an Standort 1 schwanken kann, ohne dass sich
# das Optimum ändert.

kl1unten, kl1oben = U21.objSA([xkl1])
print(kl1unten, kl1oben)
print("\nIntervall für ZFK des kleinen Produkts an Standort 1: von ", 
      kl1unten, " bis ", kl1oben)

# Herausfinden, in welchem Intervall die Lagerkapazität, 
# also die rechte Seite der NB zur Lagerkapazität,
# an Standort 1 schwanken kann, ohne dass sich der Dualwert der NB
# im Optimum ändert.

lager1unten, lager1oben = U21.rhsSA([standort1lager])
print("\nIntervall rechte Seite bei Lager an Standort 1: von ",
      lager1unten, " bis ", lager1oben)

# Es ist auch möglich, die Intervalle für mehrere ZFK 
# bzw. mehrere rechte Seiten in eins abzufragen. Dann gibt man einfach
# eine Liste von denjenigenen an, die man abfragen will. Im folgenden 
# werden die Intervalle für alle ZFK und alle rechten Seiten abgefragt.
# Im Folgenden sieht man auch, dass man anstelle der Namen der Variablen
# bzw. NB auch deren Nummern nehmen kann. Die Variablen und NB werden dabei von
# 0 an in der Reihenfolge nummeriert, indem sie dem Modell hinzugefügt 
# wurden.

ZFKunten, ZFKoben = U21.objSA([0, 1, 2, 3, 4, 5, 6, 7, 8, 9])
print("\nIntervalle für ZFK: von ", ZFKunten, " bis ", ZFKoben)

RHSunten, RHSoben = U21.rhsSA([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11])
print("\nIntervalle für rechte Seiten: von ", RHSunten, " bis ", RHSoben)