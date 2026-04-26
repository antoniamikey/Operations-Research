"""
Abschlussprojekt
Modell-Datei
@author: Antonia Schikora, Cedric Schuster, Jonas Stassiuk, Pascal Moroschan
"""

import xpress as xp
import useful_functions as use
import UpgradefrageAbschlussprojekt_dat as dat

lastgang = xp.problem()

################################
### VARIABLEN                ###
################################

# Gesamtproduktion aller Generatoren von Typ g in Stunde s an Tag t (MW)
ProdV = lastgang.addVariables(dat.GEN, dat.STUNDE, dat.TAG,
                              name="ProdV", lb=0, vartype=xp.continuous)

# Anzahl laufender Generatoren von Typ g in Stunde s an Tag t (ganzzahlig)
LaufGV = lastgang.addVariables(dat.GEN, dat.STUNDE, dat.TAG,
                               name="LaufGV", lb=0, vartype=xp.integer)


# Begründung für continuous: NeuGV erbt Ganzzahligkeit von LaufGV (da NeuGV >= LaufGV[t] - LaufGV[t-1],
# und LaufGV ganzzahlig ist, nimmt NeuGV im Optimum automatisch ganzzahlige Werte an)
NeuGV = lastgang.addVariables(dat.GEN, dat.STUNDE, dat.TAG,
                              name="NeuGV", lb=0, vartype=xp.continuous)

# Binärvariable für Netzwerkstabilität
# NetzV[s,t] = 1, wenn mehr als 3 Generatoren von Typ 1 laufen
NetzV = lastgang.addVariables(dat.STUNDE, dat.TAG,
                              name="NetzV", lb=0, vartype=xp.binary)

################################
### ZIELFUNKTION             ###
################################

gesamtkosten = xp.Sum(
    dat.anlassko[g]   * NeuGV[g, s, t]                                         # Anlasskosten
    + dat.minlastko[g]  * LaufGV[g, s, t]                                      # Fixkosten auf Minimallast
    + dat.kouebermin[g] * (ProdV[g, s, t] - dat.minlast[g] * LaufGV[g, s, t])  # Variable Kosten darüber
    for g in dat.GEN for s in dat.STUNDE for t in dat.TAG)

lastgang.setObjective(gesamtkosten, sense=xp.minimize)

################################
### NEBENBEDINGUNGEN         ###
################################

# Bedarfsdeckung (>= erlaubt minimale Überproduktion, die ZF verhindert sie automatisch)
bedarfsdeckung = [xp.Sum(ProdV[g, s, t] for g in dat.GEN) >= dat.bedarf[s, t]
                  for s in dat.STUNDE for t in dat.TAG]

# Minimalproduktion
produktionsminimum = [ProdV[g, s, t] >= dat.minlast[g] * LaufGV[g, s, t]
                      for g in dat.GEN for s in dat.STUNDE for t in dat.TAG]

# Maximalproduktion
produktionsmaximum = [ProdV[g, s, t] <= dat.maxlast[g] * LaufGV[g, s, t]
                      for g in dat.GEN for s in dat.STUNDE for t in dat.TAG]

# Obergrenze für Anzahl laufender Generatoren.
# Ohne diese NB könnte der Solver theoretisch mehr Generatoren einsetzen als verfügbar
maxAnzGen = [LaufGV[g, s, t] <= dat.verfgen[g]
             for g in dat.GEN for s in dat.STUNDE for t in dat.TAG]

# Oberer Puffer: laufende Generatoren müssen 20% Mehrbelastung ohne Neustart abdecken
obererPuffer = [xp.Sum(dat.maxlast[g] * LaufGV[g, s, t] for g in dat.GEN) >= (1 + dat.puffer_oben) * dat.bedarf[s, t]
                for s in dat.STUNDE for t in dat.TAG]

# Unterer Puffer: Minimallast darf bei 15% Rückgang nicht unterschritten werden
untererPuffer = [xp.Sum(dat.minlast[g] * LaufGV[g, s, t] for g in dat.GEN) <= (1 - dat.puffer_unten) * dat.bedarf[s, t]
                 for s in dat.STUNDE for t in dat.TAG]

# Netzwerkstabilität: nicht mehr als 3 Typ-1 UND mehr als 3 Typ-2 gleichzeitig
# Hinweis: GEN[0] = Typ 1, GEN[1] = Typ 2 — Reihenfolge muss der Excel entsprechen
M1 = dat.verfgen[dat.GEN[0]] - 3  # entspricht hier 3
M2 = dat.verfgen[dat.GEN[1]] - 3  # entspricht hier 5

netzwerkstabilitaet1 = [LaufGV[dat.GEN[0], s, t] <= 3 + M1 * NetzV[s, t]
                        for s in dat.STUNDE for t in dat.TAG]
netzwerkstabilitaet2 = [LaufGV[dat.GEN[1], s, t] <= 3 + M2 * (1 - NetzV[s, t])
                        for s in dat.STUNDE for t in dat.TAG]

# Anlasskosten: Woche ist periodisch
# Übergang Sonntag 23:00 auf Montag 0:00: gleiche Generatoren = keine Anlasskosten
# STUNDE[-1] = letzte Stunde, TAG[-1] = Sonntag
neugestartet_wochenanfang = [
    NeuGV[g, dat.STUNDE[0], dat.TAG[0]] >= LaufGV[g, dat.STUNDE[0], dat.TAG[0]] - LaufGV[g, dat.STUNDE[-1], dat.TAG[-1]]
    for g in dat.GEN]

neugestartet_stunde = [
    NeuGV[g, dat.STUNDE[s], t] >= LaufGV[g, dat.STUNDE[s], t] - LaufGV[g, dat.STUNDE[s-1], t]
    for g in dat.GEN for s in range(1, len(dat.STUNDE)) for t in dat.TAG]

neugestartet_tag = [
    NeuGV[g, dat.STUNDE[0], dat.TAG[t]] >= LaufGV[g, dat.STUNDE[0], dat.TAG[t]] - LaufGV[g, dat.STUNDE[0], dat.TAG[t-1]]
    for g in dat.GEN for t in range(1, len(dat.TAG))]

# Alle Nebenbedingungen hinzufügen
lastgang.addConstraint(
    bedarfsdeckung, produktionsminimum, produktionsmaximum, maxAnzGen,
    obererPuffer, untererPuffer,
    netzwerkstabilitaet1, netzwerkstabilitaet2,
    neugestartet_wochenanfang, neugestartet_stunde, neugestartet_tag)

################################
### INSTANZ LOESEN & AUSGABE ###
################################

lastgang.mipOptimize()

# Loesung anfordern
prod_plan  = lastgang.getSolution(ProdV)
laufend    = lastgang.getSolution(LaufGV)
neugestart = lastgang.getSolution(NeuGV)
gesamtkosten_wert = lastgang.attributes.objval

# Konsolenausgabe
use.output("Produktion (MW)", prod_plan, dat.GEN, dat.STUNDE, dat.TAG)
use.output("Laufende Generatoren", laufend, dat.GEN, dat.STUNDE, dat.TAG)
use.output("Neu gestartete Generatoren", neugestart, dat.GEN, dat.STUNDE, dat.TAG)
print("\nGesamtkosten:", gesamtkosten_wert, "EUR")

################################
### EXCEL-AUSGABE            ###
################################

# HINWEIS: Vor dem ersten Ausführen dieses Skripts muss im selben Ordner
# eine leere Excel-Datei mit dem Namen "ProjektErgebnisUpgrade.xlsx" erstellt werden.
print("HINWEIS: Vor dem ersten Ausführen dieses Skripts muss im selben Ordner eine leere Excel-Datei mit dem Namen 'ProjektErgebnisUpgrade.xlsx' erstellt werden")
# Dazu einfach eine neue Excel-Datei anlegen, speichern und wieder schließen.
# Die Datei muss beim Ausführen des Skripts geschlossen sein.
workbook_out, sheet_out = use.open_sheet("ProjektErgebnisUpgrade.xlsx", 0)

# Produktionsplan: Gesamtproduktion je Generatortyp, Stunde und Tag (MW).
use.write_thead(sheet_out, "A1", "Produktion (MW)", "Generator", "Stunde", "Wochentag")
use.write_tbody(sheet_out, "A1", prod_plan, dat.GEN, dat.STUNDE, dat.TAG)


# Anzahl laufender Generatoren je Typ, Stunde und Tag (ganzzahlig).
use.write_thead(sheet_out, "K1", "Laufende Generatoren", "Generator", "Stunde", "Wochentag")
use.write_tbody(sheet_out, "K1", laufend, dat.GEN, dat.STUNDE, dat.TAG)

# Kostenaufschlüsselung

# Anlasskosten je Generatortyp (EUR):
# anlassko[g] × Summe aller Neustarts über alle Stunden und Tage
anlasskosten_gen = {g: sum(dat.anlassko[g] * neugestart[g, s, t]
                           for s in dat.STUNDE for t in dat.TAG)
                    for g in dat.GEN}

# Fixkosten auf Minimallast je Generatortyp (EUR):
# minlastko[g] × Summe aller Laufstunden
fixkosten_gen = {g: sum(dat.minlastko[g] * laufend[g, s, t]
                        for s in dat.STUNDE for t in dat.TAG)
                 for g in dat.GEN}

# Variable Kosten oberhalb der Minimallast je Generatortyp (EUR):
# kouebermin[g] × (Gesamtproduktion − Minimallast × Anzahl laufender Generatoren)
varkosten_gen = {g: sum(dat.kouebermin[g] * (prod_plan[g, s, t] - dat.minlast[g] * laufend[g, s, t])
                        for s in dat.STUNDE for t in dat.TAG)
                 for g in dat.GEN}

# Ausgabe der drei Kostenkomponenten
use.write_thead(sheet_out, "U1", "Anlasskosten (EUR)", "Generator")
use.write_tbody(sheet_out, "U1", anlasskosten_gen, dat.GEN)

use.write_thead(sheet_out, "U6", "Fixkosten auf Minimallast (EUR)", "Generator")
use.write_tbody(sheet_out, "U6", fixkosten_gen, dat.GEN)

use.write_thead(sheet_out, "U11", "Variable Kosten oberhalb Minimallast (EUR)", "Generator")
use.write_tbody(sheet_out, "U11", varkosten_gen, dat.GEN)

# Gesamtkosten (= Zielfunktionswert)
use.write_scalar(sheet_out, "U16", gesamtkosten_wert, "Gesamtkosten (EUR)")

# Ergebnisse in der Ausgabedatei speichern.
use.save_sheet(workbook_out, "ProjektErgebnisUpgrade.xlsx")