"""
U26
Verallgemeinertes Investmentproblem: Portfolio-Selektion
@author: Kai Helge Becker
"""

import xpress as xp
import useful_functions as use

# Verknuepfe mit Datenfile
import U26investment_dat as dat

# Initialisiere Instanz
LP1 = xp.problem()

################################
### PARAMETER, MENGEN        ###
################################

### Skalarparameter ##########

# Anzahl Zeitperioden
T = dat.T  
# Anzahl Investmentprodukte
J = dat.J 
# Budget zum Investieren
B = dat.B

### Indexmengen ##############

# Zeitperioden
ZEIT  = dat.ZEIT 
# Investmentprodukte
PROD  = dat.PROD

### Vektorparameter ##########

# Erster moegl. Beginnzeitpunkt des Investmentproduktes (Anfang der Periode) 
start = dat.start # PROD
# Letzter moegl. Beginnzeitpunkt des Investmentproduktes (Anfang der Periode)
end = dat.end # PROD
# Gewinn durch Investmentprodukt (0 < gewinn < 1)
gewinn = dat.gewinn # PROD
# Laufzeit des Investmentproduktes in Perioden
lauf = dat.lauf # PROD


################################
### VARIABLEN                ###
################################

# Budget am Anfang einer Periode
BudgetV = LP1.addVariables(ZEIT, name="BudgetV", 
                           lb=0, vartype=xp.continuous)

# Betrag, der in ein Projekt investiert wird
InvestV = LP1.addVariables(ZEIT, PROD, name="InvestV", 
                           lb=0, vartype=xp.continuous)


################################
### ZIELFUNKTION             ###
################################

# Maximiere Budget am Anfang der letzten Periode

endbudget = BudgetV[T]

LP1.setObjective(endbudget, sense=xp.maximize)

################################
### NEBENBEDINGUNGEN         ###
################################

# Anfangsbudget
anfang = BudgetV[1] == B

# Budgetbilanz fuer weitere Perioden
budgetbilanz = [BudgetV[t]
                == BudgetV[t-1]
                 - xp.Sum(InvestV[t-1, p] 
                          for p in PROD
                          if t-1 >= start[p] and t-1 <= end[p]) 
                 + xp.Sum((1+gewinn[p])*InvestV[t-lauf[p], p] 
                          for p in PROD
                          if t-lauf[p] >= start[p] and t-lauf[p] <= end[p])
               for t in ZEIT if t > 1]

# Maximales Investment
maximal = [xp.Sum(InvestV[t, p] for p in PROD)
           <= BudgetV[t] for t in ZEIT]

# Nebenbedingungen zur Instanz hinzufuegen 
LP1.addConstraint(anfang, budgetbilanz, maximal)

################################
### INSTANZ LOESEN & AUSGABE ###
################################

# Instanz loesen
LP1.lpOptimize()

# Loesung anfordern
budget = LP1.getSolution(BudgetV)
invest = LP1.getSolution(InvestV)
zielfunktion = LP1.attributes.objval

# Ausgabe
use.output("Budget", budget, ZEIT)
use.output("Investmentprodukt", invest, ZEIT, PROD)
print("\nGesamtgewinn = ", zielfunktion - B)
