from pyomo.environ import *

A = ['hammer','wrench','screwdriver','towel']
B = {'hammer':8,'wrench':3,'screwdriver':6,'towel':11}
C = {'hammer':5,'wrench':7,'screwdriver':4,'towel':3}
WMAX = 14

# Modelo de optimizacion de la mochila
mochila = ConcreteModel()

# Definir variable de decisión binaria
mochila.articulo = Var(A, within = Binary)

# Definir función objetivo
mochila.FO = Objective(
    expr = sum(B[i]*mochila.articulo[i] for i in A),
    sense = maximize)

mochila.RPeso = Constraint (
    expr = sum(C[i]*mochila.articulo[i] for i in A) <= WMAX)

opt = SolverFactory('cbc')

mochila.write("archivo.lp",io_options={"symbolic_solver_labels":True})

resultado = opt.solve(mochila, tee=True)

mochila.pprint()