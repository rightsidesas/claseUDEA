from pyomo.environ import *
import pandas as pd
from openpyxl import load_workbook
from os import getcwd
import xlwings as xw
from pandasql import sqldf

#Definir objeto con archivo de excel
xlFile = getcwd() + r"\despacho.XLSx"            

# # Definir objeto con información de excel
xl_datos = pd.ExcelFile(xlFile)

#Definir modelo de optimizacion
modelo = ConcreteModel("taller6")

#LECTURA DE DATOS;
df_DEMANDA = xl_datos.parse('DEMANDA').set_index(['periodo']) 
df_GENERADORES = xl_datos.parse('GENERADORES').set_index(['nombre','periodo']) 
iGenerador  = sqldf("SELECT distinct(nombre) FROM df_GENERADORES;", locals())

#Indices del modelo
modelo.GENERADOR = Set(initialize=iGenerador.nombre)
modelo.PERIODO = Set(initialize=df_DEMANDA.index)

#Definir variables de decision
modelo.despacho = Var( modelo.GENERADOR, modelo.PERIODO,domain=PositiveReals )
modelo.binaria = Var( modelo.GENERADOR, modelo.PERIODO, domain=Binary )

#Definir funcion objetivo
def obj_rule(modelo):
    expr  = sum ( df_GENERADORES.precio[i,t] * modelo.despacho[i,t] 
                     for i in modelo.GENERADOR
                     for t in modelo.PERIODO
                )
    return expr
modelo.FO = Objective(rule=obj_rule, sense=minimize)

#Definir restricciones
#restriccion Balance demanda
def r1_rule(modelo,t):
    ecuacion = sum(modelo.despacho[i,t] 
                        for i in modelo.GENERADOR
                    )
    return ecuacion >= df_DEMANDA.demanda[t]
modelo.r1 = Constraint(modelo.PERIODO,rule=r1_rule)    

#restriccion Maximo
def r2_rule(modelo,i,t):
    return modelo.despacho[i,t]  - df_GENERADORES.maximo[i,t]*modelo.binaria[i,t] <= 0
modelo.r2 = Constraint(modelo.GENERADOR,modelo.PERIODO,rule=r2_rule)    

#restriccion Mínimo AGC
def r3_rule(modelo,i,t):
    return modelo.despacho[i,t]  - df_GENERADORES.minimo[i,t]*modelo.binaria[i,t] >= 0
modelo.r3 = Constraint(modelo.GENERADOR,modelo.PERIODO,rule=r3_rule)    

#Definir Optimizador
opt = SolverFactory('cbc')

#Escribir archivo .lp
modelo.write("archivo.lp",io_options={"symbolic_solver_labels":True})

#Ejecutar el modelo
results = opt.solve(modelo,tee=0,logfile ="archivo.log", keepfiles= 0,symbolic_solver_labels=True)

if (results.solver.status == SolverStatus.ok) and (results.solver.termination_condition == TerminationCondition.optimal):

    #Imprimir Resultados
    print ()
    print ("funcion objetivo ", value(modelo.FO))

    #Escribir los resultados
    wb = xw.Book(xlFile)
    hoja_out = wb.sheets['despacho']

    columnas = ["GENERADOR"]
    for t in modelo.PERIODO:
        columnas.append(t)

    out_ = pd.DataFrame(columns=columnas)

    fila = 0
    for i in modelo.GENERADOR:
        fila += 1
        salida = []
        salida.append(i)
        for t in modelo.PERIODO:
            salida.append(modelo.despacho[i,t].value) 
        out_.loc[fila] = salida
    
    hoja_out.clear_contents()   
    hoja_out.range('A1').value = out_

elif (results.solver.termination_condition == TerminationCondition.infeasible):
	print()
	print("EL PROBLEMA ES INFACTIBLE")

elif(results.solver.termination_condition == TerminationCondition.unbounded):
	print()
	print("EL PROBLEMA ES INFACTIBLE")
else:
    print("TERMINÓ EJECUCIÓN CON ERRORES")