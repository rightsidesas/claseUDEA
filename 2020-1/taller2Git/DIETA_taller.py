from pyomo.environ import *
import pandas as pd
from openpyxl import load_workbook
from os import getcwd
import xlwings as xw

#Definir archivo de excel

# Definir objeto con información de excel
xl_dieta = pd.ExcelFile(xlFile)

#Definir modelo de optimizacion


#Indices del modelo
modelo.S_TIPO = Set(initialize=["ACOMPAÑANTE","CARNE","POSTRE"])

#LECTURA DE DATOS;

#Definir variables de decision

#Definir funcion objetivo
def valorAlmuerzo_rule(modeloRobo):
    expr  = sum (DATOS1.COSTO[i]*modeloDieta.x[i] for i in  DATOS1.index) 
    return expr

modeloDieta.ValorAlmuerzo = Objective(rule=valorAlmuerzo_rule, sense=minimize)

#Definir restricciones
#restriccion de Carbohidratos

#restriccion de Vitaminas

#restriccion de Proteinas

#restriccion de Grasas

#restriccion de Tipo de alimento

#Definir Optimizador
opt = SolverFactory('cbc')

#Escribir archivo .lp
modeloDieta.write("archivo.lp",io_options={"symbolic_solver_labels":True})

#Ejecutar el modelo
results = opt.solve(modeloDieta,tee=0,logfile ="dieta.log", keepfiles= 0,symbolic_solver_labels=True)

if (results.solver.status == SolverStatus.ok) and (results.solver.termination_condition == TerminationCondition.optimal):

    #Imprimir Resultados
    print ()
    print ("Valor del almuerzo: ", value(modeloDieta.ValorAlmuerzo))
    print ()

    #Escribir los resultados
    wb = xw.Book(xlFile)
    hoja_out = wb.sheets['resultado']

    out_almuerzo = pd.DataFrame(columns=["Alimento","valor"])

    fila = 0
    for i in DATOS1.index:
        fila += 1
        salida = []
        salida.append(i)
        salida.append(modeloDieta.x[i].value) 
        out_almuerzo.loc[fila] = salida

    hoja_out.range('A1').value = out_almuerzo

elif (results.solver.termination_condition == TerminationCondition.infeasible):
	print()
	print("EL PROBLEMA ES INFACTIBLE")

elif(results.solver.termination_condition == TerminationCondition.unbounded):
	print()
	print("EL PROBLEMA ES INFACTIBLE")
else:
    print("TERMINÓ EJECUCIÓN CON ERRORES")