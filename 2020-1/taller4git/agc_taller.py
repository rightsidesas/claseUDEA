from pyomo.environ import *
import pandas as pd
from openpyxl import load_workbook
from os import getcwd
import xlwings as xw

#Definir objeto con archivo de excel
xlFile = getcwd() + r"\agc.XLSx"            

# # Definir objeto con información de excel
xl_datos = pd.ExcelFile(xlFile)

#Definir modelo de optimizacion
modelo = ConcreteModel("taller4")

#LECTURA DE DATOS;

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
    hoja_out = wb.sheets['resultado']

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
            salida.append(modelo.agc[i,t].value) 
        out_.loc[fila] = salida

    hoja_out.range('A1').value = out_

elif (results.solver.termination_condition == TerminationCondition.infeasible):
	print()
	print("EL PROBLEMA ES INFACTIBLE")

elif(results.solver.termination_condition == TerminationCondition.unbounded):
	print()
	print("EL PROBLEMA ES INFACTIBLE")
else:
    print("TERMINÓ EJECUCIÓN CON ERRORES")