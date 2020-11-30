# -*- coding: utf-8 -*-
"""
Created on Tue Nov 24 17:53:40 2020

@author: Joe
"""

# Codigo pronóstico Demanda SIN
import pandas as pd 
import os 
# lectura datos

cwd = os.getcwd()

# Lectura Datos
DemandaSin = pd.read_excel(cwd  + r'\data.xlsx',
                           sheet_name='DemandaSIN') 
# Días de la semana
DemandaSin.Fecha = pd.to_datetime(DemandaSin.Fecha)
DemandaSin['Dia']= DemandaSin.Fecha.apply(lambda x : x.strftime("%A")) # tipo de día

# Gráficos Serie de demanda
import plotly.graph_objects as go
from plotly.offline import plot
fig = go.Figure()
fig.add_trace(go.Scatter(x=DemandaSin['Fecha'][:365], y=DemandaSin['Demanda'][:365],
                         name='Demanda',
                         mode='lines+markers',
                         marker = dict(color = 'rgb(255,0,0)'),
                         line=dict(color='rgba(7,7,7,0.5)', width=1)))
fig.update_layout(
    template = "plotly_white",
     title="Demanda Energia SIN Mwh",
     yaxis_title="GWh",
)
plot(fig, filename='Demanda.html')

# Días festivos
import holidays_co
DiaFestivo = ["Festivo" 
              if holidays_co.is_holiday_date(i) 
              else "Ordinario" 
              for i in DemandaSin.Fecha]
DemandaSin['DiaFestivo'] = DiaFestivo
# descripción
AnalisisDatos = DemandaSin[['Dia','Demanda']].groupby('Dia').agg('describe')
AnalisisDatos.to_excel('descripciondatos.xlsx')

import plotly.express as px

# Gráfico 1
fig = px.box(DemandaSin.iloc[:365], x="Dia", y="Demanda", points="all")
fig.update_layout(
    template = "plotly_white",
     title="Demanda Energia SIN Mwh",
     yaxis_title="GWh",
)
plot(fig, filename='DemandaDia.html')

# Grafico 2
fig = px.box(DemandaSin.iloc[:365], x="Dia", y="Demanda", points="all",
             color="DiaFestivo")
fig.update_layout(
    template = "plotly_white",
     title="Demanda Energia SIN Mwh",
     yaxis_title="GWh",
)
plot(fig, filename='DemandaDiaFestivo.html')


# Modelo 
data = DemandaSin.copy()
maximo = max(data.Demanda)
minimo = min(data.Demanda)
data['estandar'] = (data.Demanda - minimo)/(maximo - minimo)


# Rezagos variables
data['R7'] = data.estandar.shift(7)
data['R14'] = data.estandar.shift(14)
data.dropna(axis = 0, inplace = True)


# matriz datos

import numpy as np

# variables independientes
xmodelo = np.array(data[['R7','R14']])
xdim = xmodelo.shape
xmodelo = xmodelo.reshape(xdim[0],1,xdim[1])

# variable dependiente
ymodelo=np.array(data['estandar'])
ydim = ymodelo.shape
ymodelo=ymodelo.reshape(ydim[0],1)

# datos entreno
xentreno = xmodelo[:-10]
xvalidacion = xmodelo[-10:]
yentreno = ymodelo[:-10]
yvalidacion = ymodelo[-10:]



from keras.models import Sequential
from keras.layers import LSTM , Dense
from keras.optimizers import RMSprop

# Definición modelo
modelo = Sequential()
modelo.add(LSTM(20, return_sequences=True,activation='elu',
               input_shape=(1, xdim[1])))
#modelo.add(LSTM(5,activation='relu', return_sequences=True))
modelo.add(LSTM(1, activation='elu'))
modelo.compile(loss='mean_squared_error',
              optimizer='rmsprop')
modelo.fit(xentreno, yentreno,
          batch_size=20, epochs=5,
          validation_data=(xvalidacion, yvalidacion))


# Pronóstico

data.set_index(['Fecha'], inplace = True)
pronosticoDT = data.loc['2018-3-1':'2018-3-10']

# matriz
xpred = np.array(pronosticoDT[['R7','R14']])
xdimp = xpred.shape
xpred = xpred.reshape(xdimp[0],1,xdimp[1])

# pronóstico
pronostico_m = modelo.predict(xpred)

salida = [float(i) for i in pronostico_m] 
salida = [i*(maximo-minimo)+minimo for i in salida]

datafinal = pronosticoDT.Demanda.tolist()

Resultado = pd.DataFrame({'Demanda': datafinal, 'Pronostico': salida,
              'Fecha' : pronosticoDT.index.tolist()})



Resultado['ERROR'] = (Resultado.Demanda - Resultado.Pronostico)/Resultado.Demanda

fig = go.Figure()
fig.add_trace(go.Scatter(x=Resultado['Fecha'], y=Resultado['Demanda'],
                         name='Demanda',
                         mode='lines+markers',
                         marker = dict(color = 'rgb(255,0,0)'),
                         line=dict(color='rgba(7,7,7)', width=1)))
fig.add_trace(go.Scatter(x=Resultado['Fecha'], y=Resultado['Pronostico'],
                         name='Pronostico',
                         mode='lines+markers',
                         marker = dict(color = 'rgb(7,7,7)'),
                         line=dict(color='rgb(34,139,34)', width=1)))

fig.update_layout(
    template = "plotly_white",
     title="Demanda Energia SIN Mwh",
     yaxis_title="GWh",
)
plot(fig, filename='ResultadoFinal.html')


Resultado['ERROR'] = ((Resultado.Demanda - Resultado.Pronostico)/Resultado.Demanda)*100

fig = go.Figure()
fig.add_trace(go.Scatter(x=Resultado['Fecha'], y=Resultado['ERROR'],
                         name='Residual',
                         mode='lines+markers',
                         marker = dict(color = 'rgb(255,0,0)'),
                         line=dict(color='rgba(7,7,7)', width=1)))
fig.update_layout(
    template = "plotly_white",
     title="Residuales Demanda Energia SIN Mwh",
     yaxis_title="GWh",
)
plot(fig, filename='Residual.html')


Resultado.to_excel('Resultado final.xlsx')


# pronósticar dias festivos, como hacemos para que el modelo aprenda este tipo de dias

 

















