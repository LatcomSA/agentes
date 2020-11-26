# -*- coding: utf-8 -*-
"""
Created on Wed Oct 14 17:28:35 2020

@author: jasso
"""
import pandas as pd
from scipy.stats import poisson
import math
import numpy as np
import openpyxl 

serie = pd.read_csv('pronostico_wfm.csv', sep=',', header=0, index_col=None)

del(serie['Nombre de cola'])
agents = np.zeros((serie.shape[0],1))
serie = serie.to_numpy()
serie = np.concatenate((serie,agents), axis=1)


interval = 1800
AWT = 20
sla_per = 0.8
for x in range(serie.shape[0]):    
    agents = 1
    try:
        AHT = float(serie[x,4])        
        if AHT == 0:
           serie[x,5] =  0
           continue
        offered = float(serie[x,3])
        arrv_rate = offered/interval
        traf_inte = arrv_rate*AHT
        occup = traf_inte/agents
        erlang_c = poisson.pmf(agents,traf_inte,loc=0)/(poisson.pmf(agents,traf_inte,loc=0)
                                            +(1-occup)*(poisson.cdf(agents-1,traf_inte,loc=0)))
        SLA = 1-erlang_c*math.exp(-(agents-traf_inte)*(AWT/AHT))
    except:
        serie[x,5] =  0
    else:    
        while SLA < sla_per:
              agents += 1
              occup = traf_inte/agents
              erlang_c = poisson.pmf(agents,traf_inte,loc=0)/(poisson.pmf(agents,traf_inte,loc=0)
                                            +(1-occup)*(poisson.cdf(agents-1,traf_inte,loc=0)))
              SLA = 1-erlang_c*math.exp(-(agents-traf_inte)*(AWT/AHT))
        #ASA = (erlang_c*AHT)/(agents*(1-occup))
        serie[x,5] = agents 

sig_moth  = openpyxl.Workbook()
sheet_moth = sig_moth.active
for x in serie:
    sheet_moth.append([x[0],x[1],x[2],x[3],x[4],x[5]])
sig_moth.save('./agentes_convergencia.xlsx')  