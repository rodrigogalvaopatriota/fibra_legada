#git remote add origin git@github.com:rodrigogalvaopatriota/fibra_legada.git
#git branch -M master
#git add fibra_legada_.ipynb
#git commit -m "message"
#git push -u origin master.


import pandas as pd
import numpy as np
import os
import datetime as dt
from calendar import monthrange

#GET LAST DAY OF MONTH
data_atual = dt.datetime.today()
last_date = data_atual.replace(day=monthrange(data_atual.year,data_atual.month)[1])
last_day_month = last_date.day


#teste
dt_now = dt.datetime.now()
sec = dt_now.second
minute = dt_now.minute
hour = dt_now.hour
day_now = dt_now.day
month = dt_now.month
month = month 
#end_day = 30 - day


#rompimento(prioridade 97,98,99) e atenuacao(prioridade 21):soma da qtde div km regional ou uf * 1000.TETE tete.b
#backbone np (colun backbone: bbn, bbr) divide dado da coluna ay: np pelo total
#backbone tmr (colun backbone: bbn, bbr) tempo medio coluna bj:TEMPO_FIBRA_PSR_EM_HORAS
#acesso np (colun backbone: ninf, bba, outros) divide dado da coluna ay: np pelo total
#acesso tmr (colun backbone: ninf, bba, outros) tempo medio coluna bj:TEMPO_FIBRA_PSR_EM_HORAS
#ftth primario col f = fo, col g = prefo,col v = ora ou vazio, col bi = s 
#ftth secundario col f = fo, col g = prfth, col bi = s 


path = os.getcwd()
path_telegram = '//node'
file = '//Fibra_Optica_ico_base.xlsx'
file_gh = '//LEGAL RE1.xlsx'
file_cod_enc = '//COD_ENCER.xlsx'
file_causa_portal = '//causas.xlsx'
file_telegram = '//bot-legada'
file_km = '//KM_FO.xlsx'

table_base = pd.read_excel(path+file)
table_cod_enc = pd.read_excel(path+file_cod_enc)
table_causa_portal = pd.read_excel(path+file_causa_portal)
table_gh = pd.read_excel(path+file_gh)
table_km = pd.read_excel(path+file_km)


tb_base = table_base
tb_cod_enc = table_cod_enc
tb_causa_portal = table_causa_portal
tb_gh = table_gh
tb_km = table_km

tb_base = tb_base[['UF','BA','AREA_TECNICA','COS','COS_ORIGEM','PRIORIDADE','NOME_TECNICO','MATRICULA_TECNICO','BACKBONE','RAMIFICACAO','FTTH','ABERTURA','TEMPO_FIBRA_SEGUNDO','TEMPO_FIBRA_PSR_EM_HORAS','INI_ACIONAMENTO','ENCE_ACIONAMENTO','PRAZO_PSR_2016','EXPURGO','COD_ENCERRAMENTO']].copy()
tb_cod_enc = tb_cod_enc[['COD_ENCERRAMENTO','CAUSA3']].copy()
tb_causa_portal = tb_causa_portal[['BA','CAUSA']].copy()


#tb_gh = tb_gh[['Nome Complet','Desc. Coord.']].copy()
#tb_gh = tb_gh[['funcionario_id','coordenador','gerente_negocios']].copy()


#tb_gh = tb_gh.rename(columns={'RT' : 'MATRICULA_TECNICO'})
#tb_gh = tb_gh.rename(columns={'Nome Complet' : 'NOME_TECNICO'})
tb_gh = tb_gh.rename(columns={'descricao' : 'NOME_TECNICO'})

tb_causa_portal = tb_causa_portal.rename(columns={'CAUSA':'CAUSA_PORTAL'})

vl_base = tb_base['UF'].count()
print('count original '+str(vl_base))

#FILTERS
#tb_base = tb_base[tb_base['UF'].isin(['PR','AM','PA','BA','AP','RR'])]
tb_base = tb_base[tb_base['UF'].isin(['PR'])]

vl_uf = tb_base['UF'].count()
print('count_after_uf '+str(vl_uf))

tb_base = tb_base[tb_base['AREA_TECNICA'].isin(['FO'])]
#tb_base = tb_base[tb_base['FTTH'].isin(['N'])]
tb_base = tb_base[tb_base['EXPURGO'].isin(['N','T'])]




#tb_base['COS_REDUCE'] = tb_base['COS'].str[-3:]
#tb_base = tb_base[tb_base['COS_REDUCE'].isin(['EFO'])]

tb_base['MONTH'] = pd.to_datetime(tb_base['ABERTURA']).dt.month
tb_base['YEAR'] = pd.to_datetime(tb_base['ABERTURA']).dt.year

tb_base = tb_base[tb_base['MONTH'].isin([month])]
vl_others_filters = tb_base['UF'].count()
print('count_after_others_filters '+str(vl_others_filters))


#ATENUACAO, ROMPIMENTO
tb_base['STATUS_AT_ROMP']=np.where(tb_base['PRIORIDADE']==21,'ATENUACAO',np.where(tb_base['PRIORIDADE']==97,'ROMPIMENTO',np.where(tb_base['PRIORIDADE']==98,'ROMPIMENTO',np.where(tb_base['PRIORIDADE']==99,'ROMPIMENTO','OTHER'))))
#BACKBONE,ACESSO
tb_base['STATUS_BACKBONE_ACESSO']=np.where(tb_base['BACKBONE']=='BBN','BACKBONE',np.where(tb_base['BACKBONE']=='BBR','BACKBONE',np.where(tb_base['BACKBONE']=='NINF','ACESSO',np.where(tb_base['BACKBONE']=='BBA','ACESSO',np.where(tb_base['BACKBONE']=='OUTROS','ACESSO','OTHER')))))


#FTTH
tb_base = tb_base.fillna('DADOS N LOCALIZADOS')
tb_base['kee_status_ftth'] = tb_base['COS'].map(str)+'_'+tb_base['RAMIFICACAO'].map(str)+'_'+tb_base['FTTH'].map(str)
tb_base['STATUS_FTTH'] = np.where((tb_base['kee_status_ftth'] == 'PREFO_ORA_S') | (tb_base['kee_status_ftth'] == 'PREFO_DADOS N LOCALIZADOS_S'),'PRIMARIO',
                         np.where(tb_base['kee_status_ftth'].str[:5] == 'PRFTH','SECUNDARIO','OUTROS'))




#tb_base['TMR2'] = (pd.to_datetime(tb_base['ENCE_ACIONAMENTO']))-(pd.to_datetime(tb_base['ENCE_ACIONAMENTO']))
#tb_base['TMR2'] = tb_base['ENCE_ACIONAMENTO'] - tb_base['INI_ACIONAMENTO']
tb_base['NOME_TECNICO'] = tb_base['NOME_TECNICO'].astype(str)
tb_gh['NOME_TECNICO'] = tb_gh['NOME_TECNICO'].astype(str)




#tb_base = tb_base.fillna('DADOS N LOCALIZADOS')

print('executado.')



#TMR
tb_base['result_hour'] = tb_base['TEMPO_FIBRA_SEGUNDO'] / 60 / 60
#tb_base['result_hour'] 

hour_meta_backbone = 5
hour_meta_acesso = 7

tb_base['tmr_meta'] = np.where((tb_base['result_hour'] > hour_meta_acesso)&(tb_base['STATUS_BACKBONE_ACESSO'] == 'ACESSO'),'FORA DA META',np.where((tb_base['result_hour'] > hour_meta_backbone)&(tb_base['STATUS_BACKBONE_ACESSO'] == 'BACKBONE'),'FORA DA META','DENTRO DA META'))

tb_base['status_tmr_kee'] = tb_base['STATUS_BACKBONE_ACESSO'].map(str)+'_'+tb_base['tmr_meta'].map(str)



#ROMPIMENTO
km = 22200

meta_rompimento = 7


tb_base_rompimento = tb_base[tb_base['PRIORIDADE'].isin([97,98,99])]
tb_base_rompimento = tb_base_rompimento[tb_base_rompimento['AREA_TECNICA'].isin(['FO'])]
tb_base_rompimento = tb_base_rompimento[tb_base_rompimento['COS'].isin(['PREFO'])]
tb_base_rompimento = tb_base_rompimento[tb_base_rompimento['FTTH'].isin(['N'])]
tb_base_rompimento = tb_base_rompimento[tb_base_rompimento['MONTH'].isin([month])]
vl_rompimento = tb_base_rompimento['UF'].count()

result_rompimento =(vl_rompimento / km) * 1000
#PROJECAO
meta_rompimento_dia = meta_rompimento / last_day_month
meta_rompimento_projecao = ((result_rompimento / day_now)* last_day_month)
print('projecao_rompimento: '+str(meta_rompimento_projecao))

print('VALORES rompimento:')
print('count_rompimento '+str(vl_rompimento))
print('result_rompimento '+str(result_rompimento))



#ATENUACAO  
meta_atenuacao =  0.4
tb_base_atenuacao = tb_base[tb_base['PRIORIDADE'].isin([21])]
tb_base_atenuacao = tb_base_atenuacao[tb_base_atenuacao['AREA_TECNICA'].isin(['FO'])]
tb_base_atenuacao = tb_base_atenuacao[tb_base_atenuacao['COS'].isin(['PREFO'])]
tb_base_atenuacao = tb_base_atenuacao[tb_base_atenuacao['FTTH'].isin(['N'])]
tb_base_atenuacao = tb_base_atenuacao[tb_base_atenuacao['MONTH'].isin([month])]
vl_atenuacao = tb_base_atenuacao['UF'].count()

result_atenuacao =(vl_atenuacao / km) * 1000
#PROJECAO
meta_atenuacao_dia = meta_atenuacao / last_day_month
meta_atenuacao_projecao = ((result_atenuacao / day_now)* last_day_month)
print('projecao_atenuacao: '+str(meta_atenuacao_projecao))

print('VALORES atenuacao:')
print('count_atenuacao '+str(vl_atenuacao))
print('result_atenuacao '+str(result_atenuacao))

#LEGADA: NO PRAZO, TMR 

#BACKBONE
#FILTERS
meta_backbone_np = 0.85
tb_base_backbone = tb_base[tb_base['BACKBONE'].isin(['BBN','BBR'])]
tb_base_backbone = tb_base_backbone[tb_base_backbone['AREA_TECNICA'].isin(['FO'])]
tb_base_backbone = tb_base_backbone[tb_base_backbone['COS'].isin(['PREFO'])]
tb_base_backbone = tb_base_backbone[tb_base_backbone['FTTH'].isin(['N'])]
tb_base_backbone = tb_base_backbone[tb_base_backbone['MONTH'].isin([month])]
vl_backbone = tb_base_backbone['UF'].count()
#TMR
vl_backbone_media_tmr = tb_base_backbone['result_hour'].mean()
#NP
tb_base_backbone_no_prazo = tb_base_backbone[tb_base_backbone['PRAZO_PSR_2016'].isin(['NP'])]
tb_base_backbone_fora_prazo = tb_base_backbone[tb_base_backbone['PRAZO_PSR_2016'].isin(['FP'])]

vl_backbone_np = tb_base_backbone_no_prazo['UF'].count()
vl_backbone_fp = tb_base_backbone_fora_prazo['UF'].count()
backbone_result = vl_backbone_np / vl_backbone

#PROJECAO NP
meta_backbone_dia = meta_backbone_np / last_day_month
meta_backbone_projecao = ((backbone_result / day_now)* last_day_month)
print('projecao_backbone: '+str(meta_backbone_projecao))

#PROJECAO TMR BACKBONE
meta_backbone_dia_tmr = hour_meta_backbone / last_day_month
meta_backbone_projecao_tmr = ((vl_backbone_media_tmr / day_now)* last_day_month)
print('projecao_backbone: '+str(meta_backbone_projecao_tmr))

print('VALORES NP BACKBONE:')
print('count_backbone '+str(vl_backbone))
print('np_backbone '+str(vl_backbone_np))
print('fp_backbone '+str(vl_backbone_fp))
print('np_result_backbone '+str(backbone_result))
print('mean_backbone_tmr '+str(vl_backbone_media_tmr))


#ACESSO
#FILTERS
meta_acesso_np = 0.85
tb_base_acesso = tb_base[tb_base['BACKBONE'].isin(['NINF','BBA','OUTROS'])]
tb_base_acesso = tb_base_acesso[tb_base_acesso['AREA_TECNICA'].isin(['FO'])]
tb_base_acesso = tb_base_acesso[tb_base_acesso['COS'].isin(['PREFO'])]
tb_base_acesso = tb_base_acesso[tb_base_acesso['FTTH'].isin(['N'])]
tb_base_acesso = tb_base_acesso[tb_base_acesso['MONTH'].isin([month])]
vl_acesso = tb_base_acesso['UF'].count()
#TMR
vl_acesso_media_tmr = tb_base_acesso['result_hour'].mean()
#NP
tb_base_acesso_no_prazo = tb_base_acesso[tb_base_acesso['PRAZO_PSR_2016'].isin(['NP'])]
tb_base_acesso_fora_prazo = tb_base_acesso[tb_base_acesso['PRAZO_PSR_2016'].isin(['FP'])]

vl_acesso_np = tb_base_acesso_no_prazo['UF'].count()
vl_acesso_fp = tb_base_acesso_fora_prazo['UF'].count()
acesso_result = vl_acesso_np / vl_acesso

#PROJECAO NP
meta_acesso_dia = meta_acesso_np / last_day_month
meta_acesso_projecao = ((acesso_result / day_now)* last_day_month)
print('projecao_acesso: '+str(meta_acesso_projecao))

#PROJECAO TMR ACESSO
meta_acesso_dia_tmr = hour_meta_acesso / last_day_month
meta_acesso_projecao_tmr = ((vl_acesso_media_tmr / day_now)* last_day_month)
print('projecao_acesso: '+str(meta_acesso_projecao_tmr))

print('VALORES NP ACESSO:')
print('count_acesso '+str(vl_acesso))
print('np_acesso '+str(vl_acesso_np))
print('fp_acesso '+str(vl_acesso_fp))
print('np_result_acesso '+str(acesso_result))
print('mean_acesso_tmr '+str(vl_acesso_media_tmr))


#NO PRAZO FTTH
#PRIMARIO

tb_base_ftth_primario = tb_base[tb_base['STATUS_FTTH'].isin(['PRIMARIO'])]
tb_base_ftth_primario = tb_base_ftth_primario[tb_base_ftth_primario['MONTH'].isin([month])]
tb_base_ftth_primario = tb_base_ftth_primario[tb_base_ftth_primario['FTTH'].isin(['S'])]
vl_ftth_primario = tb_base_ftth_primario['UF'].count()
vl_ftth_primario_media = tb_base_ftth_primario['result_hour'].mean()

tb_base_ftth_primario_no_prazo = tb_base_ftth_primario[tb_base_ftth_primario['PRAZO_PSR_2016'].isin(['NP'])]
tb_base_ftth_primario_fora_prazo = tb_base_ftth_primario[tb_base_ftth_primario['PRAZO_PSR_2016'].isin(['FP'])]

vl_ftth_primario_np = tb_base_ftth_primario_no_prazo['UF'].count()
vl_ftth_primario_fp = tb_base_ftth_primario_fora_prazo['UF'].count()
ftth_primario_result = vl_ftth_primario_np / vl_ftth_primario

print('VALORES NP ftth_PRIMARIO:')
print('count_ftth_primario '+str(vl_ftth_primario))
print('np_ftth_primario '+str(vl_ftth_primario_np))
print('fp_ftth_primario '+str(vl_ftth_primario_fp))
print('np_result_ftth_primario '+str(ftth_primario_result))
print('mean_ftth_primario '+str(vl_ftth_primario_media))


#SECUNDARIO

tb_base_ftth_secundario = tb_base[tb_base['STATUS_FTTH'].isin(['SECUNDARIO'])]
tb_base_ftth_secundario = tb_base_ftth_secundario[tb_base_ftth_secundario['MONTH'].isin([month])]
tb_base_ftth_secundario = tb_base_ftth_secundario[tb_base_ftth_secundario['FTTH'].isin(['S'])]
vl_ftth_secundario = tb_base_ftth_secundario['UF'].count()
vl_ftth_secundario_media = tb_base_ftth_secundario['result_hour'].mean()

tb_base_ftth_secundario_no_prazo = tb_base_ftth_secundario[tb_base_ftth_secundario['PRAZO_PSR_2016'].isin(['NP'])]
tb_base_ftth_secundario_fora_prazo = tb_base_ftth_secundario[tb_base_ftth_secundario['PRAZO_PSR_2016'].isin(['FP'])]

vl_ftth_secundario_np = tb_base_ftth_secundario_no_prazo['UF'].count()
vl_ftth_secundario_fp = tb_base_ftth_secundario_fora_prazo['UF'].count()
ftth_secundario_result = vl_ftth_secundario_np / vl_ftth_secundario

print('VALORES NP ftth_SENCUNDARIO:')
print('count_ftth_secundario '+str(vl_ftth_secundario))
print('np_ftth_secundario '+str(vl_ftth_secundario_np))
print('fp_ftth_secundario '+str(vl_ftth_secundario_fp))
print('np_result_ftth_secundario '+str(ftth_secundario_result))
print('mean_ftth_ecundario '+str(vl_ftth_secundario_media))




dt = pd.DataFrame({

                   #ROMPIMENTO
                   'rompimento':[1000,km,vl_rompimento,result_rompimento],
                   #'np_backbone_month':[],
                   

                   #ATENUACAO
                   'atenuacao':[1000,km,vl_atenuacao,result_atenuacao],
                   #'np_backbone_month':[],
                  

                   #NP BACKCBONE
                   'np_backbone':[vl_backbone,vl_backbone_np,vl_backbone_fp,backbone_result],
                   #'np_backbone_month':[],

                   #TMR BACKBONE
                   'tmr_backbone':['','','',vl_backbone_media_tmr],
                   #'np_ftth_secundario_month':[],
                   

                   #NP ACESSO
                   'np_acesso':[vl_acesso,vl_acesso_np,vl_acesso_fp,acesso_result],
                   #'np_acesso_month':[],


                    #TMR ACESSO
                   'tmr_acesso':['','','',vl_acesso_media_tmr],
                   #'np_ftth_secundario_month':[],
                   
                  
                    #FTTH PRIMARIO NP
                   'np_ftth_primario':[vl_ftth_primario,vl_ftth_primario_np,vl_ftth_primario_fp,ftth_primario_result],
                   #'np_ftth_primario_month':[],

                    #FTTH PRIMARIO TMR
                   'tmr_ftth_primario':['','','',vl_ftth_primario_media],
                   #'np_ftth_primario_month':[],
                  

                    #FTTH SECUNDARIO NP
                   'np_ftth_secundario':[vl_ftth_secundario,vl_ftth_secundario_np,vl_ftth_secundario_fp,ftth_secundario_result],
                   #'np_ftth_secundario_month':[],

                   #FTTH SECUNDARIO TMR
                   'tmr_ftth_secundario':['','','',vl_ftth_secundario_media],
                   #'np_ftth_secundario_month':[],
                   

                    })

dt.to_excel(path+'//NP_indicadores.xlsx')
print('send excel to directory.')


#GENERATING EXCEL FILE
vl_base = tb_base['UF'].count()
print('count_after_filters '+str(vl_base))


tb_base = pd.merge(tb_base,tb_gh,on='NOME_TECNICO',how='left')
tb_base = pd.merge(tb_base,tb_cod_enc,on='COD_ENCERRAMENTO',how='left')
tb_base = pd.merge(tb_base,tb_causa_portal,on='BA',how='left')


tb_base = tb_base.drop_duplicates(subset='BA',keep='first')
vl_after_merge = tb_base['UF'].count()

vl_fim = vl_base
vl_result = vl_fim - vl_after_merge

#tb_base = tb_base[['UF','BA','AREA_TECNICA','COS','COS_ORIGEM','PRIORIDADE','NOME_TECNICO','MATRICULA_TECNICO','BACKBONE','RAMIFICACAO','FTTH','ABERTURA','TEMPO_FIBRA_SEGUNDO','MONTH','YEAR','INI_ACIONAMENTO','ENCE_ACIONAMENTO','EXPURGO','Desc. Coord.','STATUS_AT_ROMP','STATUS_BACKBONE_ACESSO','COD_ENCERRAMENTO','CAUSA3','CAUSA_PORTAL','STATUS_FTTH','kee_status_ftth','result_hour','PRAZO_PSR_2016','tmr_meta']].copy()
tb_base = tb_base[['UF','BA','AREA_TECNICA','COS','COS_ORIGEM','PRIORIDADE','NOME_TECNICO','MATRICULA_TECNICO','BACKBONE','RAMIFICACAO','FTTH','ABERTURA','TEMPO_FIBRA_SEGUNDO','MONTH','YEAR','INI_ACIONAMENTO','ENCE_ACIONAMENTO','EXPURGO','STATUS_AT_ROMP','STATUS_BACKBONE_ACESSO','COD_ENCERRAMENTO','CAUSA3','CAUSA_PORTAL','STATUS_FTTH','kee_status_ftth','result_hour','PRAZO_PSR_2016','tmr_meta','coordenador','gerente_negocios']].copy()


if vl_result == 0:
    print('REAL PROOF OK: '+str(vl_result))
    tb_base.to_excel(path+'\\base_legada_python_model.xlsx', index=False)
    print('Sending file to EXCEL')
    
    #print('Generating the graphics:')
    #graphic = px.treemap(tb_base,path=['STATUS_AT_ROMP','Desc. Coord.'])
    #graphic.show()
    
    #file_telegram = path+'\\bot-legada'
    #os.system(path_telegram+file_telegram)
    #print('Sending file to Telegram')


    
else:
    print('DONT EXPORTED, ERROR IN REAL PROOF.')
    
    
#file_telegram = path+'\\bot-legada'
#os.system(file_telegram)
#print('Sending file to Telegram')



#graphic = px.treemap(tb_base,path=['STATUS_AT_ROMP','NOME_TECNICO'])


   
#else:
#    print("Real proof is wrong!!")

#UPDATE SPREADSHEETS GOOGLE:
#INSTALL:
#pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
#!pip3 install gspread 
# or: pip install -U gspread
#account service in count: python1407:
#https://console.cloud.google.com/apis/api/sheets.googleapis.com/metrics?project=python1407
#PAINEL:
#https://console.cloud.google.com/apis/api/sheets.googleapis.com/metrics?project=python1407
#DOCS:
#https://docs.gspread.org/en/v5.4.0/user-guide.html

import gspread
import pandas as pd
import os

credentials = {
  "type": "service_account",
  "project_id": "python1407",
  "private_key_id": "7be9b74632a590a5c33e8b832654aff1057abc66",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDLZOkEymZuwr4t\nhi38OC05hBTyqWqSqP0cUDBgETlAP459NjyK/DPR4craZvIQWwZN9XvsK/yRj4nB\nKFrMeTxBqTO/m3OZEgklxTNs6xnslDOIFY4E82pZXSZtCRn9JDTnmutX5fWtbClH\nxLsCNeyI0/nF48Y+hGME5M32s0Max4nVwwQbrDUIUIbYeiac+WXzfHOHrZqTj1+X\nYOwKjyHK0wlEYn5Pv5lDR9cvsB8TW9uTQNOY2f7dG6N+y1wnQXvSomVc9R0R8B/U\ntA43lwYJAs97EANLA/Tqy1swRWGoRR8j5DCQbiLY1Y85QyFHKXWLoSRPRuccu6RC\naThG+HKFAgMBAAECggEAJhJ8CcEHPolmhuf8eJ9dW8xNDYVH5S8Lvf6Gp5zhvhSH\njAmYeJ2v54Qf8BTgD86yFeqzKSisrOSU8RqoMGkrLdFJ1f53u3nkS3Un5KX3YtD0\n+m6qeGPGDvdAR52yBy/9VTMrBXeOrsk1yvDY3peQcKZZNUEnLTGjxVk88oZos7y7\nr2Lvtgr2AVUKTOtl94GXE8rYxnJMTtG9uJBo+QUAWX1D5D948WScVZ0VgHt3oIli\nUrtFLHOd+OarGyj6UvJmjHT5TbaaAslqYWMUQHLe7dydrP1DItAdYXZZc2aUuTd2\nukKRvX3NekOQQZcZxnKINdqipyaXgVybtMKrSZHl+wKBgQDopc4nd4rZtA7ojfOi\nshzHIbyTs6L/M+eG60QCGhmzZHqJ1NIV+YHwCQA7wgb+uW9Ke9eKyg0BnF6hrlGo\nGyo19t8v26TrRzsDi08KMm7/4+mcSrnLLT6DdSiyMbZtku3YZHco/eouWR/nlRAw\ngeQe+kbvSC7zWqPvFfwrfnC1IwKBgQDfz2WkSYG8PBRiQ8U8pO5zUndAc75sJYbC\nU5u/6Y1JSTYcDQxkRTjF7VcuzD8FHkWykm1TCfy3YCGqGUMCGynYeyR9qcyoa0zB\nflaCav2HUFtvRsV29YrsxI3N4lAN5BvrBw9fknEHP1ye157/o1037N9zOsy6zK5K\ndOH+eSTYNwKBgBbeRXdnrsRbiKOfYHV7oIyKamjyXXFMftOqSJMUUbZqiAkIXGZA\nkl8v40/8cIeVXrUpmzRPTBv+bObjpa8qjGmljKa9pmZiKBDfHrPX5UVN9+afCchI\n+D4fxBJQBKicqrh8l6H145EOva4b3u2FtxC8dUCMDeFp5XdY5+K2mQmVAoGBANqT\nLezYbP9snWuqTAICAW5W52fmod30eDtoc/9lFDqyaUnT5Ho4sE18kVx+1D0nZ2IS\nZvpmEoz0MWxx52MzLBbjjKu9HMaOpBOEUvBjlN6FuAZg05BuFRNOkj6z+wLV9/38\nkyL/Xat6UfY/FmULIorvpvpePntgUgcdR2jC3xzZAoGAGieZltDVDqX6aKZ83saY\nSgRUYkui8BDhRvzVBVkiGPB6qtOo9jZBMjAOddlnIVeji9/scEDH2Xd4yD2n1Xds\nePsPnBY4bBy3zRdUj1YcV9BP5jT+v/9C2IoGP9YZvLF+D4xiQy/fNERYUrRyDOyp\nPgKM3Un5+PjVMWRuUHCiyow=\n-----END PRIVATE KEY-----\n",
  "client_email": "teste-295@python1407.iam.gserviceaccount.com",
  "client_id": "106801904945565893222",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/teste-295%40python1407.iam.gserviceaccount.com"
   
    
}

url_id = '1cIFUTxA3CohaJZO3c4WjRlMpC0TmOdeoGcAtx8YVXJE'
#url_id = '1EJ_UOMpefTFhW76AdK6chVTF4PsAS_knuOZ8M2uhtNI'
gc = gspread.service_account_from_dict(credentials)
sh = gc.open_by_key(url_id)


#SELECT SHEET FOR NAME:
worksheet_base = sh.worksheet("base")

#CLEAR
#worksheet_base.batch_clear(["A2:B30000", "D2:L30000", "my_named_range"])
worksheet_base.batch_clear(["A2:T40000"])
print('clear the spreadsheet base legada')

path = os.getcwd()
table_sv_base = pd.read_excel(path+'\\base_legada_python_model.xlsx')
print('leitura do excel: '+str(tb_base.shape))
#table_sv_base = tb_base

table_sv_base = table_sv_base.fillna('N LOCALIZADO')

print(table_sv_base.shape)

#table_sv_base = table_sv_base[['UF','BA','PRIORIDADE','NOME_TECNICO','MATRICULA_TECNICO','MONTH','YEAR','Desc. Coord.','STATUS_AT_ROMP','STATUS_BACKBONE_ACESSO','COD_ENCERRAMENTO','CAUSA3','CAUSA_PORTAL','STATUS_FTTH','result_hour','PRAZO_PSR_2016','tmr_meta']].copy()
table_sv_base = table_sv_base[['UF','BA','COS','PRIORIDADE','NOME_TECNICO','MATRICULA_TECNICO','MONTH','YEAR','STATUS_AT_ROMP','STATUS_BACKBONE_ACESSO','COD_ENCERRAMENTO','CAUSA3','CAUSA_PORTAL','STATUS_FTTH','result_hour','PRAZO_PSR_2016','tmr_meta','coordenador','gerente_negocios','FTTH']].copy()

#table_sv_base = table_sv_base[['UF']].copy()
#,'coordenador','gerente_negocios'

#UPDATE SPREADSHEET
worksheet_base.update('a2',table_sv_base.values.tolist())
#worksheet_base.update('a2',tb_base.values.tolist() )

#REAL PROOF
#GETTING CELL VALUE
val_google = worksheet_base.acell('U1').value
val_google = val_google.replace(',','.')

#int(val_google)
val_table_sv_base = table_sv_base['UF'].count()
int(val_table_sv_base)

#val_google = val_google[:-3]
val_google = int(val_google)
val_table_sv_base = int(val_table_sv_base)
print(val_google)
print(val_table_sv_base)

result = val_google - val_table_sv_base
#file_telegram = path+'\\bot-legada'
print(result)
if result == 0:
    print('REAL PROOF: '+str(result))
    #os.system(path_telegram+file_telegram)
    print('Sending file to Telegram')
else:
    print('ERROR IN REAL PROOF:'+str(result))
    
    
    
#PROJECAO
#SELECT SHEET FOR NAME:
worksheet_base = sh.worksheet("legada_projecao")

#CLEAR
worksheet_base.batch_clear(["A2:R10"])
print('clear the spreadsheet legada projecao')

path = os.getcwd()

dt = pd.DataFrame({
                    #ROMPIMENTO A
                   'meta_rompimento':[meta_rompimento],
                   'projecao_rompimento':[meta_rompimento_projecao],
                   'result_rompimento':[result_rompimento],
                                    
                    #ATENUACAO D
                    'meta_atenuacao':[meta_atenuacao],
                    'projecao_atenuacao':[meta_atenuacao_projecao],
                    'result_atenuacao':[result_atenuacao],
                    
                    #BACKBONE NP G
                    'meta_backbone_np':[meta_backbone_np],
                    'projecao_backbone_np':[meta_backbone_projecao],
                    'result_backbone_np':[backbone_result],
                    
                     #BACKBONE TMR J
                    'meta_backbone_tmr':[hour_meta_backbone],
                    'projecao_backbone_tmr':[meta_backbone_projecao_tmr],
                    'result_backbone_tmr':[vl_backbone_media_tmr],
                    
                     #ACESSO NP M
                    'meta_acesso_np':[meta_acesso_np],
                    'projecao_acesso_np':[meta_acesso_projecao],
                    'result_acesso_np':[acesso_result],
                    
                    #ACESSO TMR P
                    'meta_acesso_tmr':[hour_meta_acesso],
                    'projecao_acesso_tmr':[meta_acesso_projecao_tmr],
                    'result_acesso_tmr':[vl_acesso_media_tmr],
                   

                                      
                })

table_sv_base = dt
table_sv_base = table_sv_base[[
                                'meta_rompimento','projecao_rompimento','result_rompimento',
                                'meta_atenuacao','projecao_atenuacao','result_atenuacao',
                                'meta_backbone_np','projecao_backbone_np','result_backbone_np',
                                'meta_backbone_tmr','projecao_backbone_tmr','result_backbone_tmr',
                                'meta_acesso_np','projecao_acesso_np','result_acesso_np',
                                'meta_acesso_tmr','projecao_acesso_tmr','result_acesso_tmr'
                                
                                
                                ]].copy()


#UPDATE SPREADSHEET
worksheet_base.update('a2',table_sv_base.values.tolist())
print('finish proccess projection')

#rompimento(prioridade 97,98,99) e atenuacao(prioridade 21):soma da qtde div km regional ou uf * 1000
#backbone np (colun backbone: bbn, bbr) divide dado da coluna ay: np pelo total
#backbone tmr (colun backbone: bbn, bbr) tempo medio coluna bj:TEMPO_FIBRA_PSR_EM_HORAS
#acesso np (colun backbone: ninf, bba, outros) divide dado da coluna ay: np pelo total
#acesso tmr (colun backbone: ninf, bba, outros) tempo medio coluna bj:TEMPO_FIBRA_PSR_EM_HORAS
#ftth primario col f = fo, col g = prefo,col v = ora ou vazio, col bi = s 
#ftth secundario col f = fo, col g = prfth, col bi = s 

#pyinstaller --onefile fibra_legada.py