import easygui
import requests
import googlemaps
import pandas as pd
import datetime
import os
import time
import configparser

config = configparser.ConfigParser(interpolation=None)
config.read('Config_getkm.ini')
API_KEY = config['DEFAULT']['api']

SHEET = 0

# Pegando arquivo e definindo estrutura de save
file_table = easygui.fileopenbox()

path = os.path.dirname(file_table)
name_file = os.path.basename(file_table)
name_base = os.path.splitext(name_file)[0]
extensao =(os.path.splitext(name_file)[1]).lower()

# Abrindo a planilha e criando coluna de KM
table = pd.ExcelFile(file_table)
table_read = pd.read_excel(file_table)

# table_read.insert(loc=2,column="KM",value="")    
table_read.insert(loc=6,column="KM",value="")    


#Preparando a estrutura para loop
cidade_origem = table_read.Origem
cidade_destino = table_read.Destino



#Definindo a cidade de Origem
# origem = "LUIS EDUARDO MAGALHAES - BA"


#Inicio do loop para capturar o KM




# gmaps = googlemaps.Client(key=api_key)

# Pegando Lat Long

# geocode_origem = gmaps.geocode(origem)

# lat_origem = geocode_origem[0]['geometry']['location']['lat']

# lng_origem = geocode_origem[0]['geometry']['location']['lng']

# geocode_destino = gmaps.geocode(destino)

# lat_destino = geocode_destino[0]['geometry']['location']['lat']

# lng_destino = geocode_destino[0]['geometry']['location']['lng'] 


#Realizando Consulta e registrando infos.

datainicio = datetime.datetime.now()

name_output = os.path.join(path,f"{name_base} - KM ok {extensao}")
with pd.ExcelWriter(name_output) as writer:
    table_read.to_excel(writer,sheet_name="KM",index=False)


# def escrita_excel(file,sheet,infos):
#     base_output = pd.read_excel(file)
#     n_linhas = len(base_output)
#     # if n_linhas != 0:
#     #     n_linhas = len(base_output) - 1
        
    
#     base_output.loc[n_linhas] = infos
    
#     with pd.ExcelWriter(file) as writer:
#         base_output.to_excel(writer,sheet_name=sheet,index=False)


def escrita_excel(name_output,index,info):
    file = name_output
    out_table = table_read
    # base_output = pd.read_excel(name_output)
    # n_linhas = len(base_output)
    # if n_linhas != 0:
    #     n_linhas = len(base_output) - 1
        
    out_table.loc[index,'KM'] = info
    # base_output.loc[n_linhas] = infos
    
    with pd.ExcelWriter(file) as writer:
        out_table.to_excel(writer,sheet_name='KM',index=False)
        
    # return n_linha

def get_km():
    controlador = 50
    for index,(origem,destino) in enumerate(zip(cidade_origem,cidade_destino)):
        
        try:
            # destino = cidade
            url = "https://maps.googleapis.com/maps/api/distancematrix/json?"
            
            
            r = requests.get(url + "origins="+origem+"&destinations="+destino+"&key="+API_KEY)
            # print(url + "origins="+origem+"&destinations="+destino+"&key="+API_KEY)
            # distancia_valor = r.json()['rows'][0]["elements"][0]["distance"]["value"]
            
            distancia_texto = r.json()['rows'][0]["elements"][0]["distance"]["text"]
            # print(distancia_texto)
            # tempo_valor = r.json()['rows'][0]["elements"][0]["duration"]["value"]
            
            # tempo_texto = r.json()['rows'][0]["elements"][0]["duration"]["text"] 
        
            table_read.loc[index,'KM'] = distancia_texto
            escrita_excel(name_output,index,distancia_texto)
            # infos = table_read.loc[index,'KM'] = distancia_texto
            
            # coluna_index+=1
            # controlador+=1
            
            if index == controlador:
                time.sleep(2)
                controlador = controlador+50
            else:
                pass
            print(f"{destino} - ok")
            
            
        except:
            table_read.loc[index,'KM'] = "nao localizado"
            print (f"{destino} não localizada")
            escrita_excel(name_output,index,"nao localizado")
            # coluna_index+=1
            # controlador+=1
            pass


get_km()


datafim = datetime.datetime.now()
data= (str(datafim.strftime("%y.%m.%d_%H.%M.%S")))

name_output = os.path.join(path,f"{name_base} - KM ok_{data}{extensao}")

with pd.ExcelWriter(name_output) as writer:
    table_read.to_excel(writer,sheet_name="KM",index=False)
    
delta = str(datafim - datainicio)

    
print(f"Tempo de Execuçao{delta} - {name_output}")


# Printando Infos
# print(distancia_valor)

# print(distancia_texto)

# print(tempo_valor)

# print(tempo_texto)

# print("\n")

# print("***Origem***")

# print(f'Latitude: {lat_origem}, Longitude: {lng_origem}')

# print("\n")

# print("***destino***")

# print(f'Latitude: {lat_destino}, Longitude: {lng_destino}')


