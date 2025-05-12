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
api_key = config['Default']['api']



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
table_read.insert(loc=1,column="KM",value="")    


#Preparando a estrutura para loop
cidade_estado = table_read.Cidade_estado
coluna_index = 0
controlador = 0


#Definindo a cidade de Origem
origem = "LUIS EDUARDO MAGALHAES - BA"


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


for cidade in cidade_estado:
    
    try:
        destino = cidade
        url = "https://maps.googleapis.com/maps/api/distancematrix/json?"
        
        r = requests.get(url + "origins="+origem+"&destinations="+destino+"&key="+api_key)
        
        # distancia_valor = r.json()['rows'][0]["elements"][0]["distance"]["value"]
        
        distancia_texto = r.json()['rows'][0]["elements"][0]["distance"]["text"]
        
        # tempo_valor = r.json()['rows'][0]["elements"][0]["duration"]["value"]
        
        # tempo_texto = r.json()['rows'][0]["elements"][0]["duration"]["text"] 
    
        table_read.loc[coluna_index,'KM'] = distancia_texto
        coluna_index+=1
        controlador+=1
        
        if controlador == 50:
            time.sleep(2)
            controlador = 0
        else:
            pass
        print(f"{destino} - ok")
        
    except:
        table_read.loc[coluna_index,'KM'] = "nao localizado"
        print (f"{cidade} não localizada")
        coluna_index+=1
        controlador+=1
        pass



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


