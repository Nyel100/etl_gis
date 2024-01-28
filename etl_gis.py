import geopy
import openpyxl
import pandas as pd
import geopandas as gpd 
from geopandas.tools import geocode 
import psycopg2 
from psycopg2 import sql


#Função para ler arquivo excel e retornar dataframe pandas
def ler_excel(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo)
        return df
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return None
    

#Função para geocodificação
def geocodificar_dataframe(df):
    try:
        dados_geocodificados = geocode(df['ENDEREÇO'], provider = 'nominatim', timeout = None, user_agent = "my_geocoder") 
        df['Latitude'] = dados_geocodificados['geometry'].y
        df['Longitude'] = dados_geocodificados['geometry'].x 
        gdf = gpd.GeoDataFrame(df, geometry = gpd.points_from_xy(df['Longitude'], df['Latitude']))
        return gdf 
    except Exception as e:
        print(f"Erro ao realizar a geocodificação: {e}")
        return None


#chamando o arquivo excel
filename = (r"E:\vscode_projects\etl_gis_project\etl_gis\PontosApoio.xlsx")
apoio_df = ler_excel(filename)

if apoio_df is not None:
    print("Tabela Excel lida com sucesso")
    print(apoio_df.head())
    
    apoio_df = geocodificar_dataframe(apoio_df)
    
    if apoio_df is not None:
        print("Dados geocodificados com sucesso: ")
        print(apoio_df.head())
        
    else:
        print("Falha ao realizar a geocodificação")
    
else:
    print("Falha ao ler o arquivo Excel")
