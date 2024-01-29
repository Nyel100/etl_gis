import os
from dotenv import load_dotenv
import geopy
import openpyxl
import pandas as pd
import geopandas as gpd
from geopandas.tools import geocode
import psycopg2
from psycopg2 import sql

load_dotenv()

def ler_excel(caminho_arquivo):
    
    """
    Lê um arquivo Excel e retorna um DataFrame do Pandas.

    Parâmetros:
    - caminho_arquivo (str): O caminho do arquivo Excel a ser lido.

    Retorna:
    - pandas.DataFrame: Um DataFrame contendo os dados do arquivo Excel.

    Exceções:
    - Retorna None se houver um erro ao ler o arquivo.
    
    """
    
    try:
        df = pd.read_excel(caminho_arquivo)
        return df
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return None

def geocodificar_dataframe(df):
    
    """
    Realiza a geocodificação de um DataFrame utilizando o geocode do GeoPandas.

    Parâmetros:
    - df (pandas.DataFrame): O DataFrame contendo os dados a serem geocodificados.
    
    Retorna:
    - geopandas.GeoDataFrame: Um GeoDataFrame contendo os dados geocodificados.

    Exceções:
    - Retorna None se houver um erro ao realizar a geocodificação.
    
    """
    try:
        dados_geocodificados = geocode(df['ENDEREÇO'], provider='nominatim', timeout=None, user_agent="my_geocoder")
        df['Latitude'] = dados_geocodificados['geometry'].y
        df['Longitude'] = dados_geocodificados['geometry'].x
        gdf = gpd.GeoDataFrame(df, geometry=gpd.points_from_xy(df['Longitude'], df['Latitude']))
        return gdf
    except Exception as e:
        print(f"Erro ao realizar a geocodificação: {e}")
        return None

# Função para conectar ao banco de dados PostgreSQL
def conectar_bd():
    try:
        conn = psycopg2.connect(
            dbname = os.getenv('DATABASE'),
            user = os.getenv('USER'),
            password = os.getenv('PASSWORD'),
            host = os.getenv('HOST'),
            port = os.getenv('PORT')
        )
        return conn
    except Exception as e:
        print(f"Erro ao conectar ao banco de dados: {e}")
        return None

# Função para criar a tabela no banco de dados se ela ainda não existir
def criar_tabela(conn):
    try:
        cursor = conn.cursor()
        tabela_nome = 'pontos_de_apoio'
        colunas = [
            ('id', 'SERIAL PRIMARY KEY'),
            ('comunidades', 'VARCHAR'),
            ('locais_apoio', 'VARCHAR' ),
            ('endereco', 'VARCHAR'),
            ('latitude', 'DOUBLE PRECISION'),
            ('longitude', 'DOUBLE PRECISION'),
            ('geometry', 'GEOMETRY(Point)')
        ]
        sql_criar_tabela = sql.SQL(
            "CREATE TABLE IF NOT EXISTS {} ({})"
        ).format(
            sql.Identifier(tabela_nome),
            sql.SQL(', ').join(
                sql.SQL('{} {}').format(sql.Identifier(nome), sql.SQL(tipo))
                for nome, tipo in colunas
            )
        )
        cursor.execute(sql_criar_tabela)
        conn.commit()
        cursor.close()
        print("Tabela criada com sucesso.")
    except Exception as e:
        print(f"Erro ao criar tabela: {e}")

# Função para inserir os dados geocodificados na tabela do banco de dados
def inserir_dados(conn, gdf):
    try:
        cursor = conn.cursor()
        tabela_nome = 'pontos_de_apoio'
        for index, row in gdf.iterrows():
            comunidades = row['COMUNIDADES']
            endereco = row['ENDEREÇO']
            locais_apoio = row['PONTOS DE APOIO']
            latitude = row['Latitude']
            longitude = row['Longitude']
            geometry = row['geometry']

            # Usando o parâmetro 'geometry' diretamente para PostGIS
            sql_inserir_dados = sql.SQL(
                "INSERT INTO {} (comunidades, locais_apoio, endereco, latitude, longitude, geometry) VALUES (%s, %s, %s, %s, %s, ST_GeomFromText(%s, 4326))"
            ).format(sql.Identifier(tabela_nome))

            cursor.execute(sql_inserir_dados, (comunidades, locais_apoio, endereco, latitude, longitude, f"POINT({longitude} {latitude})"))
            
        conn.commit()
        cursor.close()
        print("Dados inseridos com sucesso.")
    except Exception as e:
        print(f"Erro ao inserir dados: {e}")


def main():
    # Exemplo de uso
    caminho_do_arquivo = r"E:\vscode_projects\etl_gis_project\etl_gis\PontosApoio.xlsx"
    apoio_df = ler_excel(caminho_do_arquivo)

    if apoio_df is not None:
        print("Dados do Excel lidos com sucesso:")
        print(apoio_df.head())

        apoio_gdf = geocodificar_dataframe(apoio_df)

        if apoio_gdf is not None:
            print("Dados geocodificados com sucesso:")
            print(apoio_gdf.head())

            conexao_bd = conectar_bd()

            if conexao_bd is not None:
                criar_tabela(conexao_bd)
                inserir_dados(conexao_bd, apoio_gdf)
                conexao_bd.close()
        else:
            print("Falha ao realizar a geocodificação.")
    else:
        print("Falha ao ler os dados do Excel.")


if __name__ == "__main__":
    main()


