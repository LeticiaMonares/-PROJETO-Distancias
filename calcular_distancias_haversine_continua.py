
import pandas as pd
import numpy as np
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
import time
import os

# Caminho do arquivo original
input_file = "distancias.xlsx"
output_file = "distancias_com_linha_reta.xlsx"

# Coordenadas da empresa (São Caetano do Sul - SP)
origem_lat = -23.4704
origem_lon = -46.5735

# Função de Haversine (linha reta entre dois pontos geográficos)
def haversine(lat1, lon1, lat2, lon2):
    R = 6371  # raio da Terra em km
    phi1 = np.radians(lat1)
    phi2 = np.radians(lat2)
    delta_phi = np.radians(lat2 - lat1)
    delta_lambda = np.radians(lon2 - lon1)

    a = np.sin(delta_phi / 2.0)**2 + np.cos(phi1) * np.cos(phi2) * np.sin(delta_lambda / 2.0)**2
    c = 2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a))
    return R * c

# Inicializar geolocalizador
geolocator = Nominatim(user_agent="geo_dist_calculator")

# Carregar planilha
df = pd.read_excel(input_file)

# Se o arquivo já existir, carregue-o para continuar de onde parou
if os.path.exists(output_file):
    df_saida = pd.read_excel(output_file)
else:
    df["Distância (linha reta km)"] = None
    df_saida = df.copy()

for i, row in df.iterrows():
    cidade = row["Cidade"]
    estado = row["Estado"]

    if pd.notnull(df_saida.at[i, "Distância (linha reta km)"]):
        print(f"[{i+1}] {cidade} - {estado} ... já calculado, pulando.")
        continue

    try:
        location = geolocator.geocode(f"{cidade}, {estado}, Brasil", timeout=10)
        if location:
            distancia = haversine(origem_lat, origem_lon, location.latitude, location.longitude)
            df_saida.at[i, "Distância (linha reta km)"] = round(distancia, 2)
            print(f"[{i+1}] {cidade} - {estado}: {round(distancia, 2)} km")
        else:
            print(f"[{i+1}] {cidade} - {estado}: localização não encontrada.")
    except GeocoderTimedOut:
        print(f"[{i+1}] {cidade} - {estado}: tempo esgotado.")
        continue

    # Salvar progresso após cada cidade
    df_saida.to_excel(output_file, index=False)

    time.sleep(1)

print("✅ Distâncias calculadas e salvas em:", output_file)
