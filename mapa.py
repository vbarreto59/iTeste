import pandas as pd
import requests
import time
from fake_useragent import UserAgent

# Função para expandir o link encurtado
def expandir_link(link_encurtado):
    try:
        # Verifica se o link começa com http:// ou https://
        if not link_encurtado.startswith(("http://", "https://")):
            print(f"Link inválido: {link_encurtado}. Pulando...")
            return None
        
        # Cria um User-Agent aleatório para simular um navegador real
        headers = {
            "User-Agent": UserAgent().random
        }
        response = requests.get(link_encurtado, headers=headers, allow_redirects=True, timeout=10)
        return response.url  # Retorna a URL expandida
    except requests.RequestException as e:
        print(f"Erro ao expandir o link {link_encurtado}: {e}")
        return None

# Função para extrair latitude e longitude do link expandido
def extrair_coordenadas(link_expandido):
    try:
        # Verifica se o link contém coordenadas no formato /@lat,long
        if "/@" not in link_expandido:
            print(f"Link não contém coordenadas: {link_expandido}")
            return None, None
        
        # Extrai a parte da URL que contém as coordenadas
        partes = link_expandido.split("/@")
        coordenadas = partes[1].split(",")
        
        # Verifica se há pelo menos duas partes (latitude e longitude)
        if len(coordenadas) < 2:
            print(f"Formato de coordenadas inválido: {link_expandido}")
            return None, None
        
        lat = coordenadas[0]
        long = coordenadas[1]
        return lat, long
    except Exception as e:
        print(f"Erro ao extrair coordenadas do link {link_expandido}: {e}")
        return None, None

# Carrega a planilha da aba "mapa"
caminho_planilha = "mapas.xlsx"
df = pd.read_excel(caminho_planilha, sheet_name="mapa")

# Itera sobre as linhas da planilha
for index, row in df.iterrows():
    link_mapa = row["LinkMapa"]  # Corrigido para "LinkMapa" (com "M" maiúsculo)
    
    # Verifica se o campo LinkMapa está vazio
    if pd.isna(link_mapa) or link_mapa.strip() == "":
        print(f"Linha {index + 1} está em branco. Pulando...")
        continue  # Pula para a próxima linha
    
    # Expande o link encurtado
    link_expandido = expandir_link(link_mapa)
    
    if link_expandido:
        # Extrai as coordenadas
        lat, long = extrair_coordenadas(link_expandido)
        
        if lat and long:
            # Atualiza os campos Latitude e Longitude
            df.at[index, "Latitude"] = lat
            df.at[index, "Longitude"] = long
            print(f"Empreendimento ID {row['Empresa_ID']} atualizado. Latitude: {lat}, Longitude: {long}")
        else:
            print(f"Não foi possível extrair coordenadas para o empreendimento ID {row['Empresa_ID']}.")
    else:
        print(f"Não foi possível expandir o link para o empreendimento ID {row['Empresa_ID']}.")

    # Pausa de 15 segundos entre as requisições
    print("Aguardando 15 segundos antes da próxima requisição...")
    time.sleep(15)

# Salva a planilha atualizada na mesma aba "mapa"
with pd.ExcelWriter(caminho_planilha, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df.to_excel(writer, sheet_name="mapa", index=False)

print("Planilha atualizada e salva com sucesso!")