import os
import requests
import zipfile

# Diretório atual de execução
LOCAL_REPO_PATH = os.getcwd()
CURRENT_VERSION_FILE = os.path.join(LOCAL_REPO_PATH, "VERSAO.txt")

def get_latest_release():
    """Obtém a última versão publicada no GitHub."""
    url = "https://api.github.com/repos/theoravi/SEIAutomacao/releases/latest"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        return data["tag_name"], data["zipball_url"]
    else:
        print("Erro ao acessar o GitHub:", response.status_code)
        return None, None

def get_current_version():
    """Lê a versão atual do arquivo VERSAO.txt."""
    if os.path.exists(CURRENT_VERSION_FILE):
        with open(CURRENT_VERSION_FILE, "r") as f:
            return f.read().strip()
    return None

def download_and_extract_zip(url, target_path):
    """Baixa o ZIP da última release e substitui os arquivos no diretório local."""
    zip_path = os.path.join(target_path, "update.zip")

    # Baixar o arquivo ZIP
    print("Baixando a última versão...")
    response = requests.get(url, stream=True)
    if response.status_code == 200:
        with open(zip_path, "wb") as f:
            f.write(response.content)
        print("Download concluído. Extraindo arquivos...")

        # Extrair o ZIP
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(target_path)
        os.remove(zip_path)  # Remover o arquivo ZIP após extração
        print("Arquivos extraídos com sucesso!")
    else:
        print("Erro ao baixar a última versão:", response.status_code)

def main():
    print("Verificando atualizações...")

    # Obter informações da última release
    latest_version, zip_url = get_latest_release()
    current_version = get_current_version()

    if not latest_version:
        print("Não foi possível obter a última versão do GitHub.")
        return

    print(f"Versão atual: {current_version if current_version else 'não identificada'}")
    print(f"Última versão: {latest_version}")

    # Verificar se há uma nova versão
    if latest_version != current_version:
        print("Nova versão encontrada! Atualizando...")
        download_and_extract_zip(zip_url, LOCAL_REPO_PATH)

        # Atualizar o arquivo VERSAO.txt
        with open(CURRENT_VERSION_FILE, "w") as f:
            f.write(latest_version)
        print("Atualização concluída para a versão", latest_version)
    else:
        print("A automação já está na última versão.")

if __name__ == "__main__":
    main()
