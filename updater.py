import os
import requests
import zipfile
import shutil
import stat

# Diretório atual de execução
LOCAL_REPO_PATH = os.getcwd()
MAIN_PATH = os.path.join(LOCAL_REPO_PATH, "../main")
CURRENT_VERSION_FILE = os.path.join(LOCAL_REPO_PATH, "../VERSAO.txt")

def handle_remove_readonly(func, path, exc_info):
    """Força a remoção de arquivos somente leitura."""
    if func in (os.unlink, os.rmdir):
        os.chmod(path, stat.S_IWRITE)
        func(path)
    else:
        raise

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

def download_and_replace_main(url):
    """Baixa o ZIP da última release e substitui toda a pasta main."""
    zip_path = os.path.join(LOCAL_REPO_PATH, "../update.zip")
    temp_extraction_path = os.path.join(LOCAL_REPO_PATH, "../temp_update")

    # Criar diretório temporário para extração
    os.makedirs(temp_extraction_path, exist_ok=True)

    # Baixar o arquivo ZIP
    print("Baixando a última versão...")
    response = requests.get(url, stream=True)
    if response.status_code == 200:
        with open(zip_path, "wb") as f:
            f.write(response.content)
        print("Download concluído. Extraindo arquivos...")

        # Extrair o ZIP no diretório temporário
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(temp_extraction_path)
        os.remove(zip_path)  # Remover o arquivo ZIP após extração
        print("Arquivos extraídos com sucesso!")

        # Encontrar a pasta que foi extraída
        extracted_folder = next((folder for folder in os.listdir(temp_extraction_path)
                                 if os.path.isdir(os.path.join(temp_extraction_path, folder))), None)
        
        if not extracted_folder:
            print("Erro: pasta principal não encontrada na extração.")
            return

        # Caminho da nova pasta 'main' extraída
        new_main_path = os.path.join(temp_extraction_path, extracted_folder, "../main")

        # Caminho da pasta 'main' original
        original_main_path = os.path.join(LOCAL_REPO_PATH, "../main")
        
        # Renomear a pasta 'main' antiga, se existir
        if os.path.exists(original_main_path):
            backup_path = os.path.join(LOCAL_REPO_PATH, "../main_backup")
            print(f"Renomeando pasta antiga para: {backup_path}")
            if os.path.exists(backup_path):
                shutil.rmtree(backup_path, onerror=handle_remove_readonly)  # Remover backups antigos, se necessário
            input("Movendo a main original para a pasta backup")
            shutil.move(original_main_path, backup_path)
        
        # Mover a nova pasta 'main' para o local correto
        input("Movendo a nova pasta main para o diretorio correto")
        shutil.move(new_main_path, original_main_path)
        print("Pasta 'main' atualizada com sucesso!")

        # Remover o diretório temporário
        shutil.rmtree(temp_extraction_path, onerror=handle_remove_readonly)
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
        download_and_replace_main(zip_url)

        # Atualizar o arquivo VERSAO.txt
        with open(CURRENT_VERSION_FILE, "w") as f:
            f.write(latest_version)
        print("Atualização concluída para a versão", latest_version)
    else:
        print("A automação já está na última versão.")

if __name__ == "__main__":
    main()
