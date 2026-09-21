"""
Script para monitorar o arquivo "index.html" e, toda vez que ele for alterado
(salvo/colado na pasta), automaticamente:
    1. Esperar alguns segundos para garantir que a gravação do arquivo terminou
    2. git add index.html
    3. git commit -m "..."
    4. git pull --rebase --autostash (sincroniza com o GitHub antes de enviar)
    5. git push

Fica rodando em loop, verificando a cada poucos segundos se a data de
modificação do arquivo mudou. Não precisa instalar nada além do Git já
configurado na pasta (não usa pywin32, não depende do Excel).

Para PARAR: feche a janela ou pressione Ctrl+C.
"""

import os
import time
import subprocess
import traceback
import datetime

# ===================== CONFIGURAÇÕES =====================
# Pasta raiz do repositório Git (a pasta que contém a pasta .git)
CAMINHO_REPO_GIT = r"C:\Users\jaildo.junior\Desktop\ANALISE_PCP_RFK"

# Nome do arquivo a ser monitorado (relativo à pasta acima)
ARQUIVO_MONITORADO = "index.html"

MENSAGEM_COMMIT = "Atualização automática do index.html"

# De quantos em quantos segundos o script verifica se o arquivo mudou
INTERVALO_VERIFICACAO_SEGUNDOS = 5

# Depois de detectar uma mudança, espera esses segundos e confere de novo se
# o arquivo parou de mudar (evita subir um arquivo pela metade caso o editor
# ainda esteja gravando).
TEMPO_ESPERA_ESTABILIZAR_SEGUNDOS = 3
# ===========================================================


def obter_mtime(caminho_arquivo):
    try:
        return os.path.getmtime(caminho_arquivo)
    except FileNotFoundError:
        return None


def subir_para_github():
    caminho_completo = os.path.join(CAMINHO_REPO_GIT, ARQUIVO_MONITORADO)

    print(f"Alteração detectada em {ARQUIVO_MONITORADO}. Enviando para o GitHub...")

    subprocess.run(["git", "add", ARQUIVO_MONITORADO], cwd=CAMINHO_REPO_GIT, check=True)

    # Confere se realmente há mudança no conteúdo (não só na data do arquivo)
    resultado_status = subprocess.run(
        ["git", "diff", "--cached", "--quiet", "--", ARQUIVO_MONITORADO],
        cwd=CAMINHO_REPO_GIT,
    )
    houve_mudanca = resultado_status.returncode != 0

    if not houve_mudanca:
        print("Arquivo foi salvo, mas o conteúdo é igual ao já commitado. Nada a enviar.")
        return

    agora = time.strftime("%d/%m/%Y %H:%M:%S")
    mensagem = f"{MENSAGEM_COMMIT} - {agora}"
    subprocess.run(["git", "commit", "-m", mensagem], cwd=CAMINHO_REPO_GIT, check=True)

    print("Sincronizando com o repositório remoto (git pull --rebase)...")
    resultado_pull = subprocess.run(
        ["git", "pull", "--rebase", "--autostash"],
        cwd=CAMINHO_REPO_GIT,
    )

    if resultado_pull.returncode != 0:
        subprocess.run(["git", "rebase", "--abort"], cwd=CAMINHO_REPO_GIT)
        print(
            "ERRO: não foi possível sincronizar automaticamente com o GitHub "
            "(conflito entre mudanças locais e remotas). O rebase foi abortado. "
            "Resolva manualmente com 'git pull' na pasta do repositório antes "
            "de salvar o arquivo de novo."
        )
        return

    subprocess.run(["git", "push"], cwd=CAMINHO_REPO_GIT, check=True)
    print(f"'{ARQUIVO_MONITORADO}' enviado com sucesso para o GitHub!\n")


def monitorar():
    caminho_completo = os.path.join(CAMINHO_REPO_GIT, ARQUIVO_MONITORADO)

    print("============================================")
    print(f" Monitorando alterações em: {caminho_completo}")
    print(f" Verificando a cada {INTERVALO_VERIFICACAO_SEGUNDOS}s. Pressione Ctrl+C para parar.")
    print("============================================\n")

    ultimo_mtime = obter_mtime(caminho_completo)

    while True:
        time.sleep(INTERVALO_VERIFICACAO_SEGUNDOS)
        mtime_atual = obter_mtime(caminho_completo)

        if mtime_atual is None:
            continue  # arquivo não existe no momento (ex: sendo substituído)

        if mtime_atual != ultimo_mtime:
            # Espera um pouco e confere de novo, para não pegar o arquivo
            # "pela metade" enquanto ainda está sendo salvo/copiado.
            time.sleep(TEMPO_ESPERA_ESTABILIZAR_SEGUNDOS)
            mtime_confirmacao = obter_mtime(caminho_completo)

            if mtime_confirmacao != mtime_atual:
                # Ainda estava mudando, aguarda o próximo ciclo do loop.
                continue

            try:
                subir_para_github()
            except Exception:
                print("ERRO ao tentar enviar para o GitHub:")
                traceback.print_exc()

            ultimo_mtime = mtime_confirmacao


if __name__ == "__main__":
    try:
        monitorar()
    except KeyboardInterrupt:
        print("\nMonitoramento encerrado pelo usuário.")
    except Exception:
        print("\nERRO inesperado:")
        traceback.print_exc()
        input("Pressione Enter para fechar esta janela...")
