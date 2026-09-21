"""
Script para automatizar o fluxo da BASE DE VENDAS:
1. Abrir a planilha "base_vendas.xlsx" (com Power Query conectado ao dataset do Power BI)
2. Atualizar as consultas (equivalente ao Alt+F5)
3. Exportar os dados atualizados para base_vendas.csv
4. Subir o CSV atualizado para o GitHub (add, commit, push)

Requisitos:
    pip install pywin32

Rode este script no Windows (não funciona em Mac/Linux, pois depende do Excel via COM).
"""

import time
import subprocess
import traceback
import win32com.client as win32

# ===================== CONFIGURAÇÕES =====================
# Caminho completo da planilha de vendas (ajuste a extensão se for .xlsm em vez de .xlsx)
CAMINHO_PLANILHA_VENDAS = r"C:\Users\jaildo.junior\Desktop\DASHBOARD_RFK\DADOS_VENDA\base_vendas.xlsx"

# Nome da aba/planilha que contém a tabela final tratada, vinda do Power BI
NOME_ABA_TABELA = "BASE_VENDAS"

# Caminho de saída do CSV
CAMINHO_SAIDA_CSV = r"C:\Users\jaildo.junior\Desktop\ANALISE_PCP_RFK\base_vendas.csv"

# Pasta raiz do repositório Git (a pasta que contém a pasta .git)
CAMINHO_REPO_GIT = r"C:\Users\jaildo.junior\Desktop\ANALISE_PCP_RFK"

# Nome do arquivo dentro do repositório para o git add (relativo ao CAMINHO_REPO_GIT)
ARQUIVO_NO_REPO = "base_vendas.csv"

MENSAGEM_COMMIT = "Atualização automática da base de vendas"
# ===========================================================


def atualizar_e_exportar():
    # DispatchEx força a criação de uma instância NOVA e isolada do Excel,
    # em vez de reaproveitar uma instância já aberta (não fecha nem interfere
    # em outras planilhas que você já tenha aberto manualmente).
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        print("Abrindo planilha de vendas...")
        wb = excel.Workbooks.Open(CAMINHO_PLANILHA_VENDAS)

        print("Atualizando Power Query (RefreshAll)...")
        wb.RefreshAll()

        # Consultas do Power Query rodam em segundo plano (assíncronas).
        # Isso força o Excel a esperar até todas terminarem.
        excel.CalculateUntilAsyncQueriesDone()

        # Pequena margem de segurança extra
        time.sleep(3)

        print("Salvando planilha de vendas com os dados atualizados...")
        wb.Save()

        print(f"Exportando aba '{NOME_ABA_TABELA}' como CSV...")
        aba = wb.Worksheets(NOME_ABA_TABELA)
        aba.Copy()  # cria um novo workbook temporário só com essa aba
        novo_wb = excel.ActiveWorkbook
        novo_wb.SaveAs(CAMINHO_SAIDA_CSV, FileFormat=62)  # 62 = CSV UTF-8 (preserva acentos)
        novo_wb.Close(SaveChanges=False)

        wb.Close(SaveChanges=False)  # já foi salva explicitamente acima com wb.Save()
        print(f"CSV salvo com sucesso em: {CAMINHO_SAIDA_CSV}")
    finally:
        excel.Quit()


def subir_para_github():
    print("Enviando alterações para o GitHub...")
    # "-A" adiciona TODAS as mudanças da pasta (não só o base_vendas.csv) —
    # assim, se você editar este script, o index.html, ou qualquer outro
    # arquivo do repositório, ele também é commitado e enviado automaticamente
    # na próxima execução, sem precisar rodar git add/commit/push manualmente.
    subprocess.run(["git", "add", "-A"], cwd=CAMINHO_REPO_GIT, check=True)

    agora = time.strftime("%d/%m/%Y %H:%M:%S")

    # Verifica se há alguma mudança pendente (em qualquer arquivo, não só o CSV)
    resultado_status = subprocess.run(
        ["git", "diff", "--cached", "--quiet"],
        cwd=CAMINHO_REPO_GIT,
    )
    houve_mudanca = resultado_status.returncode != 0

    if houve_mudanca:
        mensagem = f"{MENSAGEM_COMMIT} - {agora}"
        subprocess.run(["git", "commit", "-m", mensagem], cwd=CAMINHO_REPO_GIT, check=True)
    else:
        # Cria um commit vazio só para registrar que o script rodou e verificou os dados
        mensagem = f"Verificação automática vendas (sem mudanças) - {agora}"
        subprocess.run(
            ["git", "commit", "--allow-empty", "-m", mensagem],
            cwd=CAMINHO_REPO_GIT,
            check=True,
        )

    # Antes de enviar, sincroniza com o que já existe no GitHub (evita o erro
    # "rejected... fetch first" quando o repositório remoto tem commits que
    # ainda não estão aqui localmente).
    print("Sincronizando com o repositório remoto (git pull --rebase)...")
    # --autostash guarda temporariamente qualquer mudança não commitada em
    # OUTROS arquivos da pasta (ex.: o próprio script .py sendo editado),
    # faz o rebase, e devolve essas mudanças depois. Sem isso, o pull falha
    # com "You have unstaged changes" mesmo sem ter nada a ver com o CSV.
    resultado_pull = subprocess.run(
        ["git", "pull", "--rebase", "--autostash"],
        cwd=CAMINHO_REPO_GIT,
    )

    if resultado_pull.returncode != 0:
        # Provavelmente houve conflito entre as mudanças locais e remotas.
        # Não dá pra resolver isso automaticamente sem risco de perder dados,
        # então abortamos o rebase e avisamos para resolver manualmente.
        subprocess.run(["git", "rebase", "--abort"], cwd=CAMINHO_REPO_GIT)
        raise RuntimeError(
            "Não foi possível sincronizar automaticamente com o GitHub: houve "
            "conflito entre as mudanças locais e as que já estão no repositório "
            "remoto. O rebase foi abortado para não perder nada. Abra o "
            "PowerShell/CMD na pasta do repositório e resolva manualmente com "
            "'git pull' (ou peça ajuda para resolver o conflito) antes de rodar "
            "este script de novo."
        )

    subprocess.run(["git", "push"], cwd=CAMINHO_REPO_GIT, check=True)
    print("Alterações enviadas com sucesso para o GitHub!")


if __name__ == "__main__":
    try:
        atualizar_e_exportar()
        subir_para_github()
        print("\n============================================")
        print("Execução finalizada.")
        print("============================================")
    except Exception:
        print("\n============================================")
        print("ERRO durante a execução do script:")
        print("============================================")
        traceback.print_exc()
        print("============================================")
    finally:
        # Isso garante que a janela SEMPRE fique aberta ao final,
        # mesmo que tenha ocorrido um erro acima.
        input("Pressione Enter para fechar esta janela...")
