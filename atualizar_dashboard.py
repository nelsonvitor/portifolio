"""
Script para automatizar o fluxo:
1. Abrir a planilha de origem (com Power Query)
2. Atualizar as consultas (equivalente ao Alt+F5)
3. Exportar os dados atualizados para dados.csv
4. Subir o CSV atualizado (e qualquer outra mudança na pasta) para o GitHub (add, commit, push)

Requisitos:
    pip install pywin32

Rode este script no Windows (não funciona em Mac/Linux, pois depende do Excel via COM).
"""

import os
import sys
import time
import subprocess
import traceback
import datetime
import win32com.client as win32

# ===================== CONFIGURAÇÕES =====================
# Caminho completo da planilha de origem (ajuste a extensão se for .xlsm em vez de .xlsx)
CAMINHO_PLANILHA_ORIGEM = r"C:\Users\jaildo.junior\Desktop\DASHBOARD_RFK\DADOS_INSUMO\TESTE_ABC_DASHBOARD_1.xlsx"

# Nome da aba/planilha que contém a tabela final que deve virar o dados.csv
NOME_ABA_TABELA = "BASE DE DADOS"

# Caminho de saída do CSV
CAMINHO_SAIDA_CSV = r"C:\Users\jaildo.junior\Desktop\ANALISE_PCP_RFK\dados.csv"

# Pasta raiz do repositório Git (a pasta que contém a pasta .git)
# >>> CONFIRME se é essa mesma pasta <<<
CAMINHO_REPO_GIT = r"C:\Users\jaildo.junior\Desktop\ANALISE_PCP_RFK"

MENSAGEM_COMMIT = "Atualização automática dos dados"

# Pasta onde os logs de cada execução serão gravados (criada automaticamente se não existir).
# Como o Agendador de Tarefas roda o script sem tela visível, o log é o único jeito
# de conferir depois se a execução das 09h, 10h, 11h etc. funcionou.
PASTA_LOGS = os.path.join(CAMINHO_REPO_GIT, "logs")
# ===========================================================


class Logger:
    """Espelha tudo que é 'print' no console (se existir) e também em um arquivo de log."""

    def __init__(self, caminho_arquivo):
        self.terminal = sys.stdout
        os.makedirs(os.path.dirname(caminho_arquivo), exist_ok=True)
        self.arquivo = open(caminho_arquivo, "a", encoding="utf-8")

    def write(self, mensagem):
        try:
            self.terminal.write(mensagem)
        except Exception:
            pass  # quando rodado sem console (agendador), pode não haver terminal
        self.arquivo.write(mensagem)

    def flush(self):
        try:
            self.terminal.flush()
        except Exception:
            pass
        self.arquivo.flush()


def atualizar_e_exportar():
    # DispatchEx força a criação de uma instância NOVA e isolada do Excel,
    # em vez de reaproveitar uma instância já aberta (que poderia estar visível).
    # Isso não fecha nem interfere em outras planilhas que você já tenha aberto manualmente.
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # Evita o pop-up "Este arquivo contém vínculos para outros arquivos.
    # Deseja atualizá-los?" que trava o script esperando clique manual.
    excel.AskToUpdateLinks = False

    try:
        print("Abrindo planilha de origem...")
        # UpdateLinks=0 -> não atualiza vínculos externos ao abrir (evita o prompt acima)
        wb = excel.Workbooks.Open(CAMINHO_PLANILHA_ORIGEM, UpdateLinks=0)

        # Se o arquivo foi copiado/baixado/sincronizado (OneDrive, rede, e-mail etc.),
        # o Windows pode marcá-lo como "bloqueado" e o Excel abre em Modo de Exibição
        # Protegida (somente leitura), exigindo clique manual em "Habilitar Edição".
        # O bloco abaixo detecta isso e libera a edição automaticamente.
        if excel.ProtectedViewWindows.Count > 0:
            print("Arquivo abriu em Modo de Exibição Protegida. Habilitando edição automaticamente...")
            pvw = excel.ProtectedViewWindows.Item(1)
            wb = pvw.Edit()  # converte a janela protegida em um Workbook editável normal

        if wb.ReadOnly:
            raise RuntimeError(
                "A planilha foi aberta como SOMENTE LEITURA e não foi possível liberar "
                "a edição automaticamente. Verifique se o arquivo não está aberto por "
                "outra pessoa/processo, ou se está marcado como 'Somente leitura' nas "
                "propriedades do arquivo no Windows (botão direito > Propriedades > "
                "desmarcar 'Somente leitura')."
            )

        print("Atualizando Power Query (RefreshAll)...")
        wb.RefreshAll()

        # Consultas do Power Query rodam em segundo plano (assíncronas).
        # Isso força o Excel a esperar até todas terminarem.
        excel.CalculateUntilAsyncQueriesDone()

        # Pequena margem de segurança extra
        time.sleep(3)

        print("Salvando planilha de origem com os dados atualizados...")
        wb.Save()

        print(f"Exportando aba '{NOME_ABA_TABELA}' como CSV...")
        aba = wb.Worksheets(NOME_ABA_TABELA)
        aba.Copy()  # cria um novo workbook temporário só com essa aba
        novo_wb = excel.ActiveWorkbook
        novo_wb.SaveAs(CAMINHO_SAIDA_CSV, FileFormat=62)  # 62 = CSV UTF-8 (preserva acentos como em "MÊS")
        novo_wb.Close(SaveChanges=False)

        wb.Close(SaveChanges=False)  # já foi salva explicitamente acima com wb.Save()
        print(f"CSV salvo com sucesso em: {CAMINHO_SAIDA_CSV}")
    finally:
        excel.Quit()


def subir_para_github():
    print("Enviando alterações para o GitHub...")
    # "-A" adiciona TODAS as mudanças da pasta (não só o dados.csv) — assim,
    # se você editar este script, o index.html, ou qualquer outro arquivo do
    # repositório, ele também é commitado e enviado automaticamente na
    # próxima execução, sem precisar rodar git add/commit/push manualmente.
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
        mensagem = f"Verificação automática (sem mudanças nos dados) - {agora}"
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
    # outros arquivos da pasta, faz o rebase, e devolve essas mudanças depois.
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
    nome_log = datetime.datetime.now().strftime("%Y-%m-%d") + ".log"
    caminho_log = os.path.join(PASTA_LOGS, nome_log)
    sys.stdout = sys.stderr = Logger(caminho_log)

    # Roda sem pausa quando não há um console interativo por trás (ex: Agendador de
    # Tarefas). Quando você dá duplo clique manualmente no .py/.bat, sys.stdin.isatty()
    # é True e a pausa no final continua aparecendo normalmente.
    modo_interativo = sys.stdin is not None and sys.stdin.isatty()

    print(f"\n\n===== Execução iniciada em {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')} =====")
    try:
        atualizar_e_exportar()
        subir_para_github()
        print("============================================")
        print("Execução finalizada com sucesso.")
        print("============================================")
    except Exception:
        print("============================================")
        print("ERRO durante a execução do script:")
        print("============================================")
        traceback.print_exc()
        print("============================================")
    finally:
        if modo_interativo:
            input("Pressione Enter para fechar esta janela...")
