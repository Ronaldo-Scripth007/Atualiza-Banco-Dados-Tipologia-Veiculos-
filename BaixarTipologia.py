import os
import re
import csv
import time
import unicodedata
from pathlib import Path
import psycopg2
from psycopg2.extras import execute_values
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# =====================================================================
# XPATHS
# =====================================================================
# Login
XPATH_CAMPO_DOMINIO = "/html/body/form/input[1]"
XPATH_CAMPO_CPF = "/html/body/form/input[2]"
XPATH_CAMPO_USUARIO = "/html/body/form/input[3]"
XPATH_CAMPO_SENHA = "/html/body/form/input[4]"
XPATH_BTN_ENTRAR = "/html/body/form/a"

# Menu principal: campo "Opção"
XPATH_MENU_OPCAO = "//input[@name='f3']"

# Opção 45: campo "Arquivo Excel" (S / N)
XPATH_ARQUIVO_EXCEL = "//*[@id='tp_rel']"

RE_PLACA = re.compile(r"^[A-Z]{3}[0-9][A-Z0-9][0-9]{2}$")  # ABC1234 ou ABC1D23


# =====================================================================
# CONFIGURAÇÕES (.env na mesma pasta deste arquivo)
# =====================================================================
PASTA_BASE = Path(__file__).resolve().parent

load_dotenv(PASTA_BASE / ".env")

SSW_DOMINIO = os.getenv("SSW_DOMINIO", "").strip()
SSW_CPF = os.getenv("SSW_CPF", "").strip()
SSW_USER = os.getenv("SSW_USER", "").strip()
SSW_PASS = os.getenv("SSW_PASS", "").strip()
SUPABASE_DB_URL = os.getenv("SUPABASE_DB_URL", "").strip()

if not SUPABASE_DB_URL:
    raise RuntimeError("Falta SUPABASE_DB_URL no arquivo .env.")

# Pasta onde o Chrome vai salvar o relatório baixado
PASTA_DOWNLOAD = PASTA_BASE / "downloads"
PASTA_DOWNLOAD.mkdir(exist_ok=True)


# =====================================================================
# FUNÇÕES AUXILIARES
# =====================================================================
def preencher(xpath, valor):
    campo = espera.until(EC.presence_of_element_located((By.XPATH, xpath)))
    campo.clear()
    campo.send_keys(valor)

    return campo


def esperar_download(pasta, arquivos_antes, timeout=120):
    """Espera aparecer um arquivo novo na pasta e terminar de baixar."""
    fim = time.time() + timeout

    while time.time() < fim:
        novos = [a for a in pasta.glob("*") if a not in arquivos_antes]

        if novos and not any(a.suffix in (".crdownload", ".tmp") for a in novos):
            return novos[0]

        time.sleep(1)

    raise TimeoutError("O download não terminou a tempo.")


def limpar_relatorio(caminho):
    """Exclui a linha 1 (título) e a coluna A. Salva um arquivo novo e mantém o original."""
    with open(caminho, newline="", encoding="latin-1") as f:
        linhas = list(csv.reader(f, delimiter=";"))

    linhas = linhas[1:]  # exclui a linha 1
    linhas = [linha[1:] for linha in linhas]  # exclui a coluna A

    saida = caminho.with_name(caminho.stem + "_limpo.csv")

    with open(saida, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f, delimiter=";").writerows(linhas)

    return saida


def normalizar_coluna(nome):
    """'DATA/HORA CADASTRO' -> 'data_hora_cadastro'"""
    nome = unicodedata.normalize("NFD", str(nome))
    nome = "".join(c for c in nome if unicodedata.category(c) != "Mn")

    return re.sub(r"[^a-zA-Z0-9]+", "_", nome).strip("_").lower()


RENOMEAR = {
    "modelo_e_n_rastreador_1": "modelo_n_rastreador_1",
    "modelo_e_n_rastreador_2": "modelo_n_rastreador_2",
}


def ajustar_colunas(cabecalho):
    """Normaliza os nomes e numera VALIDADE / NRO AUTORIZAÇÃO pela gerenciadora anterior."""
    colunas = []
    ger = None

    for c in cabecalho:
        n = normalizar_coluna(c)

        if n.startswith("gerenciadora_"):
            ger = n.split("_")[-1]
        elif n in ("validade", "nro_autorizacao") and ger:
            n = f"{n}_gerenciadora_{ger}"

        colunas.append(RENOMEAR.get(n, n))

    return colunas


def enviar_para_supabase(caminho_csv):
    with open(caminho_csv, newline="", encoding="utf-8-sig") as f:
        leitor = csv.reader(f, delimiter=";")
        cabecalho = next(leitor)
        linhas = list(leitor)

    colunas = ajustar_colunas(cabecalho)

    # Descarta colunas sem nome (a vazia do final do arquivo)
    manter = [i for i, c in enumerate(colunas) if c]
    colunas = [colunas[i] for i in manter]
    linhas = [[(l[i] if i < len(l) else "") for i in manter] for l in linhas]

    if "placa" not in colunas:
        raise RuntimeError("Coluna PLACA não encontrada no relatório.")

    # Vazio vira NULL; se a mesma placa aparecer duas vezes, fica a última
    por_placa = {}
    ignoradas = 0
    idx_placa = colunas.index("placa")

    for linha in linhas:
        valores = [(v.strip() or None) for v in linha]
        valores = (valores + [None] * len(colunas))[: len(colunas)]

        if valores[idx_placa] and RE_PLACA.match(valores[idx_placa]):
            por_placa[valores[idx_placa]] = valores
        elif valores[idx_placa]:
            ignoradas += 1  # linha que não é placa (ex.: CNPJ de empresa)

    dados = list(por_placa.values())
    print(f"[info] Placas no relatório: {len(dados)} (linhas ignoradas por não serem placa: {ignoradas})")

    conn = psycopg2.connect(SUPABASE_DB_URL)

    try:
        with conn, conn.cursor() as cur:
            # Confere se todas as colunas do relatório existem na tabela
            cur.execute(
                "select column_name from information_schema.columns "
                "where table_schema = 'public' and table_name = 'VEICULOS'"
            )
            colunas_banco = {r[0] for r in cur.fetchall()} - {"id"}
            desconhecidas = [c for c in colunas if c not in colunas_banco]

            if desconhecidas:
                raise RuntimeError(f"Colunas do relatório que não existem na tabela: {desconhecidas}")

            lista = ", ".join(f'"{c}"' for c in colunas)
            definicao = ", ".join(f'"{c}" text' for c in colunas)

            # Tabela temporária com os dados do relatório
            cur.execute(f"create temp table staging ({definicao}) on commit drop")
            execute_values(cur, f"insert into staging ({lista}) values %s", dados, page_size=1000)

            cur.execute(
                "select count(*) from staging s where not exists "
                '(select 1 from public."VEICULOS" v where v.placa = s.placa)'
            )
            novas = cur.fetchone()[0]

            # Insere as novas e atualiza só as que estão diferentes
            sets = ", ".join(f'"{c}" = excluded."{c}"' for c in colunas if c != "placa")
            antigo = ", ".join(f'v."{c}"' for c in colunas if c != "placa")
            novo = ", ".join(f'excluded."{c}"' for c in colunas if c != "placa")

            cur.execute(
                f'insert into public."VEICULOS" as v ({lista}) select {lista} from staging '
                f"on conflict (placa) do update set {sets} "
                f"where ({antigo}) is distinct from ({novo})"
            )
            afetadas = cur.rowcount

        print(f"[ok] Placas novas inseridas: {novas}")
        print(f"[ok] Placas atualizadas: {afetadas - novas}")

    finally:
        conn.close()


# =====================================================================
# CONFIGURAÇÃO DO NAVEGADOR
# =====================================================================
# region Conf NAVEGADOR
OCULTAR_NAVEGADOR = True  # False = Teste

opcoes = Options()

# if OCULTAR_NAVEGADOR:
#     opcoes.add_argument("--headless=new")
#     opcoes.add_argument("--window-size=1920,1080")
# else:
#     opcoes.add_argument("--start-maximized")

if OCULTAR_NAVEGADOR:
    opcoes.add_argument("--headless=new")
    opcoes.add_argument("--window-size=1920,1080")
    opcoes.add_argument("--disable-gpu")
    opcoes.add_argument("--no-sandbox")
    opcoes.add_argument("--disable-blink-features=AutomationControlled")
    opcoes.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36"
    )
else:
    opcoes.add_argument("--start-maximized")

opcoes.add_experimental_option(
    "prefs",
    {
        "download.default_directory": str(PASTA_DOWNLOAD),
        "download.prompt_for_download": False,
    },
)

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=opcoes)
espera = WebDriverWait(navegador, 15)
# endregion


# =====================================================================
# SSW - LOGIN
# =====================================================================
navegador.get("https://tms.ssw.inf.br/bin/ssw0422")

preencher(XPATH_CAMPO_DOMINIO, SSW_DOMINIO)
preencher(XPATH_CAMPO_CPF, SSW_CPF)
preencher(XPATH_CAMPO_USUARIO, SSW_USER)
preencher(XPATH_CAMPO_SENHA, SSW_PASS)

espera.until(EC.element_to_be_clickable((By.XPATH, XPATH_BTN_ENTRAR))).click()
espera.until(EC.url_contains("menu01"))

print("[ok] Login realizado no SSW.")


# =====================================================================
# SSW - OPÇÃO 45 (muda o foco para a nova janela)
# =====================================================================
janelas_antes = navegador.window_handles[:]

campo_opcao = preencher(XPATH_MENU_OPCAO, "45")
campo_opcao.send_keys(Keys.ENTER)

print("[ok] Opção 45 enviada.")

try:
    espera.until(lambda nav: len(nav.window_handles) > len(janelas_antes))

    nova_janela = [h for h in navegador.window_handles if h not in janelas_antes][-1]
    navegador.switch_to.window(nova_janela)

    print(f"[ok] Foco alterado para a nova janela: {navegador.title}")

except Exception:
    print("[info] A opção 45 abriu na mesma janela.")


# =====================================================================
# SSW - ARQUIVO EXCEL = S (já dispara o download do relatório)
# =====================================================================
arquivos_antes = set(PASTA_DOWNLOAD.glob("*"))

preencher(XPATH_ARQUIVO_EXCEL, "S")

print("[ok] Arquivo Excel = S")

arquivo = esperar_download(PASTA_DOWNLOAD, arquivos_antes)

print(f"[ok] Relatório baixado: {arquivo}")

navegador.quit()

# =====================================================================
# LIMPEZA E ENVIO PARA O SUPABASE
# =====================================================================
arquivo_limpo = limpar_relatorio(arquivo)

print(f"[ok] Relatório limpo salvo em: {arquivo_limpo}")

enviar_para_supabase(arquivo_limpo)



# Apagando os arquivos baixados, na pasta.
for f in (arquivo, arquivo_limpo):
    try:
        f.unlink()
        print(f"[ok] Arquivo apagado: {f.name}")
    except FileNotFoundError:
        pass