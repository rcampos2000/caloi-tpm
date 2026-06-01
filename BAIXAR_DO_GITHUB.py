"""
BAIXAR DO GITHUB → DESKTOP
===========================
Baixa a versão mais recente do repositório GitHub
e atualiza os arquivos locais na pasta Manutencao_TPM.
Use após editar os arquivos diretamente no GitHub ou Codespaces.
"""

import os, sys, json, base64, urllib.request, urllib.error

# ── Configurações ──────────────────────────────────────────────────────────────
OWNER  = "rcampos2000"
REPO   = "caloi-tpm"
BRANCH = "main"
BASE   = os.path.dirname(os.path.abspath(__file__))

# ── Ler token ─────────────────────────────────────────────────────────────────
TOKEN_FILE = os.path.join(BASE, "github_token.txt")
TOKEN = ""
if os.path.exists(TOKEN_FILE):
    with open(TOKEN_FILE) as f:
        TOKEN = f.read().strip()

HEADERS = {
    "Accept": "application/vnd.github.v3+json",
    "User-Agent": "CaloriTPM-Downloader/1.0",
}
if TOKEN:
    HEADERS["Authorization"] = f"token {TOKEN}"

# ── Arquivos que NÃO devem ser sobrescritos (dados locais) ────────────────────
NAO_SOBRESCREVER = {
    "Manutencao_TPM.xlsx",
    "Plano_de_Acao.xlsx",
    "github_token.txt",
    "UPLOAD_GITHUB.py",
    "BAIXAR_DO_GITHUB.py",
    "EXECUTAR_UPLOAD.bat",
    "BAIXAR_DO_GITHUB.bat",
}
NAO_BAIXAR_DIRS = {"config", "assinaturas", "backups", "fotos"}

def api_get(url):
    req = urllib.request.Request(url, headers=HEADERS)
    try:
        resp = urllib.request.urlopen(req, timeout=20)
        return json.loads(resp.read()), resp.status
    except urllib.error.HTTPError as e:
        return None, e.code

def listar_conteudo(caminho=""):
    url = f"https://api.github.com/repos/{OWNER}/{REPO}/contents/{caminho}?ref={BRANCH}"
    dados, status = api_get(url)
    if status != 200 or not dados:
        return []
    return dados

def baixar_arquivo(url_download, destino_local):
    req = urllib.request.Request(url_download, headers=HEADERS)
    resp = urllib.request.urlopen(req, timeout=30)
    conteudo = resp.read()
    os.makedirs(os.path.dirname(destino_local), exist_ok=True)
    with open(destino_local, "wb") as f:
        f.write(conteudo)

def processar_dir(caminho_repo="", caminho_local_base=BASE):
    itens = listar_conteudo(caminho_repo)
    baixados = 0
    pulados  = 0

    for item in itens:
        nome = item["name"]
        tipo = item["type"]

        # Pular pastas de dados locais
        if tipo == "dir" and nome in NAO_BAIXAR_DIRS:
            print(f"  ⏭  Pasta ignorada: {nome}/")
            continue

        caminho_local = os.path.join(caminho_local_base, nome)

        if tipo == "dir":
            sub_repo  = f"{caminho_repo}/{nome}" if caminho_repo else nome
            sub_local = caminho_local
            b, p = processar_dir(sub_repo, sub_local)
            baixados += b
            pulados  += p

        elif tipo == "file":
            if nome in NAO_SOBRESCREVER:
                print(f"  ⏭  Mantido local: {nome}")
                pulados += 1
                continue

            nome_rel = os.path.relpath(caminho_local, BASE).replace("\\", "/")
            print(f"  ⬇  {nome_rel[:60]}", end=" ", flush=True)
            try:
                baixar_arquivo(item["download_url"], caminho_local)
                print("✅")
                baixados += 1
            except Exception as e:
                print(f"❌ {e}")

    return baixados, pulados

# ── EXECUÇÃO ──────────────────────────────────────────────────────────────────
print("=" * 60)
print("  CALOI TPM — BAIXAR VERSÃO MAIS RECENTE DO GITHUB")
print("=" * 60)
print()

# Verificar conexão
print("🔗 Conectando ao GitHub...")
dados, status = api_get(f"https://api.github.com/repos/{OWNER}/{REPO}")
if status != 200 or not dados:
    print(f"❌ Não foi possível conectar. Status: {status}")
    print("   Verifique a conexão com a internet.")
    input("\nPressione Enter para fechar...")
    sys.exit(1)

ultimo_commit_url = f"https://api.github.com/repos/{OWNER}/{REPO}/commits/{BRANCH}"
commit_data, _ = api_get(ultimo_commit_url)
if commit_data:
    msg   = commit_data["commit"]["message"]
    autor = commit_data["commit"]["author"]["name"]
    data  = commit_data["commit"]["author"]["date"][:10]
    print(f"✅ Repositório: {dados['full_name']}")
    print(f"   Último commit: {msg[:50]} ({autor}, {data})")
else:
    print(f"✅ Repositório: {dados['full_name']}")

print()
print("📥 Baixando arquivos...")
print()

baixados, pulados = processar_dir()

print()
print("=" * 60)
print(f"✅ {baixados} arquivo(s) atualizados")
print(f"⏭  {pulados} arquivo(s) mantidos (dados locais)")
print()
print("🚀 Sua pasta local está sincronizada com a nuvem!")
print("=" * 60)
input("\nPressione Enter para fechar...")
