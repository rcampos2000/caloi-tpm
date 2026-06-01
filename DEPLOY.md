# Caloi TPM — Guia de Deploy (GitHub → Railway)

O TPM já está publicado em **https://manutencao-tpm.up.railway.app** e ligado ao repositório
**github.com/rcampos2000/caloi-tpm**. Para que as mudanças (ex.: cadastro de colaboradores)
apareçam online, é preciso **enviar ao GitHub** — o Railway então faz o **redeploy**.

---

## Parte 0 — Tirar do OneDrive (faça uma vez)
O git não funciona bem dentro do OneDrive e há um `.git` acidental na pasta de usuário
(`C:\Users\campo`). Por isso o envio falha ali.

1. Dê dois cliques em **`MOVER_PARA_PROJETOS.bat`** → copia o TPM para `C:\Projetos\Manutencao_TPM`
   (com os dados, para continuar funcionando; o `.gitignore` impede que os dados/token subam).
2. Daqui em diante, rode os `.bat` a partir de **`C:\Projetos\Manutencao_TPM`**.

> Recomendado: deixe **SGM** e **TPM** lado a lado em `C:\Projetos\` — assim o SGM acha os Excel do TPM automaticamente.

---

## Parte 1 — Enviar ao GitHub
(a partir de `C:\Projetos\Manutencao_TPM`)

- Dois cliques em **`SUBIR_GITHUB.bat`**.
  - Na 1ª vez pede a URL: `https://github.com/rcampos2000/caloi-tpm.git`.
  - Tem trava de segurança: se o git apontar para a pasta errada, ele aborta.
- Ou manual:
```
git init
git branch -M main
git remote add origin https://github.com/rcampos2000/caloi-tpm.git
git add -A
git commit -m "Atualizacao TPM"
git push -u origin main
```

> O `.gitignore` já exclui os dados (`*.xlsx`, `config/usuarios.json`, `assinaturas/`, `backups/`)
> e o token (`github_token.txt`, `UPLOAD_GITHUB.py`) — nada sensível vai ao repositório.

---

## Parte 2 — Railway (redeploy)
- Se o projeto Railway do TPM está **conectado ao repo** (Deploy from GitHub), o push **dispara o redeploy** automaticamente. Acompanhe em railway.app → projeto do TPM → Deployments.
- Se não estiver automático: no serviço, **Deploy / Redeploy**.
- Variáveis já configuradas no Railway do TPM (conferir): `DATA_DIR=/data`, `SECRET_KEY`, e o **Volume** montado em `/data` (onde ficam os Excel e o config).

---

## Segurança (pendência)
- `github_token.txt` e `UPLOAD_GITHUB.py` ainda existem na pasta e contêm um token.
  Eles **não vão para o GitHub** (gitignore), mas o ideal é **revogar o token** no GitHub
  (Settings → Developer settings → Personal access tokens) e apagar esses arquivos.

---

## Resumo
1. `MOVER_PARA_PROJETOS.bat` (uma vez) → `C:\Projetos\Manutencao_TPM`.
2. `SUBIR_GITHUB.bat` → envia ao GitHub (caloi-tpm).
3. Railway faz o redeploy → mudanças (cadastro de colaboradores etc.) entram no ar.
