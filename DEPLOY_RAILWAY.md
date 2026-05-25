# Deploy do Caloi TPM v2.0 no Railway

## Pré-requisitos
- Conta no GitHub (https://github.com)
- Conta no Railway (https://railway.app) — faça login com o GitHub
- Git instalado no computador (https://git-scm.com/download/win)

---

## PASSO 1 — Criar repositório no GitHub

1. Acesse https://github.com/new
2. Nome do repositório: `caloi-tpm`
3. Marque **Private** (para manter os dados da empresa privados)
4. Clique em **Create repository**

---

## PASSO 2 — Enviar o código para o GitHub

Abra o **Prompt de Comando** dentro da pasta `Manutencao_TPM` e execute:

```bash
git init
git add .
git commit -m "Caloi TPM v2.0 - deploy inicial"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/caloi-tpm.git
git push -u origin main
```

> Substitua `SEU_USUARIO` pelo seu nome de usuário do GitHub.

---

## PASSO 3 — Criar o projeto no Railway

1. Acesse https://railway.app e faça login com o GitHub
2. Clique em **New Project**
3. Selecione **Deploy from GitHub repo**
4. Escolha o repositório `caloi-tpm`
5. Railway vai detectar Python automaticamente. Clique em **Deploy Now**

---

## PASSO 4 — Adicionar volume persistente (IMPORTANTE!)

Sem isso, os dados do Excel são apagados a cada reinicialização.

1. No painel do Railway, clique no seu serviço
2. Vá em **Volumes** (menu lateral)
3. Clique em **Add Volume**
4. Mount path: `/data`
5. Clique em **Save**

---

## PASSO 5 — Configurar variáveis de ambiente

No painel do Railway → seu serviço → aba **Variables**, adicione:

| Variável     | Valor                              |
|--------------|------------------------------------|
| `DATA_DIR`   | `/data`                            |
| `SECRET_KEY` | (gere uma senha aleatória longa)   |
| `PORT`       | (Railway preenche automaticamente) |

Para gerar SECRET_KEY, use: https://passwordsgenerator.net/ (32+ caracteres)

---

## PASSO 6 — Obter a URL pública

1. No Railway → seu serviço → aba **Settings**
2. Em **Networking**, clique em **Generate Domain**
3. Você receberá uma URL tipo: `caloi-tpm-production.up.railway.app`

Compartilhe essa URL com os técnicos — ela funciona em qualquer computador ou celular.

---

## PASSO 7 — Primeiro acesso

1. Acesse a URL gerada
2. Login padrão: **admin** / **admin123**
3. Vá em Configurações → Usuários e troque a senha imediatamente!

---

## Atualizar o sistema (após mudanças no código)

```bash
git add .
git commit -m "Descrição da atualização"
git push
```

O Railway faz o redeploy automaticamente após cada push.

---

## Custos estimados Railway

| Plano    | Custo       | Para este sistema         |
|----------|-------------|---------------------------|
| Hobby    | ~$5/mês     | Suficiente (até 5 usuários)|
| Pro      | ~$20/mês    | Para times maiores         |

O volume de 1 GB de disco está incluído no plano Hobby.

---

## Suporte

- Documentação Railway: https://docs.railway.app
- Em caso de dúvida, abra o chat do Claude e peça ajuda!
