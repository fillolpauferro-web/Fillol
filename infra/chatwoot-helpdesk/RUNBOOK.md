# Runbook de deploy — Helpdesk Chatwoot (fase 1 / MVP)

Passo a passo para colocar o helpdesk no ar, seguindo a spec aprovada em
`docs/superpowers/specs/2026-07-28-chatwoot-helpdesk-design.md`. Fase 1 cobre
os 3 canais (painel web, e-mail via Google Workspace, chat ao vivo) para até
15 agentes, sem domínio próprio ainda.

## 1. Provisionar o servidor

1. Criar uma VPS (DigitalOcean Droplet 4 vCPU/8GB ou Hetzner equivalente),
   Ubuntu 22.04 LTS.
2. Instalar Docker e o plugin Docker Compose:
   ```bash
   curl -fsSL https://get.docker.com | sh
   apt-get install -y docker-compose-plugin
   ```
3. Configurar o firewall para expor só o necessário:
   ```bash
   ufw allow OpenSSH
   ufw allow 80/tcp
   ufw allow 443/tcp
   ufw enable
   ```

## 2. Copiar o stack para o servidor

Copiar a pasta `infra/chatwoot-helpdesk/` deste repositório para o servidor
(ex: `/opt/chatwoot-helpdesk`), via `scp` ou `git clone`.

## 3. Configurar variáveis de ambiente

```bash
cd /opt/chatwoot-helpdesk
cp .env.example .env
```

Preencher no `.env`:

- `SECRET_KEY_BASE`: gerar com `openssl rand -hex 64`.
- `POSTGRES_PASSWORD` e `REDIS_PASSWORD`: gerar senhas fortes (ex: `openssl rand -hex 24`).
- `REDIS_URL`: colar a mesma senha definida em `REDIS_PASSWORD` no lugar de `__REDIS_PASSWORD__`.
- `FRONTEND_URL`: `http://<IP_PUBLICO_DO_SERVIDOR>` (sem domínio ainda).
- Bloco de e-mail: ver passo 6 abaixo antes de preencher `SMTP_PASSWORD`.

## 4. Subir o stack

```bash
docker compose up -d
docker compose exec rails bundle exec rails db:chatwoot_prepare
```

Isso cria o banco e imprime as credenciais do usuário administrador padrão
no terminal — anotar e trocar a senha no primeiro login.

## 5. Primeiro acesso e configuração dos times

1. Acessar `http://<IP_PUBLICO_DO_SERVIDOR>` e logar como administrador.
2. Em **Settings → Teams**, criar os 3 times: `Vendas`, `Atendimento Externo`,
   `Suporte Interno`.
3. Em **Settings → Inboxes**, criar uma inbox por canal (ver passos 6 e 7) e
   associar cada uma ao time correspondente.

## 6. Conectar o e-mail (Google Workspace)

1. Na conta `suporte@empresa.com` do Google Workspace, ativar a verificação
   em duas etapas e gerar uma **senha de app** em
   `myaccount.google.com/apppasswords`.
2. Colar essa senha em `SMTP_PASSWORD` (e `IMAP` se for usar recebimento via
   IMAP direto) no `.env`, depois `docker compose restart rails sidekiq`.
3. No painel, criar a inbox de e-mail (**Settings → Inboxes → Add Inbox →
   Email**), informando o mesmo endereço e as credenciais SMTP/IMAP.
4. Associar essa inbox ao time correto (`Atendimento Externo` ou o time que
   for dono da caixa `suporte@empresa.com`).

## 7. Embutir o widget de chat no site

1. No painel, criar a inbox de **Website** (**Settings → Inboxes → Add Inbox
   → Website**).
2. Copiar o script de embed gerado e colar antes do `</body>` do site
   institucional (em ambiente de staging primeiro).
3. Associar essa inbox ao time correspondente.

## 8. Cadastrar os agentes

Em **Settings → Agents**, convidar os até 15 agentes por e-mail, definindo
papel (`agent` ou `administrator`) e associando cada um ao(s) time(s) que
vai atender.

## 9. Automatizar backup e healthcheck

1. Preencher no `.env` as credenciais de object storage
   (`BACKUP_S3_ENDPOINT`, `BACKUP_S3_BUCKET`, `BACKUP_S3_ACCESS_KEY`,
   `BACKUP_S3_SECRET_KEY`) e o `HEALTHCHECK_ALERT_EMAIL`.
2. Instalar `aws-cli` e um MTA local (ex: `apt install awscli mailutils`).
3. Adicionar ao crontab do servidor (`crontab -e`):
   ```cron
   0 3 * * * /opt/chatwoot-helpdesk/scripts/backup.sh >> /var/log/chatwoot-backup.log 2>&1
   */5 * * * * /opt/chatwoot-helpdesk/scripts/healthcheck.sh >> /var/log/chatwoot-healthcheck.log 2>&1
   ```

## 10. Validação final

Confirmar cada item antes de considerar a fase 1 concluída (mesmo checklist
da spec aprovada):

- [ ] Painel acessível via IP público, login de administrador funcionando.
- [ ] Times `Vendas`, `Atendimento Externo` e `Suporte Interno` criados.
- [ ] E-mail de teste para `suporte@empresa.com` vira ticket na inbox certa;
      resposta do agente volta por e-mail.
- [ ] Widget de chat no site (staging) recebe mensagem de teste e a resposta
      chega ao visitante.
- [ ] Ticket criado manualmente pelo painel; atribuição e resolução funcionam.
- [ ] `docker compose stop sidekiq && docker compose start sidekiq` simulando
      falha — confirmar que o `restart: unless-stopped` recupera o serviço
      sozinho após um crash real (ex: `docker kill chatwoot-sidekiq`).
- [ ] `./scripts/backup.sh` rodado manualmente uma vez; dump aparece no
      object storage.
- [ ] Os até 15 agentes reais cadastrados, com papéis e times corretos.

## 11. Quando houver domínio próprio

1. Apontar o registro DNS `A` do subdomínio (ex: `suporte.empresa.com`) para
   o IP do servidor.
2. Editar `Caddyfile`: comentar o bloco `:80` e descomentar o bloco de
   domínio.
3. Atualizar `FRONTEND_URL` no `.env` para `https://suporte.empresa.com`.
4. `docker compose restart caddy rails sidekiq`.

Itens fora de escopo desta fase (branding, automações/SLA, base de
conhecimento) ficam para uma fase futura, conforme a spec aprovada.
