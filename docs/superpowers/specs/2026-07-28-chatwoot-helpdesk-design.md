# Design: Helpdesk estilo Zendesk (self-hosted Chatwoot)

**Data:** 2026-07-28
**Status:** Aprovado para planejamento de implementação

## Contexto e objetivo

A empresa precisa de um sistema de atendimento/tickets estilo Zendesk, acessível
de mais de 10 máquinas (agentes acessando via navegador), usado por três frentes:
Vendas/ADM, atendimento a clientes externos e suporte interno.

Em vez de construir um clone do Zendesk do zero (projeto de vários meses),
a decisão estratégica é **adotar e self-hostar o Chatwoot**, uma plataforma
open-source que já cobre tickets, e-mail, chat ao vivo e múltiplos
agentes/times prontos para produção, customizando apenas o necessário
(times, canais, marca) para a empresa.

Esta é a **fase 1 (MVP)** do projeto: deploy da infraestrutura + configuração
inicial dos 3 canais/departamentos. Fases futuras (fora de escopo aqui):
branding customizado, regras de automação/distribuição, SLAs, base de
conhecimento (help center), relatórios avançados, domínio próprio com SSL.

## Escopo confirmado

- **Quem usa:** equipe ADM/Vendas, atendimento a clientes externos, suporte interno — até 15 agentes no total.
- **Canais no MVP:** formulário/painel web, e-mail (Google Workspace), chat ao vivo embutido no site institucional existente.
- **Hospedagem:** nuvem, self-hosted (não é Chatwoot Cloud SaaS).
- **Domínio:** ainda não existe; acesso inicial via IP público do servidor. Documentar migração para domínio próprio quando disponível.

## 1. Arquitetura

- **VPS única** (ex: DigitalOcean Droplet 4 vCPU/8GB ou Hetzner equivalente) — suficiente para até 15 agentes.
- **Docker Compose** com o stack oficial do Chatwoot, sem containers customizados.
- **Caddy** como reverse proxy. SSL automático (Let's Encrypt) fica pendente até haver domínio próprio; até lá, acesso via HTTP no IP do servidor.
- **Acesso dos agentes:** navegador, de qualquer uma das 10+ máquinas da empresa, via `http://<ip-do-servidor>` (depois `https://suporte.empresa.com`).
- **Acesso de clientes externos:** widget de chat embutido no site institucional + formulário público de contato do Chatwoot.

## 2. Componentes

Serviços do Docker Compose (imagens oficiais do Chatwoot):

- **PostgreSQL** — dados principais (tickets, contatos, conversas, usuários).
- **Redis** — filas de background jobs e cache.
- **Rails app (Chatwoot web)** — painel do agente, API, formulário público.
- **Sidekiq (worker)** — processamento assíncrono (e-mails, notificações, webhooks).
- **Caddy** — reverse proxy / SSL.

Configuração inicial via interface admin do Chatwoot (não é código):

- 3 times/inboxes: **Vendas**, **Atendimento Externo**, **Suporte Interno**.
- Inbox de e-mail conectada ao Google Workspace via IMAP/SMTP, usando conta de app dedicada (ex: `suporte@empresa.com` + senha de app do Google).
- Inbox de chat ao vivo (widget) com script de embed gerado para o site institucional.
- Inbox de criação manual de tickets pelo painel.
- Cadastro dos até 15 agentes, com papéis (agente/administrador) e atribuição aos times corretos.

## 3. Fluxo de dados (ciclo de vida de um ticket)

1. **Entrada:** ticket chega por e-mail (`suporte@empresa.com`), pelo widget de chat do site, ou criação manual por um agente.
2. **Roteamento:** o Chatwoot cria a conversa na inbox correspondente automaticamente, conforme o canal de origem.
3. **Atribuição:** um agente do time daquela inbox assume a conversa manualmente (regras automáticas de distribuição ficam para fase futura).
4. **Interação:** agente responde pelo painel; a resposta sai pelo mesmo canal de origem (e-mail ou chat), preservando o histórico.
5. **Resolução:** agente marca a conversa como resolvida; histórico fica associado ao contato para consultas futuras.

## 4. Backups, segurança e resiliência

- **Backup diário automático** do PostgreSQL (dump + upload para object storage — DigitalOcean Spaces ou Backblaze B2), retenção de 7-14 dias.
- **Restart policy `unless-stopped`** em todos os containers.
- **Segredos** (senha do Postgres, credenciais SMTP/IMAP, `SECRET_KEY_BASE`) em `.env` fora do controle de versão, nunca commitado.
- **Firewall:** apenas portas 80/443 e SSH expostas publicamente; Postgres/Redis acessíveis só internamente entre containers.
- **Atualizações:** subir a versão da imagem oficial no compose + rodar migrations; processo documentado, não automatizado nesta fase.
- **Monitoramento básico:** healthcheck periódico do endpoint público do Chatwoot, com alerta por e-mail em caso de indisponibilidade.

## 5. Plano de validação

Checklist manual pós-deploy (este projeto é deploy/configuração, não desenvolvimento de software autoral):

- [ ] Painel do Chatwoot acessível via IP público, login de administrador funcionando.
- [ ] Os 3 times/inboxes (Vendas, Atendimento Externo, Suporte Interno) criados e visíveis.
- [ ] E-mail de teste para `suporte@empresa.com` vira ticket na inbox certa; resposta do agente chega de volta por e-mail.
- [ ] Widget de chat embutido no site (staging) recebe mensagem de teste, aparece no painel, resposta chega ao visitante.
- [ ] Criação manual de ticket pelo painel, fluxo de atribuição e resolução funcionando.
- [ ] Container do Sidekiq derrubado manualmente se recupera sozinho (restart policy).
- [ ] Backup manual executado uma vez, dump aparece no object storage.
- [ ] Os até 15 agentes reais cadastrados, com papéis e times corretos.

## Fora de escopo (fases futuras)

- Domínio próprio com SSL automático.
- Regras de automação/distribuição de tickets.
- SLAs e relatórios avançados.
- Base de conhecimento / help center.
- Branding customizado (cores, logo, domínio do painel).
