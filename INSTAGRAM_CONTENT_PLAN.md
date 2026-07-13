# Plano de conteúdo Instagram — @shonen.cut

## 0. Bloqueios encontrados (leia antes do resto)

- **Algrow (geração de vídeo/TTS e pesquisa de tendência) está com a assinatura pausada** por pagamento pendente/cancelado (erro 402 da API). Não gerei vídeo nem áudio de fato — só o roteiro em texto abaixo. Para gerar o material, reative em `algrow.online/settings/subscription` e peça de novo.
- **Não há conta de Instagram conectada** no Windsor.ai nem no Metricool nesta sessão — só existe `tiktok_organic` (conta "Shonen"). A "pesquisa de tendência" abaixo usa o histórico real do TikTok da mesma conta como proxy, já que é o mesmo criador/nicho/audiência migrando para o Instagram.

## 1. O que os dados reais da conta mostram (Windsor.ai, `tiktok_organic`, fev–jul/2026)

Puxei o histórico completo de posts da conta (`dashboard/posts_full_2026-02-13_a_2026-07-11.json`, 97 posts). O achado principal muda a recomendação do plano anterior (`TIKTOK_VIRAL_PLAN.md`):

**Os posts tipo `PHOTO` (carrossel/slideshow com música) goleiam os `VIDEO` gravado, por uma ordem de grandeza:**

| Post | Tipo | Views | Like % | Shares |
|---|---|---:|---:|---:|
| 22/03 | PHOTO | **693.921** | 9,3% | 4.286 |
| 06/04 | PHOTO | 193.690 | 4,1% | 744 |
| 24/03 | PHOTO | 112.866 | 8,2% | 401 |
| 24/02 | PHOTO | 108.740 | 4,7% | 149 |
| 17/06 | PHOTO | 100.832 | 7,9% | 379 |
| 12/03 | PHOTO | 99.337 | 4,6% | 158 |
| 31/03 | PHOTO | 85.616 | 8,7% | 317 |
| — melhor `VIDEO` da janela (05/04) | VIDEO | 25.174 | 4,8% | 31 |

Ou seja: o "vídeo dos 100K" citado no plano anterior é, tecnicamente, um carrossel de imagens com música (formato "photo mode"), não uma gravação/edição tradicional. Isso é uma vantagem para agora: **dá pra produzir o próximo post sem depender de geração de vídeo** — é um carrossel de cards de texto/imagem, que também é o formato nativo mais forte do Instagram (carrossel > Reels em alcance orgânico para conteúdo de citação).

**Sinal de queda recente:** depois do pico de 17/06 (100.832), os posts de julho caem para a faixa de 300–1.500 views. A conta está "esfriada" — bater a fórmula vencedora de novo é a prioridade, não testar formato novo.

**Fórmula que comprovadamente funciona nesta conta (repetida em todos os picos):**
1. Gancho contrarian na primeira imagem/tela ("frase que quebra expectativa")
2. 4–6 personagens de fandoms grandes e diferentes (cross-fandom → alcança várias bolhas de hashtag)
3. Frase filosófica curta por personagem, tom sério
4. CTA final de comentário/discordância

## 2. Roteiro pronto — carrossel para Instagram (6 slides + capa)

Evita repetir personagens já usados nas partes 1 e 2 do TikTok (Itachi, Luffy, Gojo/Geto, Levi, Tanjiro, Zoro, Gaara, Nezuko, Deku, Killua) — usa fandoms fortes ainda não explorados nesse formato.

**Título de trabalho:** "5 falas de anime que pesam mais que terapia — discorda na order? Comenta 👇"

| Slide | Texto na tela | Personagem / cena sugerida |
|---|---|---|
| 1 (capa) | "ANIME NÃO ENSINA SÓ A LUTAR. ENSINA A SOBREVIVER." | tela preta, fonte bold branca |
| 2 | "GUTS — BERSERK\n'Não é sobre nunca cair. É sobre se levantar uma vez a mais do que a vida te derruba.'" | Guts caminhando ferido, cena de determinação |
| 3 | "EDWARD ELRIC — FULLMETAL ALCHEMIST\n'Trocar algo por igual não existe quando o que você perdeu é insubstituível.'" | Ed/Al cena da transmutação/perda |
| 4 | "SUNG JINWOO — SOLO LEVELING\n'Ninguém vai te salvar. Por isso você aprende a se tornar quem salva.'" | Jinwoo sozinho, cena de evolução |
| 5 | "DENJI — CHAINSAW MAN\n'Ele só queria coisas simples. E o mundo tratou isso como fraqueza.'" | Denji cena emocional/contraste |
| 6 | "REN GOKU — DEMON SLAYER\n'Não é a força que define um herói. É o que ele protege até o fim.'" | Rengoku cena final |
| 7 (fechamento) | "Qual dessas pesou mais pra você? 👇 Se acha que errei a ordem, prova que eu tô errado." | montagem dos 5 nomes numerados |

**Legenda sugerida:**
> Anime não ensina só a lutar. Ensina a sobreviver, perder e continuar mesmo assim. Qual dessas frases pesou mais? Comenta aqui embaixo — e se acha que errei a ordem, prova que eu tô errado. 👇
> #anime #otaku #animeedit #berserk #fullmetalalchemist #sololeveling #chainsawman #demonslayer #reels #shonen

**Se preferir Reels em vez de carrossel:** mesma estrutura, 3–4s por slide (28–30s total), corte seco no beat, sem narração falada (dado que a TTS do Algrow está bloqueada agora) — só texto on-screen e música em alta, igual ao padrão que já vem funcionando na conta.

## 3. Quando o Algrow voltar

Com a assinatura reativada, o próximo passo automático é: `search_viral_videos`/`channel_trends` para confirmar áudio/gancho em alta na semana, depois `generate_video` (ou `generate_image` para os 6 cards do carrossel) e `generate_tts` se decidir narrar. Nada disso foi executado ainda — este documento é só o roteiro.
