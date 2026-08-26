# -*- coding: utf-8 -*-
"""
Gera o PDF "Hoje, Não Amanhã — 45 caminhos para vencer a procrastinação".
"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, NextPageTemplate,
    PageBreak, Table, TableStyle, KeepTogether, HRFlowable
)
from reportlab.pdfgen import canvas as canvas_mod

PAGE_W, PAGE_H = A4
MARGIN = 27 * mm

# ---------- Paleta ----------
INK = colors.HexColor("#2B2420")       # texto principal, quase preto quente
RUST = colors.HexColor("#C1502E")      # laranja-terracota, cor de destaque
RUST_DARK = colors.HexColor("#9C3D22")
CREAM = colors.HexColor("#FAF3EA")     # fundo claro (capa, divisórias)
SAND = colors.HexColor("#F1E4D3")      # fundo de caixas de ação
GOLD = colors.HexColor("#B98A3E")
GREY = colors.HexColor("#6B6259")

TITLE = "HOJE, NÃO AMANHÃ"
SUBTITLE = "45 caminhos para sair do lugar e vencer a procrastinação"

OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "hoje-nao-amanha.pdf")

# ============================================================
# CONTEÚDO
# ============================================================

INTRO_TEXT = [
    "Você já perdeu horas adiando algo que levaria vinte minutos para terminar. "
    "Já viu um prazo se aproximar e, mesmo sentindo o peso dele, escolheu abrir mais uma aba, "
    "lavar uma louça que podia esperar, ou simplesmente deitar e olhar para o teto. "
    "Depois veio a mesma sensação de sempre: uma mistura de alívio por ter escapado da tarefa "
    "e de culpa por saber que só adiou o problema.",

    "Se isso soa familiar, você não está sozinho — e, mais importante, não está quebrado. "
    "Procrastinar não é um defeito de caráter nem falta de força de vontade. É um mecanismo "
    "de proteção emocional que, em algum momento, aprendeu a te afastar do desconforto de uma "
    "tarefa difícil, chata ou assustadora. O problema é que esse mecanismo, com o tempo, começa "
    "a custar mais caro do que o desconforto que ele evita.",

    "Este livro não promete uma fórmula mágica nem uma força de vontade que você simplesmente "
    "não tem hoje. Ele propõe outra coisa: uma jornada. Quarenta e cinco caminhos, organizados em "
    "seis etapas, que vão do entendimento de por que você adia, passando pela reconstrução da sua "
    "mentalidade, do seu ambiente, do seu planejamento e da sua energia, até chegar num lugar em "
    "que agir deixa de ser um esforço heroico e vira, simplesmente, quem você é.",

    "Você não precisa aplicar as 45 maneiras de uma vez. Aliás, não deveria. Escolha uma página "
    "por dia, ou uma por semana. Releia as que fizerem sentido. Risque as que não servirem para "
    "você agora — talvez sirvam depois. O único compromisso que este livro pede é o mesmo que "
    "toda mudança real pede: começar pequeno, começar hoje.",

    "Antes de virar a página, respire fundo por um instante. Você está prestes a começar uma "
    "jornada de seis atos — não uma corrida, não uma prova de resistência. Cada capítulo constrói "
    "sobre o anterior, e cada uma das 45 maneiras é uma porta, não uma obrigação. Abra as que "
    "fizerem sentido para você agora, no seu ritmo, na sua vida real.",
]

COMO_USAR = [
    "Este livro caminha em seis atos. O primeiro te ajuda a entender por que você procrastina "
    "— e isso, sozinho, já muda a forma como você lida com o problema. Os atos seguintes constroem, "
    "camada por camada, a mentalidade, o ambiente, o planejamento e a energia que sustentam a ação "
    "consistente. O último ato trata do mais difícil: transformar um esforço pontual em identidade.",

    "Cada uma das 45 maneiras termina com uma “Ação de hoje” — um passo pequeno, concreto, "
    "que você pode dar em minutos. Não pule essa parte. É nela que a leitura vira mudança real. "
    "No fim do livro você encontra um plano de ação de 30 dias e um checklist com as 45 maneiras, "
    "para revisitar sempre que precisar de um empurrão.",

    "Não existe ordem errada. Você pode ler do início ao fim, ou pular direto para o ato que mais "
    "conversa com o seu momento atual. O importante é que, ao fechar cada página, você tenha feito "
    "algo — por menor que seja — diferente do que faria se este livro não existisse.",
]

# --- estrutura dos atos ---
# cada tip: (numero, titulo, [paragrafos], acao)

ATO1 = dict(
    numero="Ato 1",
    titulo="Diagnóstico: Entendendo o Inimigo",
    epigrafe="Você não tem um problema de preguiça. Tem um problema de desconforto disfarçado de cansaço.",
    intro="Antes de mudar qualquer hábito, é preciso entender o que realmente está acontecendo quando "
          "você adia. Este primeiro trecho da jornada não pede que você produza mais — pede que você "
          "observe com honestidade. Todo o resto do livro parte daqui.",
    tips=[
        (1, "Procrastinação não é preguiça, é gestão de emoções",
         ["Quando você adia uma tarefa, raramente é porque não sabe como fazê-la ou porque é "
          "incapaz de se esforçar. Na maioria das vezes, você está evitando uma emoção desconfortável "
          "associada a ela: o medo de errar, o tédio, a ansiedade de não saber por onde começar, a "
          "insegurança de não ser bom o suficiente.",
          "Entender isso muda tudo. Você para de brigar consigo mesmo por “ser fraco” e começa a "
          "perguntar a pergunta certa: o que, exatamente, estou evitando sentir agora?"],
         "Antes de adiar a próxima tarefa, pare 10 segundos e nomeie a emoção por trás do adiamento."),

        (2, "Descubra seu tipo de procrastinador",
         ["Nem toda procrastinação nasce do mesmo lugar. O perfeccionista adia porque teme entregar "
          "algo imperfeito. O evitador de risco adia porque, se não tentar, não pode fracassar. O "
          "sonhador adia porque prefere imaginar o resultado a enfrentar o processo. Quem “só funciona "
          "sob pressão” adia porque aprendeu a confundir adrenalina com produtividade. E o ocupado "
          "demais enche a agenda de tarefas pequenas para nunca sobrar tempo para a grande.",
          "Você provavelmente se reconheceu em mais de um. Tudo bem — a maioria das pessoas transita "
          "entre esses perfis dependendo da tarefa e do momento."],
         "Escreva qual desses cinco perfis mais te descreve hoje, e em qual área da sua vida ele aparece mais forte."),

        (3, "Identifique o gatilho emocional por trás de cada adiamento",
         ["Cada tarefa que você evita tem uma história emocional específica. Pode ser um e-mail difícil "
          "que ativa medo de conflito. Pode ser um projeto ambicioso que ativa medo de não ser capaz. "
          "Tratar todo adiamento como se fosse igual é como tomar o mesmo remédio para doenças diferentes.",
          "Da próxima vez que perceber que está adiando algo específico, não tente apenas “se forçar”. "
          "Investigue: o que essa tarefa, especificamente, está tocando em você?"],
         "Escolha uma tarefa adiada agora e complete a frase: “Se eu começar isso, tenho medo de...”"),

        (4, "Mapeie seus horários de fuga",
         ["A procrastinação raramente é aleatória. Ela tem hora marcada. Pode ser sempre depois do "
          "almoço, quando a energia cai. Pode ser no início da manhã, quando a tarefa mais importante "
          "do dia ainda está intocada. Pode ser à noite, quando o cansaço vira desculpa perfeita.",
          "Conhecer o seu padrão é como acender a luz num quarto escuro: você para de tropeçar nos "
          "mesmos móveis sem saber por quê."],
         "Nos próximos três dias, anote o horário exato em que mais sente vontade de adiar algo."),

        (5, "Reconheça a “produtividade falsa”",
         ["Organizar a mesa de trabalho, responder e-mails sem importância, reorganizar a lista de "
          "tarefas pela quinta vez — tudo isso parece produtivo. E é exatamente por isso que engana "
          "tão bem: você termina o dia cansado, com a sensação de ter feito algo, mas a tarefa que "
          "realmente importava continua intocada.",
          "A produtividade falsa é procrastinação de terno e gravata. Ela precisa ser desmascarada "
          "antes de ser combatida."],
         "No fim do dia, pergunte-se: das tarefas que fiz, quais realmente moveram o que importa?"),

        (6, "Faça as pazes com o eu do passado que adiou",
         ["É tentador revisitar cada adiamento antigo com raiva de si mesmo. Mas a autocrítica dura, "
          "ao contrário do que parece, não gera mais disciplina — gera mais vergonha, e a vergonha é "
          "um dos combustíveis favoritos da procrastinação.",
          "Pesquisas sobre autocompaixão mostram, de forma consistente, que pessoas mais gentis "
          "consigo mesmas depois de uma recaída voltam à ação mais rápido do que as que se punem."],
         "Escreva uma frase perdoando você mesmo por um adiamento específico do passado."),

        (7, "Pare de se culpar: a culpa alimenta o ciclo",
         ["Culpa e procrastinação formam um ciclo vicioso: você adia, sente culpa, a culpa gera "
          "desconforto, o desconforto faz você evitar ainda mais a tarefa — para não sentir a culpa "
          "de novo. Quanto mais tempo esse ciclo roda, mais pesada a tarefa parece.",
          "Trocar “eu sou preguiçoso” por “eu adiei essa tarefa, e agora vou dar um passo” parece "
          "sutil, mas é a diferença entre ficar preso e recomeçar."],
         "Da próxima vez que se pegar se culpando, substitua a frase por: “e agora, qual é o menor passo possível?”"),
    ],
)

ATO2 = dict(
    numero="Ato 2",
    titulo="Mentalidade: Mudando a Relação com a Tarefa",
    epigrafe="Você não precisa querer fazer. Precisa apenas começar.",
    intro="Depois de entender por que você adia, a segunda etapa da jornada é reconstruir a forma "
          "como você pensa sobre as tarefas em si. Pequenas mudanças de linguagem e de expectativa "
          "têm um efeito desproporcional sobre a sua disposição para agir.",
    tips=[
        (8, "Troque “eu tenho que” por “eu escolho”",
         ["A linguagem que você usa para falar de uma tarefa muda a forma como seu cérebro reage a "
          "ela. “Eu tenho que” ativa a sensação de obrigação imposta de fora, e o cérebro humano "
          "resiste naturalmente a isso. “Eu escolho” — mesmo quando a tarefa não é opcional na "
          "prática — devolve a sensação de controle.",
          "Não é sobre se enganar. É sobre lembrar que, mesmo dentro de uma obrigação, você está "
          "escolhendo como e quando agir."],
         "Reescreva a próxima tarefa da sua lista trocando “tenho que” por “eu escolho”, em voz alta."),

        (9, "Feito é melhor que perfeito",
         ["O perfeccionismo promete qualidade, mas na prática entrega atraso. Enquanto você espera "
          "a condição ideal, o momento certo, a versão impecável, a tarefa continua parada — e uma "
          "tarefa parada não ajuda ninguém, por mais bem pensada que esteja na sua cabeça.",
          "Uma primeira versão imperfeita, entregue, vale mais do que uma versão perfeita que nunca sai do papel."],
         "Escolha uma tarefa que você vem “polindo” há dias e entregue a versão de hoje, mesmo incompleta."),

        (10, "Divida o monstro: passos de 2 minutos",
         ["Tarefas grandes assustam porque o cérebro as enxerga como um bloco único e intransponível. "
          "A saída é quebrar esse bloco em passos tão pequenos que seja quase impossível dizer não a eles.",
          "“Escrever o relatório” assusta. “Abrir o documento e escrever o título” não. É a mesma "
          "tarefa, mas com uma porta de entrada que qualquer um consegue atravessar."],
         "Pegue a tarefa mais assustadora da sua lista e escreva o menor passo possível de dar em 2 minutos."),

        (11, "Visualize o eu futuro que vai agradecer (ou sofrer)",
         ["Estudos sobre a forma como enxergamos nosso “eu futuro” mostram que, quanto mais distante "
          "e abstrato ele parece, mais fácil é sacrificá-lo por conforto imediato. Tratamos a versão de "
          "nós mesmos daqui a uma semana quase como um estranho.",
          "Aproximar essa imagem — visualizar com detalhes como você vai se sentir amanhã, aliviado ou "
          "sobrecarregado, dependendo da decisão de agora — reduz essa distância e muda a escolha."],
         "Feche os olhos por 20 segundos e visualize como o “você de amanhã” vai se sentir se você agir hoje."),

        (12, "A regra dos 5 minutos para começar qualquer coisa",
         ["Um dos maiores obstáculos não é terminar a tarefa — é começá-la. A regra dos 5 minutos "
          "resolve exatamente isso: comprometa-se a fazer apenas 5 minutos, sabendo que pode parar "
          "depois disso sem culpa.",
          "Na prática, a maior parte da resistência mora na largada. Uma vez em movimento, continuar "
          "custa muito menos energia do que começar."],
         "Escolha a tarefa mais adiada e cronometre 5 minutos de trabalho nela, agora."),

        (13, "Aceite o desconforto inicial como parte do processo",
         ["Todo início de tarefa relevante vem acompanhado de um desconforto — a mente resiste à "
          "mudança de estado, mesmo quando a mudança é boa. Esperar que esse desconforto desapareça "
          "antes de agir é esperar para sempre.",
          "A diferença entre quem age e quem procrastina raramente é a ausência de desconforto. É a "
          "disposição de agir com ele presente."],
         "Nomeie o desconforto que sente ao pensar na sua tarefa pendente e comece mesmo com ele ali."),

        (14, "Disciplina de ação mínima é mais confiável que motivação",
         ["Motivação é uma emoção, e emoções vêm e vão sem aviso. Contar com ela para agir é como "
          "contar com o clima para decidir se você vai trabalhar hoje. A ação mínima consistente — "
          "fazer um pouco, todos os dias, motivado ou não — é o que realmente sustenta um resultado.",
          "Na prática, a motivação costuma aparecer depois que você começa, não antes."],
         "Defina a menor ação possível que você consegue repetir amanhã, mesmo sem vontade nenhuma."),

        (15, "Celebre pequenas vitórias, não só o resultado final",
         ["Quando a única recompensa reconhecida é o resultado final, tarefas longas parecem um "
          "deserto sem água. Reconhecer cada pequeno progresso cria pontos de recompensa ao longo do "
          "caminho, e isso mantém o cérebro engajado.",
          "Terminar um rascunho, dar o primeiro telefonema, abrir a planilha — tudo isso merece ser "
          "notado, não só ignorado a caminho da “vitória de verdade”."],
         "No fim do dia, escreva uma pequena vitória de hoje, mesmo que a tarefa maior ainda não tenha terminado."),
    ],
)

ATO3 = dict(
    numero="Ato 3",
    titulo="Ambiente: Desenhando Seu Espaço Para o Sucesso",
    epigrafe="Seu ambiente decide por você antes que sua vontade tenha qualquer chance.",
    intro="Força de vontade é um recurso limitado e caro. Um ambiente bem desenhado faz o trabalho "
          "pesado por você, reduzindo o número de decisões conscientes necessárias para agir bem.",
    tips=[
        (16, "Elimine gatilhos de distração do campo de visão",
         ["O celular sobre a mesa, a aba do navegador aberta, a série pausada na outra tela — cada "
          "gatilho visível é um convite silencioso para adiar. Você não precisa de mais força de "
          "vontade; precisa de menos tentação diante dos olhos.",
          "Tirar o gatilho do campo de visão custa segundos. Resistir a ele o dia inteiro custa energia "
          "que você não tem de sobra."],
         "Antes de começar sua próxima tarefa importante, retire da vista tudo o que normalmente te distrai."),

        (17, "Crie um espaço sagrado só para trabalho focado",
         ["Quando o mesmo lugar serve para trabalhar, descansar e procrastinar, o cérebro não sabe "
          "qual comportamento esperar dali. Reservar um espaço — mesmo que seja só um canto da mesa "
          "— exclusivamente para foco cria uma associação clara: aqui, eu ajo.",
          "Não precisa ser grande nem perfeito. Precisa ser consistente."],
         "Escolha um local fixo que será, a partir de hoje, o seu espaço de foco — e use-o só para isso."),

        (18, "Use a fricção a seu favor",
         ["Fricção é o esforço extra necessário para fazer algo. Você pode usá-la contra si mesmo, "
          "deixando o que importa fácil e o que distrai difícil: livro aberto na página certa, "
          "aplicativo de distração deslogado, roupa de treino separada na véspera.",
          "Cada segundo de fricção a mais entre você e a distração é um segundo a mais de chance para "
          "a escolha certa vencer."],
         "Escolha uma distração recorrente e adicione um obstáculo real a ela hoje mesmo (senha, logout, distância física)."),

        (19, "Desative notificações e silencie o ruído digital",
         ["Cada notificação é uma interrupção que, além do tempo que consome, exige minutos para você "
          "recuperar o nível de foco que tinha antes dela. Um celular silencioso não é luxo — é "
          "condição básica para qualquer trabalho que exija atenção real.",
          "Você pode responder depois. Raramente algo é tão urgente quanto a notificação faz parecer."],
         "Desative as notificações não essenciais do celular antes de começar a próxima tarefa importante."),

        (20, "Prepare tudo na noite anterior",
         ["Toda decisão pequena pela manhã — o que vestir, por onde começar, onde estão os materiais "
          "— consome uma fatia da sua energia mental disponível para o que realmente importa. Preparar "
          "na noite anterior elimina esse atrito logo na largada do dia.",
          "Quem começa o dia sem precisar decidir nada trivial chega à tarefa importante com mais "
          "energia intacta."],
         "Antes de dormir hoje, separe tudo o que vai precisar para começar amanhã a tarefa mais importante."),

        (21, "Cerque-se de pessoas que te puxam para frente",
         ["Hábitos são contagiosos. Passar tempo perto de pessoas que adiam constantemente normaliza "
          "o adiamento; passar tempo perto de quem age normaliza a ação. Você não escolhe todas as "
          "pessoas ao seu redor, mas escolhe a quem dá mais espaço e atenção.",
          "Às vezes, a mudança de ambiente mais poderosa não é física — é social."],
         "Identifique uma pessoa da sua vida que te inspira a agir, e converse com ela sobre sua meta atual."),

        (22, "Torne o progresso visível",
         ["O que não é visível é fácil de esquecer, inclusive o seu próprio avanço. Um quadro, uma "
          "lista riscada, um calendário marcado — qualquer registro visual do progresso funciona como "
          "lembrete constante de que o esforço está valendo a pena.",
          "Ver o progresso acumulado é um dos combustíveis mais simples e mais subestimados contra a procrastinação."],
         "Crie um registro visual simples (papel, quadro ou app) para acompanhar seu progresso desta semana."),
    ],
)

ATO4 = dict(
    numero="Ato 4",
    titulo="Planejamento: Construindo o Mapa da Ação",
    epigrafe="Um plano simples vale mais do que uma intenção grandiosa.",
    intro="Boas intenções sem estrutura se dissolvem no primeiro imprevisto do dia. Este ato trata de "
          "transformar vontade em um plano simples o suficiente para ser seguido de verdade.",
    tips=[
        (23, "Defina a “tarefa-mãe” do dia",
         ["Quando tudo parece prioridade, nada é. Escolher, antes do dia começar, qual é a única "
          "tarefa que — se fosse a única concluída — já tornaria o dia um sucesso, dá clareza em meio "
          "ao ruído de dezenas de pendências menores.",
          "As outras tarefas continuam existindo. Mas a tarefa-mãe é a que recebe sua melhor energia."],
         "Antes de abrir qualquer outra coisa amanhã, escreva qual será a sua tarefa-mãe do dia."),

        (24, "Time blocking: dê um lar a cada tarefa",
         ["Uma tarefa sem horário reservado na agenda compete o dia inteiro contra tudo mais que "
          "aparece — e geralmente perde. Reservar um bloco específico de tempo, como se fosse um "
          "compromisso inegociável, aumenta muito a chance real de execução.",
          "Não é sobre ter uma agenda rígida. É sobre dar a cada tarefa importante um momento certo "
          "para acontecer."],
         "Abra sua agenda agora e reserve um bloco de tempo específico para a sua tarefa-mãe de amanhã."),

        (25, "Planeje à noite, decida menos pela manhã",
         ["A manhã já chega carregada de decisões — o que vestir, o que comer, por onde começar. "
          "Planejar o dia seguinte na noite anterior tira o planejamento da hora em que você tem "
          "menos energia mental disponível.",
          "Você acorda sabendo exatamente o próximo passo, em vez de gastar a primeira hora do dia "
          "decidindo o que fazer com ela."],
         "Hoje à noite, escreva as três tarefas mais importantes de amanhã, em ordem de prioridade."),

        (26, "Técnica Pomodoro para tarefas pesadas",
         ["Blocos de foco intenso — geralmente 25 minutos — seguidos de uma pausa curta tornam "
          "tarefas grandes psicologicamente suportáveis. O cérebro aceita muito mais fácil “25 minutos” "
          "do que “até terminar”.",
          "O timer também cria um limite claro: você sabe exatamente quando a pausa vai chegar, o que "
          "reduz a tentação de fugir antes disso."],
         "Escolha uma tarefa pesada e trabalhe nela por 25 minutos ininterruptos, com um timer visível."),

        (27, "Prazos artificiais menores que o prazo real",
         ["Prazos distantes reduzem a sensação de urgência, mesmo quando a tarefa é grande. Criar um "
          "prazo pessoal, alguns dias antes do prazo real, devolve urgência a uma tarefa que, "
          "de outra forma, ficaria confortavelmente esquecida até a última hora.",
          "Esse prazo extra também funciona como margem de segurança para imprevistos."],
         "Escolha uma tarefa com prazo distante e defina, agora, uma data pessoal, antecipada, para concluí-la."),

        (28, "Diga não ao multitasking",
         ["Alternar entre tarefas parece produtivo, mas cada troca de foco cobra um custo real: o "
          "cérebro leva tempo para se reorientar completamente à nova tarefa. Fazer várias coisas ao "
          "mesmo tempo, na prática, costuma significar fazer todas elas pior e mais devagar.",
          "Uma coisa de cada vez não é lentidão. É a forma mais rápida de terminar de verdade."],
         "Escolha a próxima hora do seu dia e dedique-a a uma única tarefa, sem alternar entre outras."),

        (29, "Listas de “não fazer” tanto quanto listas de “fazer”",
         ["Uma lista de tarefas diz o que fazer, mas raramente protege o seu tempo das armadilhas "
          "conhecidas: aquele grupo que consome horas, aquele hábito de checar coisas sem necessidade "
          "real. Uma lista de “não fazer” nomeia essas armadilhas com antecedência.",
          "Decidir com clareza o que evitar é tão estratégico quanto decidir o que priorizar."],
         "Escreva três coisas específicas que você não vai fazer amanhã antes de terminar sua tarefa-mãe."),

        (30, "Ritual de revisão semanal",
         ["Sem uma pausa periódica para olhar o quadro geral, é fácil passar semanas reagindo ao "
          "urgente sem nunca avançar no importante. Um ritual fixo — por exemplo, todo domingo — para "
          "revisar o que funcionou e ajustar o plano da próxima semana evita esse piloto automático.",
          "Esse momento não precisa ser longo. Precisa ser constante."],
         "Marque, agora, um horário fixo semanal para revisar seu progresso e planejar a semana seguinte."),
    ],
)

ATO5 = dict(
    numero="Ato 5",
    titulo="Foco e Energia: Sustentando o Ritmo",
    epigrafe="Disciplina sem energia é uma vela tentando queimar sem pavio.",
    intro="Nenhuma técnica de planejamento resiste a um corpo e uma mente esgotados. Este ato cuida "
          "da base física e mental que sustenta tudo o que veio antes.",
    tips=[
        (31, "Proteja seu horário de pico de energia",
         ["Todo mundo tem um período do dia em que a mente está mais alerta e disponível. Gastar "
          "esse período com tarefas triviais — e deixar a tarefa mais importante para quando a "
          "energia já caiu — é desperdiçar o seu melhor recurso na hora errada.",
          "Proteger esse horário para o que mais importa é uma das decisões de maior retorno que existem."],
         "Identifique seu horário de maior energia hoje e reserve-o, a partir de amanhã, para sua tarefa-mãe."),

        (32, "Sono e descanso como parte da produtividade, não contra ela",
         ["Cortar sono para “ganhar tempo” costuma produzir o efeito contrário: menos foco, mais "
          "erros, mais vontade de adiar no dia seguinte. Descanso não é o oposto de produtividade — é "
          "a base sobre a qual ela se sustenta.",
          "Tratar o sono como prioridade, não como sobra do dia, muda a qualidade de tudo o que vem depois."],
         "Defina um horário para dormir hoje que te garanta pelo menos 7 horas de descanso."),

        (33, "Pausas estratégicas contra o esgotamento",
         ["Trabalhar sem pausas parece disciplina, mas o esgotamento que isso gera é um dos maiores "
          "combustíveis da procrastinação futura. Pausas curtas e regulares evitam que a energia "
          "despenque a ponto de qualquer tarefa parecer impossível.",
          "Descansar no meio do caminho não é fraqueza. É manutenção."],
         "Programe uma pausa de 5 minutos a cada 50 minutos de trabalho hoje, sem exceção."),

        (34, "Alimentação pensada para energia mental",
         ["O que você come afeta diretamente sua clareza mental e disposição para agir. Refeições "
          "pesadas ou muito açucaradas antes de um período de foco costumam cobrar o preço em sonolência "
          "e dificuldade de concentração logo depois.",
          "Não é sobre dieta perfeita. É sobre notar o que sustenta sua energia e o que a derruba."],
         "Antes da sua próxima tarefa importante, escolha uma refeição ou lanche leve que sustente energia, não sono."),

        (35, "Movimente o corpo para destravar a mente",
         ["Quando a mente trava, muitas vezes o corpo é a saída mais rápida: uma caminhada curta, "
          "alongamento, alguns minutos de movimento. A circulação melhora, a tensão diminui, e a "
          "clareza costuma voltar mais rápido do que ficar parado tentando “se forçar” a pensar.",
          "Mover o corpo é, muitas vezes, o atalho mais direto para destravar a cabeça."],
         "Da próxima vez que travar numa tarefa, levante e caminhe por 5 minutos antes de voltar a ela."),

        (36, "Pratique respiração ou mindfulness antes de tarefas difíceis",
         ["Poucos minutos de respiração consciente antes de encarar uma tarefa desconfortável reduzem "
          "a ativação de ansiedade que normalmente empurra você para a fuga. É um passo pequeno que "
          "prepara o terreno emocional antes de começar.",
          "Não é sobre eliminar o desconforto. É sobre entrar nele com mais calma."],
         "Antes da sua próxima tarefa difícil, respire fundo por um minuto, contando a respiração devagar."),

        (37, "Elimine a fadiga de decisão",
         ["Cada decisão pequena do dia — o que vestir, o que comer, qual tarefa fazer primeiro — "
          "consome uma reserva limitada de energia mental. Quando essa reserva se esgota, mesmo "
          "decisões simples parecem pesadas, e adiar vira o caminho de menor resistência.",
          "Reduzir decisões triviais, com rotinas e escolhas pré-definidas, poupa energia para o que "
          "realmente exige julgamento."],
         "Escolha uma decisão trivial do seu dia (roupa, refeição, rota) e transforme-a numa rotina fixa esta semana."),

        (38, "Reconecte-se com o porquê quando a energia cair",
         ["Nos dias em que a energia está baixa, técnicas e listas ajudam menos do que lembrar, com "
          "clareza, por que aquela tarefa importa para você. O motivo verdadeiro é o combustível que "
          "resta quando a disposição já foi embora.",
          "Voltar ao porquê não resolve o cansaço, mas dá sentido suficiente para dar o próximo passo mesmo assim."],
         "Escreva, em uma frase, por que a sua tarefa-mãe de hoje realmente importa para você."),
    ],
)

ATO6 = dict(
    numero="Ato 6",
    titulo="Consolidação: Um Estilo de Vida, Não um Esforço",
    epigrafe="Não estamos buscando um dia perfeito. Estamos construindo uma identidade.",
    intro="A última etapa da jornada é a mais silenciosa e a mais decisiva: transformar tudo o que "
          "você praticou até aqui em algo que não exige mais esforço consciente — porque virou parte "
          "de quem você é.",
    tips=[
        (39, "Recompense a conclusão, não o esforço",
         ["Recompensar apenas o esforço pode, sem querer, reforçar o hábito de ficar girando em torno "
          "de uma tarefa sem nunca fechá-la. Reservar uma recompensa específica para quando a tarefa "
          "realmente termina ensina o cérebro a valorizar o fechamento, não só o movimento.",
          "Pequenas recompensas, escolhidas por você, tornam o hábito de terminar mais forte com o tempo."],
         "Defina uma pequena recompensa para quando concluir sua tarefa-mãe hoje — e só a use depois de terminar."),

        (40, "Tenha um parceiro de responsabilidade",
         ["Compromissos compartilhados com outra pessoa são mais difíceis de adiar do que compromissos "
          "só com você mesmo. Contar a alguém sua meta, e combinar de reportar o progresso, adiciona "
          "uma camada extra de motivação que a solidão da tarefa não oferece.",
          "Não precisa ser formal. Só precisa ser real."],
         "Escolha uma pessoa e conte a ela sua meta atual, combinando um dia para reportar o progresso."),

        (41, "Comemore sua sequência de dias consistentes",
         ["Existe um efeito motivador simples em ver dias consecutivos de ação acumulados: cada dia "
          "somado à sequência aumenta o custo emocional de quebrá-la. Isso não é sobre perfeição — é "
          "sobre reconhecer o valor da consistência visível.",
          "Uma sequência quebrada não apaga os dias anteriores. Ela só pede que você comece uma nova, hoje."],
         "Marque, num calendário ou app, cada dia em que você agiu conforme o plano — e observe a sequência crescer."),

        (42, "Perdoe recaídas e recomece sem drama",
         ["Toda jornada de mudança tem dias de recaída. O erro mais caro não é procrastinar de novo "
          "— é transformar essa recaída em prova de que “nada funciona” e desistir de vez. Um dia "
          "perdido é só um dia perdido, não um veredito sobre quem você é.",
          "Quem sustenta mudanças de verdade não é quem nunca falha. É quem volta rápido depois de falhar."],
         "Se você adiou algo hoje, não analise o motivo agora — apenas escolha o próximo pequeno passo e dê-o."),

        (43, "Ensine o que aprendeu",
         ["Explicar uma ideia para outra pessoa obriga você a organizá-la com mais clareza do que "
          "quando ela só existe na sua cabeça. Ensinar o que você aprendeu sobre vencer a procrastinação "
          "fixa o aprendizado de um jeito que só ler ou só praticar sozinho raramente consegue.",
          "Você não precisa ser especialista para compartilhar o que está funcionando para você."],
         "Compartilhe com alguém, hoje, uma das 45 maneiras deste livro que já fez diferença para você."),

        (44, "Escreva sua carta ao procrastinador",
         ["Nos dias difíceis, a mente esquece rápido os motivos que a trouxeram até aqui. Escrever "
          "uma carta para o “você que vai querer desistir” — com os motivos, os pequenos avanços, o "
          "porquê que importa — cria um recurso pronto para os momentos em que a vontade de adiar volta com força.",
          "Guarde essa carta em um lugar de fácil acesso. Ela é para os dias em que este livro estiver longe."],
         "Escreva uma carta curta para o seu “eu que vai querer procrastinar” amanhã, e guarde-a por perto."),

        (45, "Redefina quem você é: “sou alguém que age”",
         ["No fim, a mudança mais duradoura não é uma técnica isolada, mas a identidade que você "
          "constrói a partir da repetição de pequenas ações. Cada vez que você escolhe agir em vez de "
          "adiar, você não está apenas terminando uma tarefa — está votando em quem quer ser.",
          "Você não precisa esperar se tornar “alguém que age” para começar a agir. É o contrário: "
          "agir, mesmo pequeno, mesmo imperfeito, é como essa identidade se constrói."],
         "Termine hoje completando uma única ação pendente, e diga em voz alta: “eu sou alguém que age”."),
    ],
)

ATOS = [ATO1, ATO2, ATO3, ATO4, ATO5, ATO6]

CONCLUSAO = [
    "Você chegou ao fim de 45 caminhos, mas não ao fim da jornada. Vencer a procrastinação não é um "
    "destino que se alcança de uma vez — é uma prática que se renova a cada manhã, a cada tarefa, "
    "a cada pequena escolha entre adiar e agir.",

    "Releia este livro quantas vezes precisar. Nos dias difíceis, volte ao Ato 1 e lembre-se: você "
    "não está quebrado, só está evitando um desconforto — e agora sabe como atravessá-lo. Nos dias "
    "de energia baixa, volte ao Ato 5. Nos dias em que a identidade parecer distante, volte ao Ato 6.",

    "O que separa quem vive procrastinando de quem age não é a ausência de dúvida, medo ou cansaço. "
    "É a disposição de dar o próximo pequeno passo mesmo com tudo isso presente. Você já deu vários "
    "passos ao longo destas páginas. O próximo é seu, agora.",

    "Guarde este livro em um lugar de fácil acesso — não numa prateleira distante, mas em algum "
    "canto do seu dia a dia. Ele foi escrito para ser folheado de novo, sublinhado, dobrado nas "
    "pontas. Um guia de caminhos só cumpre seu papel quando é revisitado no momento em que o "
    "adiamento bate à porta outra vez.",
]

PLANO_30_DIAS = [
    ("Semana 1 — Diagnóstico e mentalidade",
     "Releia o Ato 1 e o Ato 2. Escolha três maneiras para praticar todos os dias: nomear a emoção "
     "por trás do adiamento, aplicar a regra dos 5 minutos e trocar “tenho que” por “eu escolho”."),
    ("Semana 2 — Ambiente e planejamento",
     "Redesenhe seu espaço de foco (Ato 3) e comece a definir sua tarefa-mãe diária, com time "
     "blocking e planejamento noturno (Ato 4). Elimine ao menos três gatilhos de distração."),
    ("Semana 3 — Energia e consistência",
     "Proteja seu horário de pico de energia, priorize o sono e use pausas estratégicas (Ato 5). "
     "Comece a marcar sua sequência de dias consistentes."),
    ("Semana 4 — Consolidação",
     "Encontre um parceiro de responsabilidade, escreva sua carta ao procrastinador e revise as 45 "
     "maneiras (Ato 6). Escolha as cinco que mais funcionaram para você e torne-as parte da sua rotina."),
]

CHECKLIST_TITULOS = []
for ato in ATOS:
    for (num, titulo, _, _) in ato["tips"]:
        CHECKLIST_TITULOS.append((num, titulo))

# ============================================================
# ESTILOS
# ============================================================
styles = getSampleStyleSheet()

def style(name, **kw):
    base = dict(fontName="Helvetica", textColor=INK, leading=15)
    base.update(kw)
    return ParagraphStyle(name, **base)

S_BODY = style("body", fontSize=13.6, leading=22.2, alignment=TA_JUSTIFY, spaceAfter=15)
S_BODY_FIRST = style("body_first", fontSize=13.6, leading=22.2, alignment=TA_JUSTIFY, spaceAfter=15)
S_ACTION = style("action", fontName="Helvetica-Oblique", fontSize=10.5, leading=15,
                  textColor=RUST_DARK, alignment=TA_LEFT)
S_ACTION_LABEL = style("action_label", fontName="Helvetica-Bold", fontSize=10.5,
                        textColor=RUST_DARK)
S_TIP_TITLE = style("tip_title", fontName="Helvetica-Bold", fontSize=17, leading=21,
                     textColor=INK, spaceBefore=6, spaceAfter=12)
S_TIP_NUM = style("tip_num", fontName="Helvetica-Bold", fontSize=10.5, leading=13,
                   textColor=RUST)
S_ATO_INTRO = style("ato_intro", fontName="Helvetica-Oblique", fontSize=10.6, leading=16,
                     textColor=GREY, alignment=TA_JUSTIFY, spaceAfter=6)
S_SUMARIO_H = style("sumario_h", fontName="Helvetica-Bold", fontSize=12.5, leading=16,
                     textColor=RUST_DARK, spaceBefore=10, spaceAfter=4)
S_SUMARIO_ITEM = style("sumario_item", fontName="Helvetica", fontSize=9.6, leading=14,
                        textColor=INK)
S_H1 = style("h1", fontName="Helvetica-Bold", fontSize=17, leading=21, textColor=RUST_DARK,
             spaceAfter=10)
S_CHECK_ATO = style("check_ato", fontName="Helvetica-Bold", fontSize=11, leading=15,
                     textColor=RUST_DARK, spaceBefore=8, spaceAfter=3)
S_CHECK_ITEM = style("check_item", fontName="Helvetica", fontSize=9.4, leading=13.5,
                      textColor=INK)
S_WEEK_TITLE = style("week_title", fontName="Helvetica-Bold", fontSize=12, leading=15,
                      textColor=RUST_DARK, spaceAfter=3)
S_WEEK_BODY = style("week_body", fontName="Helvetica", fontSize=10, leading=14.5,
                     alignment=TA_JUSTIFY, spaceAfter=10)

def rust_rule(width="100%", thickness=1.4, color=RUST, space_before=0, space_after=10):
    return HRFlowable(width=width, thickness=thickness, color=color,
                       spaceBefore=space_before, spaceAfter=space_after, lineCap='round')

# ============================================================
# CALLBACKS DE PAGINA (capa / divisorias de ato / paginas normais)
# ============================================================
# guarda o titulo do ato corrente, atualizado por um flowable invisivel
_running = {"title": "Introdução"}


def set_running_title(title):
    _running["title"] = title


class SetRunningTitle(Spacer):
    """Flowable de altura zero que atualiza o titulo corrente no cabecalho."""
    def __init__(self, title):
        Spacer.__init__(self, 0, 0)
        self.title = title

    def draw(self):
        set_running_title(self.title)


def draw_cover(c, doc):
    c.saveState()
    c.setFillColor(CREAM)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    c.setFillColor(RUST)
    c.rect(0, PAGE_H - 14 * mm, PAGE_W, 14 * mm, fill=1, stroke=0)
    c.rect(0, 0, PAGE_W, 10 * mm, fill=1, stroke=0)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 34)
    c.drawCentredString(PAGE_W / 2, PAGE_H / 2 + 26 * mm, "HOJE,")
    c.drawCentredString(PAGE_W / 2, PAGE_H / 2 + 12 * mm, "NÃO AMANHÃ")
    c.setStrokeColor(RUST)
    c.setLineWidth(1.6)
    c.line(PAGE_W / 2 - 30 * mm, PAGE_H / 2 + 4 * mm, PAGE_W / 2 + 30 * mm, PAGE_H / 2 + 4 * mm)
    c.setFont("Helvetica", 13)
    c.setFillColor(GREY)
    c.drawCentredString(PAGE_W / 2, PAGE_H / 2 - 6 * mm, "45 caminhos para sair do lugar")
    c.drawCentredString(PAGE_W / 2, PAGE_H / 2 - 13 * mm, "e vencer a procrastinação")
    c.setFont("Helvetica-Bold", 10.5)
    c.setFillColor(RUST_DARK)
    c.drawCentredString(PAGE_W / 2, 22 * mm, "UM GUIA EM SEIS ATOS")
    c.restoreState()


def draw_ato(c, doc):
    c.saveState()
    c.setFillColor(CREAM)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    c.setFillColor(RUST)
    c.rect(0, 0, 6 * mm, PAGE_H, fill=1, stroke=0)
    c.setFont("Helvetica", 8.5)
    c.setFillColor(RUST_DARK)
    c.drawCentredString(PAGE_W / 2, 11 * mm, str(c.getPageNumber()))
    c.restoreState()


def draw_normal(c, doc):
    c.saveState()
    c.setFillColor(colors.white)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    c.setFont("Helvetica", 8)
    c.setFillColor(GREY)
    c.drawString(MARGIN, PAGE_H - 14 * mm, TITLE)
    c.drawRightString(PAGE_W - MARGIN, PAGE_H - 14 * mm, _running["title"])
    c.setStrokeColor(SAND)
    c.setLineWidth(0.7)
    c.line(MARGIN, PAGE_H - 15.5 * mm, PAGE_W - MARGIN, PAGE_H - 15.5 * mm)
    c.line(MARGIN, 15.5 * mm, PAGE_W - MARGIN, 15.5 * mm)
    c.setFont("Helvetica", 8.5)
    c.setFillColor(RUST_DARK)
    c.drawCentredString(PAGE_W / 2, 11 * mm, str(c.getPageNumber()))
    c.restoreState()


# ============================================================
# MONTAGEM DO DOCUMENTO
# ============================================================
def action_box(action_text):
    p = Paragraph(
        '<font color="#9C3D22"><b>AÇÃO DE HOJE  →  </b></font>' + action_text,
        S_ACTION,
    )
    t = Table([[p]], colWidths=[PAGE_W - 2 * MARGIN - 16])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), SAND),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("TOPPADDING", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
        ("LINEBEFORE", (0, 0), (0, -1), 2.4, RUST),
    ]))
    return t


def build_tip_flowables(num, titulo, paragrafos, acao):
    flow = []
    flow.append(Paragraph(f"MANEIRA {num:02d}", S_TIP_NUM))
    flow.append(Paragraph(titulo, S_TIP_TITLE))
    for par in paragrafos:
        flow.append(Paragraph(par, S_BODY))
    flow.append(Spacer(1, 8))
    flow.append(action_box(acao))
    return [KeepTogether(flow), Spacer(1, 44), rust_rule(width=32 * mm, thickness=1, color=SAND, space_after=44)]


def build_story():
    story = []

    # ---------- CAPA ----------
    story.append(NextPageTemplate("ato"))
    story.append(PageBreak())

    # ---------- SUMARIO (usa template "ato" / fundo creme) ----------
    story.append(SetRunningTitle("Sumário"))
    story.append(Spacer(1, 4))
    story.append(Paragraph("SUMÁRIO", S_H1))
    story.append(rust_rule())
    story.append(Paragraph("Introdução  ·  Como usar este livro", S_SUMARIO_ITEM))
    story.append(Spacer(1, 10))
    for ato in ATOS:
        nums = [t[0] for t in ato["tips"]]
        story.append(Paragraph(f'{ato["numero"]} — {ato["titulo"]}', S_SUMARIO_H))
        story.append(Paragraph(
            f'Maneiras {nums[0]:02d} a {nums[-1]:02d}', S_SUMARIO_ITEM))
    story.append(Spacer(1, 10))
    story.append(Paragraph("Fechamento — Conclusão, plano de 30 dias e checklist final",
                            S_SUMARIO_H))

    # ---------- INTRODUCAO (template normal) ----------
    story.append(NextPageTemplate("normal"))
    story.append(PageBreak())
    story.append(SetRunningTitle("Introdução"))
    story.append(Paragraph("INTRODUÇÃO", S_H1))
    story.append(rust_rule())
    for par in INTRO_TEXT:
        story.append(Paragraph(par, S_BODY))

    story.append(Spacer(1, 16))
    story.append(Paragraph("COMO USAR ESTE LIVRO", S_H1))
    story.append(rust_rule())
    for par in COMO_USAR:
        story.append(Paragraph(par, S_BODY))

    # ---------- ATOS ----------
    for ato in ATOS:
        story.append(NextPageTemplate("ato"))
        story.append(PageBreak())
        story.append(SetRunningTitle(ato["titulo"]))
        story.append(Spacer(1, 60))
        story.append(Paragraph(ato["numero"].upper(), style(
            "ato_num", fontName="Helvetica-Bold", fontSize=13, textColor=RUST,
            alignment=TA_LEFT, spaceAfter=6)))
        story.append(Paragraph(ato["titulo"], style(
            "ato_title", fontName="Helvetica-Bold", fontSize=25, leading=30,
            textColor=INK, alignment=TA_LEFT, spaceAfter=16)))
        story.append(rust_rule(width=90 * mm, space_after=16))
        story.append(Paragraph("“" + ato["epigrafe"] + "”", style(
            "ato_epi", fontName="Helvetica-Oblique", fontSize=13, leading=18,
            textColor=RUST_DARK, alignment=TA_LEFT, spaceAfter=18)))
        story.append(Paragraph(ato["intro"], S_ATO_INTRO))

        story.append(NextPageTemplate("normal"))
        story.append(PageBreak())
        story.append(SetRunningTitle(ato["titulo"]))
        for (num, titulo, paragrafos, acao) in ato["tips"]:
            story.extend(build_tip_flowables(num, titulo, paragrafos, acao))

    # ---------- FECHAMENTO ----------
    story.append(NextPageTemplate("ato"))
    story.append(PageBreak())
    story.append(SetRunningTitle("Fechamento"))
    story.append(Spacer(1, 60))
    story.append(Paragraph("FECHAMENTO", style(
        "fim_num", fontName="Helvetica-Bold", fontSize=13, textColor=RUST,
        alignment=TA_LEFT, spaceAfter=6)))
    story.append(Paragraph("O Próximo Passo é Seu", style(
        "fim_title", fontName="Helvetica-Bold", fontSize=25, leading=30,
        textColor=INK, alignment=TA_LEFT, spaceAfter=16)))
    story.append(rust_rule(width=90 * mm, space_after=16))
    story.append(Paragraph("“Você não precisa terminar hoje. Precisa só começar.”", style(
        "fim_epi", fontName="Helvetica-Oblique", fontSize=13, leading=18,
        textColor=RUST_DARK, alignment=TA_LEFT)))

    story.append(NextPageTemplate("normal"))
    story.append(PageBreak())
    story.append(SetRunningTitle("Conclusão"))
    story.append(Paragraph("CONCLUSÃO", S_H1))
    story.append(rust_rule())
    for par in CONCLUSAO:
        story.append(Paragraph(par, S_BODY))

    story.append(NextPageTemplate("normal"))
    story.append(PageBreak())
    story.append(SetRunningTitle("Plano de 30 dias"))
    story.append(Paragraph("PLANO DE AÇÃO DE 30 DIAS", S_H1))
    story.append(rust_rule())
    for titulo, corpo in PLANO_30_DIAS:
        story.append(KeepTogether([
            Paragraph(titulo, S_WEEK_TITLE),
            Paragraph(corpo, S_WEEK_BODY),
        ]))

    story.append(NextPageTemplate("normal"))
    story.append(PageBreak())
    story.append(SetRunningTitle("Checklist das 45 maneiras"))
    story.append(Paragraph("CHECKLIST DAS 45 MANEIRAS", S_H1))
    story.append(rust_rule())
    idx = 0
    for ato in ATOS:
        story.append(Paragraph(f'{ato["numero"]} — {ato["titulo"]}', S_CHECK_ATO))
        for (num, titulo, _, _) in ato["tips"]:
            story.append(Paragraph(f"[  ]  {num:02d}. {titulo}", S_CHECK_ITEM))
        story.append(Spacer(1, 8))

    story.append(Spacer(1, 22))
    story.append(rust_rule(width=90 * mm))
    story.append(Paragraph(
        "“Não estamos buscando um dia perfeito. Estamos construindo uma identidade.”",
        style("final_quote", fontName="Helvetica-Oblique", fontSize=12.5, leading=17,
              textColor=RUST_DARK, alignment=TA_CENTER, spaceBefore=14)))

    return story


def main():
    doc = BaseDocTemplate(
        OUT_PATH,
        pagesize=A4,
        topMargin=MARGIN,
        bottomMargin=MARGIN,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        title=TITLE,
        author="Hoje, Não Amanhã",
    )

    frame_normal = Frame(MARGIN, MARGIN, PAGE_W - 2 * MARGIN, PAGE_H - 2 * MARGIN,
                          id="normal_frame", topPadding=6, bottomPadding=6)
    frame_ato = Frame(20 * mm, MARGIN, PAGE_W - 20 * mm - MARGIN, PAGE_H - 2 * MARGIN,
                       id="ato_frame", topPadding=6, bottomPadding=6)

    doc.addPageTemplates([
        PageTemplate(id="cover", frames=[frame_normal], onPage=draw_cover),
        PageTemplate(id="ato", frames=[frame_ato], onPage=draw_ato),
        PageTemplate(id="normal", frames=[frame_normal], onPage=draw_normal),
    ])

    story = build_story()
    doc.build(story)
    print(f"OK -> {OUT_PATH}")


if __name__ == "__main__":
    main()

