// =============================================
// ===== CURSO COMPLETO DE VBA - 25 LIÇÕES =====
// =============================================

const licoes = [
    // ========================================
    // MÓDULO 1: INTRODUÇÃO (Lições 1-5)
    // ========================================
    {
        id: 1,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "O que é VBA?",
        conceito: `
            <p><strong>VBA</strong> = Visual Basic for Applications</p>
            <br>
            <p>É a linguagem de programação que vem <strong>dentro do Excel</strong>.</p>
            <br>
            <p><strong>Para que serve?</strong></p>
            <ul>
                <li>✅ Automatizar tarefas repetitivas</li>
                <li>✅ Criar funções personalizadas</li>
                <li>✅ Manipular planilhas automaticamente</li>
            </ul>
            <br>
            <p>💡 Pense no VBA como um <strong>robô</strong> que faz tarefas para você!</p>
        `,
        codigo: `
' Meu primeiro código em VBA!
' Tudo que começa com ' é um comentário

Sub Saudacao()
    ' MsgBox mostra uma mensagem na tela
    MsgBox "Olá, eu sou o VBA!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Explicação:</strong></p>
            <p>• <span class="highlight">'</span> → Comentário (ignorado pelo computador)</p>
            <p>• <span class="highlight">Sub Saudacao()</span> → Início da macro</p>
            <p>• <span class="highlight">MsgBox</span> → Comando para mostrar mensagem</p>
            <p>• <span class="highlight">End Sub</span> → Fim da macro</p>
        `,
        passoAPasso: `
            <ol>
                <li>Abra o Excel</li>
                <li>Pressione <strong>ALT + F11</strong></li>
                <li>Clique em <strong>Inserir → Módulo</strong></li>
                <li>Cole o código</li>
                <li>Pressione <strong>F5</strong> para executar</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "O que significa a sigla VBA?",
        opcoes: [
            "Visual Basic for Applications",
            "Very Basic Automation",
            "Visual Basic Advanced",
            "Virtual Business Assistant"
        ],
        respostaCorreta: 0,
        explicacaoAcerto: "✅ Perfeito! VBA = Visual Basic for Applications.",
        explicacaoErro: "❌ VBA = Visual Basic for Applications."
    },
    {
        id: 2,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "Atalho ALT + F11",
        conceito: `
            <p><strong>ALT + F11</strong> é o atalho mais importante do VBA.</p>
            <br>
            <p>Ele <strong>abre e fecha</strong> o Editor VBA.</p>
            <br>
            <p><strong>Outros atalhos importantes:</strong></p>
            <ul>
                <li>📌 <strong>F5</strong> → Executa a macro</li>
                <li>📌 <strong>F8</strong> → Executa linha por linha (depuração)</li>
                <li>📌 <strong>Ctrl + G</strong> → Abre a Janela Imediata</li>
            </ul>
            <br>
            <p>💡 Decore o ALT + F11! Você vai usar ele o tempo todo!</p>
        `,
        codigo: `
' Atalhos essenciais do VBA:
' ALT + F11 = Abre/fecha o editor
' F5 = Executa a macro
' F8 = Executa passo a passo
' Ctrl + G = Janela Imediata
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Atalhos para decorar:</strong></p>
            <p>• <span class="highlight">ALT + F11</span> → Abre o Editor VBA</p>
            <p>• <span class="highlight">F5</span> → Executa o código</p>
            <p>• <span class="highlight">F8</span> → Depuração passo a passo</p>
            <p>• <span class="highlight">Ctrl + G</span> → Janela Imediata</p>
        `,
        passoAPasso: `
            <ol>
                <li>Abra o Excel</li>
                <li>Pressione <strong>ALT + F11</strong></li>
                <li>O Editor VBA vai abrir</li>
                <li>Pressione <strong>ALT + F11</strong> de novo</li>
                <li>O Editor VBA vai fechar</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual atalho abre e fecha o Editor VBA?",
        opcoes: [
            "Ctrl + C",
            "Alt + F11",
            "Ctrl + V",
            "Windows + D"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ ALT + F11 é o atalho para abrir/fechar o Editor VBA.",
        explicacaoErro: "❌ O atalho correto é ALT + F11."
    },
    {
        id: 3,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "O que é um Módulo?",
        conceito: `
            <p><strong>Módulo</strong> é o lugar onde você escreve o código VBA.</p>
            <br>
            <p>Pense no módulo como uma <strong>folha de papel em branco</strong> onde você vai escrever suas instruções.</p>
            <br>
            <p><strong>Como criar um módulo:</strong></p>
            <ul>
                <li>1. Abra o Editor VBA (ALT + F11)</li>
                <li>2. Clique em <strong>Inserir</strong> no menu</li>
                <li>3. Clique em <strong>Módulo</strong></li>
                <li>4. Pronto! Uma janela em branco vai aparecer</li>
            </ul>
            <br>
            <p>💡 Você pode ter vários módulos para organizar seu código!</p>
        `,
        codigo: `
' Como criar seu primeiro módulo:

' 1. ALT + F11 (abre o editor)
' 2. Inserir → Módulo
' 3. Digite seu código
' 4. F5 para executar

Sub PrimeiroModulo()
    MsgBox "Criei meu primeiro módulo!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Passos para criar um módulo:</strong></p>
            <p>• <span class="highlight">ALT + F11</span> → Abre o editor</p>
            <p>• <span class="highlight">Inserir → Módulo</span> → Cria o módulo</p>
            <p>• <span class="highlight">Digite o código</span> → Escreva no módulo</p>
            <p>• <span class="highlight">F5</span> → Executa</p>
        `,
        passoAPasso: `
            <ol>
                <li>Abra o Excel</li>
                <li>Pressione <strong>ALT + F11</strong></li>
                <li>No menu, clique em <strong>Inserir</strong></li>
                <li>Clique em <strong>Módulo</strong></li>
                <li>Uma janela em branco aparecerá</li>
                <li>Digite seu código e pressione <strong>F5</strong></li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "O que é um Módulo no VBA?",
        opcoes: [
            "Uma planilha do Excel",
            "O lugar onde você escreve o código",
            "Um tipo de gráfico",
            "Uma função matemática"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Módulo é o local onde você escreve o código VBA.",
        explicacaoErro: "❌ Módulo é o lugar onde você escreve o código."
    },
    {
        id: 4,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "Comentários no VBA",
        conceito: `
            <p><strong>Comentário</strong> é um texto que o computador ignora.</p>
            <br>
            <p>Servem para <strong>explicar o código</strong> para outras pessoas (ou para você mesmo no futuro).</p>
            <br>
            <p><strong>Como escrever um comentário:</strong></p>
            <ul>
                <li>📌 Use <strong> ' </strong> (apóstrofo) antes do texto</li>
                <li>📌 Ex: ' Isso é um comentário</li>
                <li>📌 Ex: Dim nome As String ' Guarda o nome</li>
            </ul>
            <br>
            <p>💡 <strong>Boa prática:</strong> Sempre comente seu código!</p>
        `,
        codigo: `
' ===== MEU PRIMEIRO PROGRAMA =====
' Este programa mostra uma saudação

Sub Saudacao()
    ' Declara uma variável para guardar o nome
    Dim nome As String
    
    ' Pergunta o nome do usuário
    nome = InputBox("Qual é o seu nome?")
    
    ' Mostra a saudação
    MsgBox "Olá, " & nome & "! Bem-vindo!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Uso de comentários:</strong></p>
            <p>• <span class="highlight">' ===== TÍTULO =====</span> → Separador visual</p>
            <p>• <span class="highlight">' Explica o que faz</span> → Comentário explicativo</p>
            <p>• <span class="highlight">' Declara uma variável</span> → Explica a linha seguinte</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo novo</li>
                <li>Digite o código com comentários</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Observe que os comentários são ignorados</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual símbolo é usado para criar um comentário em VBA?",
        opcoes: [
            "//",
            "' (apóstrofo)",
            "#",
            "/* */"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Em VBA, usamos o apóstrofo (') para comentários.",
        explicacaoErro: "❌ Em VBA, o símbolo de comentário é o apóstrofo (')."
    },
    {
        id: 5,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "MsgBox - Mostrando Mensagens",
        conceito: `
            <p><strong>MsgBox</strong> é o comando mais básico do VBA.</p>
            <br>
            <p>Ele serve para <strong>mostrar uma mensagem</strong> na tela.</p>
            <br>
            <p><strong>Sintaxe:</strong></p>
            <ul>
                <li>📌 <strong>MsgBox "Sua mensagem aqui"</strong></li>
                <li>📌 Ex: MsgBox "Olá, mundo!"</li>
                <li>📌 Ex: MsgBox "Resultado: " & 10 + 5</li>
            </ul>
            <br>
            <p>💡 O MsgBox é a primeira coisa que todo programador VBA aprende!</p>
        `,
        codigo: `
Sub PrimeiroMsgBox()
    ' MsgBox com texto simples
    MsgBox "Olá, mundo!"
    
    ' MsgBox com texto e variável
    Dim nome As String
    nome = "João"
    MsgBox "Bem-vindo, " & nome
    
    ' MsgBox com título
    MsgBox "Operação concluída!", vbInformation, "Sucesso"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Exemplos de MsgBox:</strong></p>
            <p>• <span class="highlight">MsgBox "Olá!"</span> → Mensagem simples</p>
            <p>• <span class="highlight">MsgBox "Olá, " & nome</span> → Texto + variável</p>
            <p>• <span class="highlight">vbInformation</span> → Ícone de informação</p>
            <p>• <span class="highlight">"Título"</span> → Título da janela</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja a mensagem aparecer!</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Para que serve o comando MsgBox?",
        opcoes: [
            "Para receber dados do usuário",
            "Para mostrar uma mensagem na tela",
            "Para salvar uma planilha",
            "Para criar uma variável"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ MsgBox mostra uma mensagem para o usuário.",
        explicacaoErro: "❌ MsgBox é usado para mostrar mensagens."
    },

    // ========================================
    // MÓDULO 2: VARIÁVEIS (Lições 6-10)
    // ========================================
    {
        id: 6,
        nivel: "📦 Módulo 2 - Variáveis",
        titulo: "O que é uma Variável?",
        conceito: `
            <p><strong>Variável</strong> é um espaço na memória para guardar informações.</p>
            <br>
            <p>Imagine uma <strong>caixa</strong> com uma <strong>etiqueta</strong>:</p>
            <br>
            <ul>
                <li>📦 <strong>nome</strong> = "João" → caixa que guarda texto</li>
                <li>📦 <strong>idade</strong> = 25 → caixa que guarda número</li>
                <li>📦 <strong>salario</strong> = 1500.50 → caixa que guarda decimal</li>
            </ul>
            <br>
            <p>💡 A variável é a coisa mais importante em programação!</p>
        `,
        codigo: `
Sub TrabalhandoComVariaveis()
    ' DECLARANDO VARIÁVEIS
    Dim nome As String      ' Texto
    Dim idade As Integer    ' Número inteiro
    Dim altura As Double    ' Número decimal
    Dim ativo As Boolean    ' Verdadeiro/Falso
    
    ' ATRIBUINDO VALORES
    nome = "Maria"
    idade = 30
    altura = 1.65
    ativo = True
    
    ' MOSTRANDO OS VALORES
    MsgBox "Nome: " & nome
    MsgBox "Idade: " & idade
    MsgBox "Altura: " & altura
    MsgBox "Ativo: " & ativo
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Explicação:</strong></p>
            <p>• <span class="highlight">Dim nome As String</span> → Declara variável para texto</p>
            <p>• <span class="highlight">Dim idade As Integer</span> → Declara variável para número inteiro</p>
            <p>• <span class="highlight">Dim altura As Double</span> → Declara variável para decimal</p>
            <p>• <span class="highlight">nome = "Maria"</span> → Atribui valor à variável</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja as mensagens com os valores das variáveis</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual comando é usado para criar uma variável em VBA?",
        opcoes: [
            "Create",
            "Dim",
            "New",
            "Variable"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Dim é a palavra-chave para declarar variáveis em VBA.",
        explicacaoErro: "❌ Em VBA, usamos 'Dim' para declarar variáveis."
    },
    {
        id: 7,
        nivel: "📦 Módulo 2 - Variáveis",
        titulo: "Tipos de Dados - String",
        conceito: `
            <p><strong>String</strong> é o tipo de dado para <strong>textos</strong>.</p>
            <br>
            <p><strong>Como usar:</strong></p>
            <ul>
                <li>📌 <strong>Dim nome As String</strong></li>
                <li>📌 <strong>nome = "João Silva"</strong></li>
            </ul>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>✅ "João"</li>
                <li>✅ "Rua das Flores, 123"</li>
                <li>✅ "vba@excel.com"</li>
            </ul>
            <br>
            <p>💡 String SEMPRE vai entre <strong>aspas duplas</strong> ("")</p>
        `,
        codigo: `
Sub TrabalhandoComString()
    Dim nome As String
    Dim endereco As String
    Dim email As String
    
    nome = "Carlos Silva"
    endereco = "Rua das Flores, 123"
    email = "carlos@email.com"
    
    MsgBox "Nome: " & nome
    MsgBox "Endereço: " & endereco
    MsgBox "E-mail: " & email
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Regras da String:</strong></p>
            <p>• <span class="highlight">As String</span> → Declara que é texto</p>
            <p>• <span class="highlight">"Entre aspas"</span> → O valor sempre entre aspas</p>
            <p>• <span class="highlight">"João" & " " & "Silva"</span> → Concatenação com &</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja como as strings são exibidas</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual tipo de dado é usado para texto em VBA?",
        opcoes: [
            "Text",
            "String",
            "Character",
            "Word"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ String é o tipo correto para texto em VBA.",
        explicacaoErro: "❌ Em VBA, String é o tipo para texto."
    },
    {
        id: 8,
        nivel: "📦 Módulo 2 - Variáveis",
        titulo: "Tipos de Dados - Integer e Double",
        conceito: `
            <p><strong>Integer</strong> = Números inteiros (sem vírgula)</p>
            <p><strong>Double</strong> = Números decimais (com vírgula)</p>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>📌 <strong>Integer:</strong> 10, 25, -3, 0, 1000</li>
                <li>📌 <strong>Double:</strong> 3.14, 2.5, -7.8, 0.99</li>
            </ul>
            <br>
            <p>💡 Use Integer para números sem vírgula e Double para números com vírgula!</p>
        `,
        codigo: `
Sub Numeros()
    ' INTEGER: números inteiros
    Dim idade As Integer
    Dim quantidade As Integer
    
    idade = 25
    quantidade = 10
    
    ' DOUBLE: números decimais
    Dim peso As Double
    Dim preco As Double
    
    peso = 70.5
    preco = 99.99
    
    MsgBox "Idade: " & idade
    MsgBox "Quantidade: " & quantidade
    MsgBox "Peso: " & peso
    MsgBox "Preço: R$ " & preco
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Diferenças:</strong></p>
            <p>• <span class="highlight">Integer</span> → Números sem vírgula (ex: 10, 25)</p>
            <p>• <span class="highlight">Double</span> → Números com vírgula (ex: 3.14, 2.5)</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja a diferença entre números inteiros e decimais</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual tipo de dado deve ser usado para guardar o valor 3.14?",
        opcoes: [
            "Integer",
            "Double",
            "String",
            "Boolean"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Double é usado para números decimais como 3.14.",
        explicacaoErro: "❌ Para números com vírgula, usamos Double."
    },
    {
        id: 9,
        nivel: "📦 Módulo 2 - Variáveis",
        titulo: "Tipos de Dados - Boolean e Date",
        conceito: `
            <p><strong>Boolean</strong> = Verdadeiro ou Falso</p>
            <p><strong>Date</strong> = Datas</p>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>📌 <strong>Boolean:</strong> True, False</li>
                <li>📌 <strong>Date:</strong> #25/12/2024#</li>
            </ul>
            <br>
            <p>💡 Boolean só tem dois valores: True (verdadeiro) ou False (falso)!</p>
        `,
        codigo: `
Sub BooleanDate()
    Dim ativo As Boolean
    Dim dataNasc As Date
    
    ativo = True
    dataNasc = #15/05/1990#
    
    MsgBox "Ativo: " & ativo
    MsgBox "Data de nascimento: " & dataNasc
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Tipos:</strong></p>
            <p>• <span class="highlight">Boolean</span> → True ou False</p>
            <p>• <span class="highlight">Date</span> → Datas (#dd/mm/aaaa#)</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Quais são os valores possíveis para uma variável Boolean?",
        opcoes: [
            "Sim e Não",
            "True e False",
            "1 e 0",
            "Yes e No"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Boolean só pode ser True (verdadeiro) ou False (falso).",
        explicacaoErro: "❌ Boolean em VBA usa True ou False."
    },
    {
        id: 10,
        nivel: "📦 Módulo 2 - Variáveis",
        titulo: "Concatenação com &",
        conceito: `
            <p><strong>Concatenação</strong> é o ato de <strong>juntar textos</strong>.</p>
            <br>
            <p>Em VBA, usamos o símbolo <strong>&</strong> para concatenar.</p>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>📌 <strong>"Olá, " & "mundo!"</strong> → "Olá, mundo!"</li>
                <li>📌 <strong>nome & " tem " & idade & " anos"</strong></li>
                <li>📌 <strong>"Total: R$ " & total</strong></li>
            </ul>
            <br>
            <p>💡 Use <strong>&</strong> para juntar textos e variáveis!</p>
        `,
        codigo: `
Sub Concatenacao()
    Dim nome As String
    Dim idade As Integer
    Dim cidade As String
    
    nome = "Ana"
    idade = 25
    cidade = "São Paulo"
    
    MsgBox "Meu nome é " & nome
    MsgBox nome & " tem " & idade & " anos e mora em " & cidade
    
    Dim valor As Double
    valor = 150.75
    MsgBox "O valor é R$ " & valor
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Sobre a concatenação:</strong></p>
            <p>• <span class="highlight">&</span> → Símbolo para juntar textos</p>
            <p>• <span class="highlight">"Texto " & variavel</span> → Junta texto com variável</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual símbolo é usado para concatenar (juntar) textos em VBA?",
        opcoes: [
            "+",
            "&",
            "|",
            "."
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ O símbolo & é usado para concatenar textos em VBA.",
        explicacaoErro: "❌ Em VBA, usamos o & para concatenar."
    },

    // ========================================
    // MÓDULO 3: COMANDOS BÁSICOS (Lições 11-15)
    // ========================================
    {
        id: 11,
        nivel: "💬 Módulo 3 - Comandos Básicos",
        titulo: "MsgBox com Botões",
        conceito: `
            <p>O MsgBox pode mostrar <strong>botões diferentes</strong> para o usuário.</p>
            <br>
            <p><strong>Tipos de botões:</strong></p>
            <ul>
                <li>📌 <strong>vbOKOnly</strong> → Só o botão OK (padrão)</li>
                <li>📌 <strong>vbYesNo</strong> → Botões Sim e Não</li>
                <li>📌 <strong>vbYesNoCancel</strong> → Sim, Não e Cancelar</li>
            </ul>
            <br>
            <p>💡 Use MsgBox com botões para <strong>interagir com o usuário</strong>!</p>
        `,
        codigo: `
Sub MsgBoxComBotoes()
    Dim resposta As Integer
    
    ' MsgBox com botões Sim e Não
    resposta = MsgBox("Deseja continuar?", vbYesNo, "Confirmação")
    
    If resposta = vbYes Then
        MsgBox "Você escolheu SIM!"
    Else
        MsgBox "Você escolheu NÃO!"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Constantes de botões:</strong></p>
            <p>• <span class="highlight">vbYesNo</span> → Mostra Sim e Não</p>
            <p>• <span class="highlight">vbYes</span> → Valor quando clica em Sim</p>
            <p>• <span class="highlight">vbNo</span> → Valor quando clica em Não</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Teste clicando nos diferentes botões</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual constante mostra botões Sim e Não no MsgBox?",
        opcoes: [
            "vbOKOnly",
            "vbYesNo",
            "vbYesNoCancel",
            "vbRetryCancel"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ vbYesNo mostra os botões Sim e Não.",
        explicacaoErro: "❌ vbYesNo é a constante para mostrar Sim e Não."
    },
    {
        id: 12,
        nivel: "💬 Módulo 3 - Comandos Básicos",
        titulo: "InputBox - Capturando Dados",
        conceito: `
            <p><strong>InputBox</strong> permite que o usuário <strong>digite informações</strong>.</p>
            <br>
            <p><strong>Sintaxe:</strong></p>
            <ul>
                <li>📌 <strong>variavel = InputBox("Pergunta")</strong></li>
                <li>📌 <strong>variavel = InputBox("Pergunta", "Título")</strong></li>
            </ul>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>✅ <strong>nome = InputBox("Qual seu nome?")</strong></li>
                <li>✅ <strong>idade = InputBox("Digite sua idade:")</strong></li>
            </ul>
            <br>
            <p>💡 O InputBox sempre retorna texto (String)!</p>
        `,
        codigo: `
Sub InputBoxExemplo()
    Dim nome As String
    Dim idade As Integer
    
    nome = InputBox("Qual é o seu nome?")
    
    If nome <> "" Then
        idade = InputBox("Qual é a sua idade?")
        idade = CInt(idade)
        MsgBox "Olá, " & nome & "! Você tem " & idade & " anos."
    Else
        MsgBox "Você não digitou seu nome!"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Sobre InputBox:</strong></p>
            <p>• <span class="highlight">InputBox("Pergunta")</span> → Pergunta simples</p>
            <p>• <span class="highlight">CInt(variavel)</span> → Converte texto para número</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Responda as perguntas do InputBox</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "O que o InputBox retorna (sempre)?",
        opcoes: [
            "Número (Integer)",
            "Texto (String)",
            "Verdadeiro/Falso (Boolean)",
            "Data (Date)"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ InputBox sempre retorna uma String (texto).",
        explicacaoErro: "❌ InputBox retorna texto."
    },
    {
        id: 13,
        nivel: "💬 Módulo 3 - Comandos Básicos",
        titulo: "Constantes (Const)",
        conceito: `
            <p><strong>Constantes</strong> são variáveis que <strong>não podem mudar de valor</strong>.</p>
            <br>
            <p><strong>Diferença:</strong></p>
            <ul>
                <li>📌 <strong>Dim</strong> → Variável (pode mudar)</li>
                <li>📌 <strong>Const</strong> → Constante (NUNCA muda)</li>
            </ul>
            <br>
            <p>💡 Use <strong>CONST</strong> para valores que <strong>nunca vão mudar</strong>!</p>
        `,
        codigo: `
Sub Constantes()
    Const PI As Double = 3.14159
    Const TAXA As Double = 0.12
    Const NOME_EMPRESA As String = "Minha Empresa"
    
    MsgBox "Empresa: " & NOME_EMPRESA
    MsgBox "Valor de PI: " & PI
    
    Dim raio As Double
    raio = InputBox("Digite o raio do círculo:")
    raio = CDbl(raio)
    
    Dim area As Double
    area = PI * raio * raio
    MsgBox "A área do círculo é: " & area
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Regras das Constantes:</strong></p>
            <p>• <span class="highlight">Const PI As Double = 3.14159</span> → Declara constante</p>
            <p>• <span class="highlight">NÃO pode ser alterada</span> → Dá erro se tentar</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja como as constantes são usadas</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual a diferença entre Dim e Const?",
        opcoes: [
            "Não tem diferença",
            "Dim pode mudar, Const não pode",
            "Const pode mudar, Dim não pode",
            "Dim é para números, Const é para texto"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Dim pode mudar de valor, Const não pode (é constante).",
        explicacaoErro: "❌ Const é constante (não muda) e Dim é variável (pode mudar)."
    },
    {
        id: 14,
        nivel: "💬 Módulo 3 - Comandos Básicos",
        titulo: "Operadores Aritméticos",
        conceito: `
            <p><strong>Operadores aritméticos</strong> são usados para fazer cálculos.</p>
            <br>
            <p><strong>Operadores disponíveis:</strong></p>
            <ul>
                <li>📌 <strong>+</strong> → Soma (ex: 10 + 5 = 15)</li>
                <li>📌 <strong>-</strong> → Subtração (ex: 10 - 5 = 5)</li>
                <li>📌 <strong>*</strong> → Multiplicação (ex: 10 * 5 = 50)</li>
                <li>📌 <strong>/</strong> → Divisão (ex: 10 / 5 = 2)</li>
                <li>📌 <strong>^</strong> → Potência (ex: 2 ^ 3 = 8)</li>
                <li>📌 <strong>Mod</strong> → Resto da divisão (ex: 10 Mod 3 = 1)</li>
            </ul>
            <br>
            <p>💡 <strong>Mod</strong> é muito útil para saber se um número é par ou ímpar!</p>
        `,
        codigo: `
Sub Operadores()
    Dim a As Double
    Dim b As Double
    
    a = InputBox("Digite o primeiro número:")
    b = InputBox("Digite o segundo número:")
    a = CDbl(a)
    b = CDbl(b)
    
    MsgBox a & " + " & b & " = " & (a + b)
    MsgBox a & " - " & b & " = " & (a - b)
    MsgBox a & " x " & b & " = " & (a * b)
    
    If b <> 0 Then
        MsgBox a & " / " & b & " = " & (a / b)
    Else
        MsgBox "Não é possível dividir por zero!"
    End If
    
    MsgBox a & " ao quadrado = " & (a ^ 2)
    MsgBox "Resto de " & a & " / " & b & " = " & (a Mod b)
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Operadores:</strong></p>
            <p>• <span class="highlight">+</span> → Soma</p>
            <p>• <span class="highlight">-</span> → Subtração</p>
            <p>• <span class="highlight">*</span> → Multiplicação</p>
            <p>• <span class="highlight">/</span> → Divisão</p>
            <p>• <span class="highlight">^</span> → Potência</p>
            <p>• <span class="highlight">Mod</span> → Resto da divisão</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Digite dois números e veja os cálculos</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual operador calcula o resto da divisão em VBA?",
        opcoes: [
            "/",
            "\\",
            "Mod",
            "Rest"
        ],
        respostaCorreta: 2,
        explicacaoAcerto: "✅ Mod calcula o resto da divisão em VBA.",
        explicacaoErro: "❌ O operador para resto da divisão é Mod."
    },
    {
        id: 15,
        nivel: "💬 Módulo 3 - Comandos Básicos",
        titulo: "Operadores de Comparação",
        conceito: `
            <p><strong>Operadores de comparação</strong> comparam valores e retornam <strong>True ou False</strong>.</p>
            <br>
            <p><strong>Operadores:</strong></p>
            <ul>
                <li>📌 <strong>=</strong> → Igual a</li>
                <li>📌 <strong><></strong> → Diferente de</li>
                <li>📌 <strong>></strong> → Maior que</li>
                <li>📌 <strong><</strong> → Menor que</li>
                <li>📌 <strong>>=</strong> → Maior ou igual</li>
                <li>📌 <strong><=</strong> → Menor ou igual</li>
            </ul>
            <br>
            <p>💡 Operadores de comparação são usados principalmente com <strong>IF</strong>!</p>
        `,
        codigo: `
Sub Comparacao()
    Dim idade As Integer
    idade = InputBox("Digite sua idade:")
    idade = CInt(idade)
    
    If idade = 18 Then
        MsgBox "Você tem exatamente 18 anos!"
    End If
    
    If idade > 18 Then
        MsgBox "Você é maior de 18 anos!"
    End If
    
    If idade >= 18 Then
        MsgBox "Você é maior de idade!"
    End If
    
    If idade < 18 Then
        MsgBox "Você é menor de idade!"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Operadores:</strong></p>
            <p>• <span class="highlight">=</span> → Igual a</p>
            <p>• <span class="highlight"><></span> → Diferente de</p>
            <p>• <span class="highlight">></span> → Maior que</p>
            <p>• <span class="highlight"><</span> → Menor que</p>
            <p>• <span class="highlight">>=</span> → Maior ou igual</p>
            <p>• <span class="highlight"><=</span> → Menor ou igual</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Digite valores e veja as comparações</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual operador testa se dois valores são diferentes em VBA?",
        opcoes: [
            "!=",
            "<>",
            "><",
            "≠"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ <> é o operador para 'diferente de' em VBA.",
        explicacaoErro: "❌ Em VBA, <> significa 'diferente de'."
    },

    // ========================================
    // MÓDULO 4: LÓGICA (Lições 16-20)
    // ========================================
    {
        id: 16,
        nivel: "🔀 Módulo 4 - Lógica",
        titulo: "IF - Tomando Decisões",
        conceito: `
            <p><strong>IF</strong> é o comando para <strong>tomar decisões</strong> no código.</p>
            <br>
            <p><strong>Estrutura básica:</strong></p>
            <ul>
                <li>📌 <strong>If</strong> condição <strong>Then</strong></li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código se for verdadeiro</li>
                <li>📌 <strong>End If</strong></li>
            </ul>
            <br>
            <p><strong>Com ELSE (senão):</strong></p>
            <ul>
                <li>📌 <strong>If</strong> condição <strong>Then</strong></li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' se for verdadeiro</li>
                <li>📌 <strong>Else</strong></li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' se for falso</li>
                <li>📌 <strong>End If</strong></li>
            </ul>
            <br>
            <p>💡 IF é como uma <strong>pergunta de sim/não</strong> para o computador!</p>
        `,
        codigo: `
Sub EstruturaIF()
    Dim idade As Integer
    idade = InputBox("Qual é a sua idade?")
    idade = CInt(idade)
    
    ' IF com ELSE
    If idade >= 18 Then
        MsgBox "Você é maior de idade!"
    Else
        MsgBox "Você é menor de idade!"
    End If
    
    ' IF com ELSEIF
    If idade < 12 Then
        MsgBox "Você é uma criança!"
    ElseIf idade < 18 Then
        MsgBox "Você é um adolescente!"
    ElseIf idade < 60 Then
        MsgBox "Você é um adulto!"
    Else
        MsgBox "Você é um idoso!"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Estrutura do IF:</strong></p>
            <p>• <span class="highlight">If condição Then</span> → Início da decisão</p>
            <p>• <span class="highlight">Else</span> → Senão (opcional)</p>
            <p>• <span class="highlight">End If</span> → Fim do IF</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Digite sua idade</li>
                <li>Veja as mensagens baseadas nas condições</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual palavra finaliza um bloco IF em VBA?",
        opcoes: [
            "End IF",
            "End If",
            "Endif",
            "Close If"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ End If é a palavra que finaliza o bloco IF.",
        explicacaoErro: "❌ A estrutura correta é End If."
    },
    {
        id: 17,
        nivel: "🔀 Módulo 4 - Lógica",
        titulo: "ELSEIF - Múltiplas Condições",
        conceito: `
            <p><strong>ELSEIF</strong> permite testar <strong>várias condições</strong> em sequência.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>If</strong> condição1 <strong>Then</strong></li>
                <li>📌 <strong>ElseIf</strong> condição2 <strong>Then</strong></li>
                <li>📌 <strong>ElseIf</strong> condição3 <strong>Then</strong></li>
                <li>📌 <strong>Else</strong></li>
                <li>📌 <strong>End If</strong></li>
            </ul>
            <br>
            <p>💡 ELSEIF é como uma <strong>escada de decisões</strong>!</p>
        `,
        codigo: `
Sub ElseIfExemplo()
    Dim nota As Double
    nota = InputBox("Digite a nota do aluno (0-10):")
    nota = CDbl(nota)
    
    If nota >= 9 Then
        MsgBox "Conceito: A - Excelente!"
    ElseIf nota >= 7 Then
        MsgBox "Conceito: B - Bom!"
    ElseIf nota >= 5 Then
        MsgBox "Conceito: C - Regular"
    Else
        MsgBox "Conceito: D - Insuficiente"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Estrutura ELSEIF:</strong></p>
            <p>• <span class="highlight">If nota >= 9 Then</span> → Primeira condição</p>
            <p>• <span class="highlight">ElseIf nota >= 7 Then</span> → Segunda condição</p>
            <p>• <span class="highlight">Else</span> → Se nenhuma for verdadeira</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Teste diferentes notas</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual comando permite testar várias condições em sequência?",
        opcoes: [
            "ELSE",
            "ELSEIF",
            "ENDIF",
            "IFELSE"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ ELSEIF permite testar múltiplas condições em sequência.",
        explicacaoErro: "❌ ELSEIF é usado para múltiplas condições."
    },
    {
        id: 18,
        nivel: "🔀 Módulo 4 - Lógica",
        titulo: "SELECT CASE",
        conceito: `
            <p><strong>SELECT CASE</strong> é uma alternativa ao ELSEIF para <strong>múltiplas decisões</strong>.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>Select Case</strong> variável</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;<strong>Case</strong> valor1</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;<strong>Case</strong> valor2</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;<strong>Case Else</strong></li>
                <li>📌 <strong>End Select</strong></li>
            </ul>
            <br>
            <p>💡 SELECT CASE é mais <strong>organizado</strong> que vários ELSEIF!</p>
        `,
        codigo: `
Sub SelectCaseExemplo()
    Dim dia As Integer
    dia = InputBox("Digite um número (1-7):")
    dia = CInt(dia)
    
    Select Case dia
        Case 1
            MsgBox "Domingo"
        Case 2
            MsgBox "Segunda-feira"
        Case 3
            MsgBox "Terça-feira"
        Case 4
            MsgBox "Quarta-feira"
        Case 5
            MsgBox "Quinta-feira"
        Case 6
            MsgBox "Sexta-feira"
        Case 7
            MsgBox "Sábado"
        Case Else
            MsgBox "Número inválido! Digite 1 a 7."
    End Select
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Tipos de CASE:</strong></p>
            <p>• <span class="highlight">Case 1</span> → Valor exato</p>
            <p>• <span class="highlight">Case 1, 3, 5</span> → Múltiplos valores</p>
            <p>• <span class="highlight">Case 1 To 5</span> → Intervalo</p>
            <p>• <span class="highlight">Case Else</span> → Valor padrão</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Teste diferentes números</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual é a vantagem do SELECT CASE sobre vários ELSEIF?",
        opcoes: [
            "É mais rápido",
            "É mais organizado e legível",
            "Faz mais cálculos",
            "Não tem vantagem"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ SELECT CASE é mais organizado e fácil de ler.",
        explicacaoErro: "❌ SELECT CASE torna o código mais organizado."
    },
    {
        id: 19,
        nivel: "🔀 Módulo 4 - Lógica",
        titulo: "Operadores Lógicos - AND e OR",
        conceito: `
            <p><strong>Operadores lógicos</strong> combinam duas ou mais condições.</p>
            <br>
            <p><strong>Principais operadores:</strong></p>
            <ul>
                <li>📌 <strong>AND</strong> → Verdadeiro se TODAS forem verdadeiras</li>
                <li>📌 <strong>OR</strong> → Verdadeiro se PELO MENOS UMA for verdadeira</li>
                <li>📌 <strong>NOT</strong> → Inverte o valor</li>
            </ul>
            <br>
            <p>💡 Use AND para condições que precisam ser <strong>todas verdadeiras</strong>!</p>
            <p>Use OR para condições que precisam de <strong>pelo menos uma verdadeira</strong>!</p>
        `,
        codigo: `
Sub OperadoresLogicos()
    Dim idade As Integer
    Dim temCarteira As Boolean
    Dim resposta As Integer
    
    idade = InputBox("Digite sua idade:")
    idade = CInt(idade)
    
    resposta = MsgBox("Possui carteira de motorista?", vbYesNo)
    temCarteira = (resposta = vbYes)
    
    ' AND: ambas devem ser verdadeiras
    If idade >= 18 And temCarteira Then
        MsgBox "Você pode dirigir!"
    Else
        MsgBox "Você NÃO pode dirigir!"
    End If
    
    ' OR: pelo menos uma verdadeira
    Dim estaChovendo As Boolean
    Dim estaFrio As Boolean
    
    resposta = MsgBox("Está chovendo?", vbYesNo)
    estaChovendo = (resposta = vbYes)
    
    resposta = MsgBox("Está frio?", vbYesNo)
    estaFrio = (resposta = vbYes)
    
    If estaChovendo Or estaFrio Then
        MsgBox "Leve um guarda-chuva e agasalho!"
    Else
        MsgBox "Dia agradável!"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Operadores Lógicos:</strong></p>
            <p>• <span class="highlight">And</span> → Verdadeiro se TODAS forem verdadeiras</p>
            <p>• <span class="highlight">Or</span> → Verdadeiro se PELO MENOS UMA for verdadeira</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja a diferença entre AND e OR</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual operador lógico exige que TODAS as condições sejam verdadeiras?",
        opcoes: [
            "OR",
            "AND",
            "NOT",
            "XOR"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ AND exige que TODAS as condições sejam verdadeiras.",
        explicacaoErro: "❌ AND só é verdadeiro se TODAS as condições forem verdadeiras."
    },
    {
        id: 20,
        nivel: "🔀 Módulo 4 - Lógica",
        titulo: "Verificando Par ou Ímpar",
        conceito: `
            <p>Para verificar se um número é par ou ímpar, usamos o operador <strong>Mod</strong>.</p>
            <br>
            <p><strong>Regra:</strong></p>
            <ul>
                <li>📌 <strong>numero Mod 2 = 0</strong> → Número é PAR</li>
                <li>📌 <strong>numero Mod 2 = 1</strong> → Número é ÍMPAR</li>
            </ul>
            <br>
            <p><strong>Outras verificações:</strong></p>
            <ul>
                <li>✅ <strong>numero Mod 3 = 0</strong> → Múltiplo de 3</li>
                <li>✅ <strong>numero Mod 5 = 0</strong> → Múltiplo de 5</li>
            </ul>
            <br>
            <p>💡 Mod retorna o <strong>resto da divisão</strong>!</p>
        `,
        codigo: `
Sub ParOuImpar()
    Dim numero As Integer
    numero = InputBox("Digite um número:")
    numero = CInt(numero)
    
    If numero Mod 2 = 0 Then
        MsgBox numero & " é PAR!"
    Else
        MsgBox numero & " é ÍMPAR!"
    End If
    
    If numero Mod 3 = 0 Then
        MsgBox numero & " é múltiplo de 3!"
    End If
    
    If numero Mod 5 = 0 Then
        MsgBox numero & " é múltiplo de 5!"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Verificações com Mod:</strong></p>
            <p>• <span class="highlight">numero Mod 2 = 0</span> → Par</p>
            <p>• <span class="highlight">numero Mod 2 = 1</span> → Ímpar</p>
            <p>• <span class="highlight">numero Mod 3 = 0</span> → Múltiplo de 3</p>
            <p>• <span class="highlight">numero Mod 5 = 0</span> → Múltiplo de 5</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Digite diferentes números</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Como verificar se um número é par em VBA?",
        opcoes: [
            "numero / 2 = 0",
            "numero Mod 2 = 0",
            "numero \ 2 = 0",
            "numero % 2 = 0"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ numero Mod 2 = 0 significa que é par.",
        explicacaoErro: "❌ Use numero Mod 2 = 0 para verificar se é par."
    },

    // ========================================
    // MÓDULO 5: LOOPS (Lições 21-25)
    // ========================================
    {
        id: 21,
        nivel: "🔄 Módulo 5 - Loops",
        titulo: "FOR - Repetindo Código",
        conceito: `
            <p><strong>FOR</strong> é um loop que repete um código <strong>um número específico</strong> de vezes.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>For</strong> i = 1 <strong>To</strong> 10</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código para repetir</li>
                <li>📌 <strong>Next</strong> i</li>
            </ul>
            <br>
            <p>💡 FOR é usado quando você <strong>sabe exatamente</strong> quantas vezes quer repetir!</p>
        `,
        codigo: `
Sub LoopFOR()
    Dim i As Integer
    
    ' Contagem simples de 1 a 5
    For i = 1 To 5
        MsgBox "Número: " & i
    Next i
    
    ' Somando números de 1 a 100
    Dim soma As Integer
    soma = 0
    
    For i = 1 To 100
        soma = soma + i
    Next i
    
    MsgBox "Soma de 1 a 100 = " & soma
    
    ' Tabuada
    Dim numero As Integer
    numero = InputBox("Digite um número para ver a tabuada:")
    numero = CInt(numero)
    
    Dim mensagem As String
    mensagem = "Tabuada do " & numero & ":" & vbCrLf
    
    For i = 1 To 10
        mensagem = mensagem & numero & " x " & i & " = " & (numero * i) & vbCrLf
    Next i
    
    MsgBox mensagem
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Estrutura do FOR:</strong></p>
            <p>• <span class="highlight">For i = 1 To 10</span> → i vai de 1 até 10</p>
            <p>• <span class="highlight">Step 2</span> → Pula de 2 em 2</p>
            <p>• <span class="highlight">Next i</span> → Vai para a próxima iteração</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja os loops executando</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Quantas vezes o código executa em: For i = 1 To 5?",
        opcoes: [
            "4 vezes",
            "5 vezes",
            "6 vezes",
            "10 vezes"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ De 1 a 5 são 5 execuções (1,2,3,4,5).",
        explicacaoErro: "❌ O loop executa 5 vezes: i = 1, 2, 3, 4, 5."
    },
    {
        id: 22,
        nivel: "🔄 Módulo 5 - Loops",
        titulo: "FOR - Preenchendo Planilhas",
        conceito: `
            <p>O loop FOR é muito útil para <strong>preencher planilhas</strong> automaticamente.</p>
            <br>
            <p><strong>Comando para células:</strong></p>
            <ul>
                <li>📌 <strong>Cells(linha, coluna)</strong> → Acessa uma célula</li>
                <li>📌 <strong>Cells(i, 1).Value = i</strong> → Coluna A recebe números</li>
            </ul>
            <br>
            <p>💡 FOR + Cells = <strong>automatizar planilhas</strong>!</p>
        `,
        codigo: `
Sub PreencherPlanilha()
    Dim i As Integer
    
    ' Preencher 10 linhas com números
    For i = 1 To 10
        Cells(i, 1).Value = i          ' Coluna A
        Cells(i, 2).Value = i * 2      ' Coluna B (dobro)
        Cells(i, 3).Value = i ^ 2      ' Coluna C (quadrado)
    Next i
    
    ' Destacar números maiores que 50
    For i = 1 To 10
        If Cells(i, 2).Value > 10 Then
            Cells(i, 2).Interior.Color = RGB(255, 255, 0) ' Amarelo
        End If
    Next i
    
    MsgBox "Planilha preenchida com sucesso!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Comandos de células:</strong></p>
            <p>• <span class="highlight">Cells(i, 1).Value = i</span> → Coluna A, linha i</p>
            <p>• <span class="highlight">Cells(i, 2).Value = i * 2</span> → Coluna B</p>
            <p>• <span class="highlight">Interior.Color = RGB(...)</span> → Cor de fundo</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Volte para a planilha (ALT + F11)</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja as células sendo preenchidas!</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "O que significa Cells(3, 2) no VBA?",
        opcoes: [
            "Linha 2, Coluna 3",
            "Linha 3, Coluna 2",
            "Célula B3",
            "Célula C2"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Cells(linha, coluna) → linha 3, coluna 2 = célula B3.",
        explicacaoErro: "❌ Cells(3, 2) = linha 3, coluna 2 = célula B3."
    },
    {
        id: 23,
        nivel: "🔄 Módulo 5 - Loops",
        titulo: "WHILE - Loop com Condição",
        conceito: `
            <p><strong>WHILE</strong> repete o código <strong>enquanto</strong> uma condição for <strong>verdadeira</strong>.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>While</strong> condição</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código para repetir</li>
                <li>📌 <strong>Wend</strong></li>
            </ul>
            <br>
            <p><strong>Diferença para FOR:</strong></p>
            <ul>
                <li>📌 <strong>FOR</strong> → Repete X vezes (número fixo)</li>
                <li>📌 <strong>WHILE</strong> → Repete enquanto for verdade</li>
            </ul>
            <br>
            <p>💡 WHILE é usado quando você <strong>não sabe</strong> quantas vezes vai repetir!</p>
        `,
        codigo: `
Sub WhileExemplo()
    Dim i As Integer
    Dim soma As Integer
    
    ' Contagem regressiva
    i = 5
    While i > 0
        MsgBox "Faltam " & i & " segundos..."
        i = i - 1
    Wend
    MsgBox "FIM!"
    
    ' Somar até passar de 100
    soma = 0
    i = 1
    While soma < 100
        soma = soma + i
        i = i + 1
    Wend
    MsgBox "Soma: " & soma
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>WHILE:</strong></p>
            <p>• <span class="highlight">While i > 0</span> → Enquanto i for maior que 0</p>
            <p>• <span class="highlight">Wend</span> → Fim do WHILE</p>
            <p>• <span class="highlight">i = i - 1</span> → Muda a condição (senão loop infinito!)</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja como o WHILE funciona</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "O que acontece se a condição do WHILE nunca ficar falsa?",
        opcoes: [
            "O programa para",
            "Loop infinito (travamento)",
            "O código não executa",
            "Mostra um erro"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Se a condição nunca ficar falsa, o loop é infinito.",
        explicacaoErro: "❌ WHILE sem condição de parada causa loop infinito."
    },
    {
        id: 24,
        nivel: "🔄 Módulo 5 - Loops",
        titulo: "DO WHILE - Executa pelo menos uma vez",
        conceito: `
            <p><strong>DO WHILE</strong> é similar ao WHILE, mas <strong>executa pelo menos uma vez</strong>.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>Do While</strong> condição</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código para repetir</li>
                <li>📌 <strong>Loop</strong></li>
            </ul>
            <br>
            <p><strong>Diferença:</strong></p>
            <ul>
                <li>📌 <strong>WHILE</strong> → Testa a condição ANTES de executar</li>
                <li>📌 <strong>DO WHILE</strong> → Testa a condição DEPOIS de executar</li>
            </ul>
            <br>
            <p>💡 DO WHILE garante que o código rode <strong>pelo menos uma vez</strong>!</p>
        `,
        codigo: `
Sub DoWhileExemplo()
    Dim numero As Integer
    
    ' Validação (executa pelo menos uma vez)
    Do
        numero = InputBox("Digite um número entre 1 e 10:")
        numero = CInt(numero)
    Loop While numero < 1 Or numero > 10
    
    MsgBox "Número válido: " & numero
    
    ' Menu com opções
    Dim resposta As Integer
    Do
        resposta = MsgBox("Escolha:" & vbCrLf & _
                          "Sim - Continuar" & vbCrLf & _
                          "Não - Sair", vbYesNo, "Menu")
        
        If resposta = vbYes Then
            MsgBox "Continuando..."
        Else
            MsgBox "Saindo..."
        End If
    Loop While resposta = vbYes
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>DO WHILE:</strong></p>
            <p>• <span class="highlight">Do ... Loop While condição</span> → Testa DEPOIS</p>
            <p>• <span class="highlight">Sempre executa pelo menos uma vez</span></p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja como o DO WHILE funciona</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual a diferença entre WHILE e DO WHILE?",
        opcoes: [
            "Não tem diferença",
            "DO WHILE executa pelo menos uma vez",
            "WHILE é mais rápido",
            "DO WHILE é para números"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ DO WHILE executa pelo menos uma vez antes de testar a condição.",
        explicacaoErro: "❌ DO WHILE garante pelo menos uma execução."
    },
    {
        id: 25,
        nivel: "🔄 Módulo 5 - Loops",
        titulo: "Saindo de um Loop - Exit For / Exit Do",
        conceito: `
            <p><strong>Exit For</strong> e <strong>Exit Do</strong> são usados para <strong>sair do loop</strong> antes do tempo.</p>
            <br>
            <p><strong>Quando usar?</strong></p>
            <ul>
                <li>✅ Quando encontrar o valor que está procurando</li>
                <li>✅ Quando ocorrer um erro</li>
                <li>✅ Quando o usuário cancelar uma operação</li>
            </ul>
            <br>
            <p>💡 Exit é útil para <strong>otimizar</strong> e <strong>controlar</strong> seus loops!</p>
        `,
        codigo: `
Sub ExitLoops()
    Dim i As Integer
    Dim procurado As Integer
    Dim encontrado As Boolean
    
    ' Exit For
    procurado = InputBox("Digite um número para procurar (1-50):")
    procurado = CInt(procurado)
    
    encontrado = False
    For i = 1 To 50
        If i = procurado Then
            encontrado = True
            MsgBox "Número " & procurado & " encontrado na posição " & i
            Exit For ' Sai do loop quando encontrar
        End If
    Next i
    
    If Not encontrado Then
        MsgBox "Número não encontrado!"
    End If
    
    ' Exit Do
    Dim senha As String
    Dim tentativas As Integer
    
    tentativas = 0
    Do
        senha = InputBox("Digite a senha:")
        tentativas = tentativas + 1
        
        If senha = "123" Then
            MsgBox "Senha correta!"
            Exit Do
        End If
        
        If tentativas >= 3 Then
            MsgBox "Muitas tentativas! Bloqueado!"
            Exit Do
        End If
    Loop
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Exit:</strong></p>
            <p>• <span class="highlight">Exit For</span> → Sai do loop FOR</p>
            <p>• <span class="highlight">Exit Do</span> → Sai do loop DO</p>
            <p>• Use quando <span class="highlight">encontrar</span> o que procura</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja como sair dos loops antes do fim</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual comando sai de um loop FOR antes do fim?",
        opcoes: [
            "End For",
            "Exit For",
            "Stop For",
            "Break For"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Exit For sai do loop FOR imediatamente.",
        explicacaoErro: "❌ Use Exit For para sair de um loop FOR."
    }
];

// =============================================
// ===== VARIÁVEIS DE CONTROLE =====
// =============================================

let licaoAtual = 0;
let acertos = 0;
let licoesCompletas = 0;
let opcaoSelecionada = -1;

// =============================================
// ===== FUNÇÕES PRINCIPAIS =====
// =============================================

function iniciarCurso() {
    document.getElementById('startBtn').style.display = 'none';
    document.getElementById('stats').style.display = 'grid';
    licaoAtual = 0;
    mostrarLicao();
}

function mostrarLicao() {
    if (licaoAtual >= licoes.length) {
        finalizarCurso();
        return;
    }
    
    const l = licoes[licaoAtual];
    
    document.getElementById('levelBadge').textContent = l.nivel;
    document.getElementById('lessonNumber').textContent = `Lição ${licaoAtual + 1} de ${licoes.length}`;
    document.getElementById('lessonTitle').textContent = l.titulo;
    document.getElementById('conceptText').innerHTML = l.conceito;
    document.getElementById('codeDisplay').textContent = l.codigo;
    document.getElementById('codeExplanation').innerHTML = l.explicacaoCodigo;
    document.getElementById('stepByStep').innerHTML = l.passoAPasso;
    
    document.getElementById('teachingArea').style.display = 'block';
    document.getElementById('quizArea').style.display = 'block';
    document.getElementById('quizQuestion').textContent = l.pergunta;
    
    const container = document.getElementById('optionsContainer');
    container.innerHTML = '';
    l.opcoes.forEach((opcao, index) => {
        const btn = document.createElement('button');
        btn.className = 'option-btn';
        btn.textContent = opcao;
        btn.onclick = function() {
            opcaoSelecionada = index;
            document.querySelectorAll('.option-btn').forEach(b => b.classList.remove('selected'));
            btn.classList.add('selected');
        };
        container.appendChild(btn);
    });
    
    document.getElementById('feedback').style.display = 'none';
    document.getElementById('nextBtn').style.display = 'none';
    opcaoSelecionada = -1;
    
    const progresso = (licaoAtual / licoes.length) * 100;
    document.getElementById('progress').style.width = progresso + '%';
    document.getElementById('progressText').textContent = Math.round(progresso) + '% concluído';
    document.getElementById('acertos').textContent = acertos;
    document.getElementById('licoes').textContent = licoesCompletas;
    document.getElementById('progresso').textContent = Math.round(progresso) + '%';
}

function verificarResposta() {
    if (opcaoSelecionada === -1) {
        alert('Por favor, selecione uma opção!');
        return;
    }
    
    const l = licoes[licaoAtual];
    const feedback = document.getElementById('feedback');
    const feedbackContent = document.getElementById('feedbackContent');
    
    const correta = opcaoSelecionada === l.respostaCorreta;
    const botoes = document.querySelectorAll('.option-btn');
    
    botoes.forEach((btn, i) => {
        btn.classList.remove('selected');
        if (i === l.respostaCorreta) {
            btn.classList.add('correct');
        } else if (i === opcaoSelecionada && !correta) {
            btn.classList.add('wrong');
        }
    });
    
    feedback.style.display = 'block';
    
    if (correta) {
        feedback.className = 'feedback correto';
        feedbackContent.innerHTML = `<h3>✅ Correto!</h3><div class="explicacao"><p>${l.explicacaoAcerto}</p></div>`;
        acertos++;
    } else {
        feedback.className = 'feedback errado';
        feedbackContent.innerHTML = `
            <h3>❌ Não foi dessa vez!</h3>
            <div class="explicacao">
                <p>${l.explicacaoErro}</p>
                <p><strong>Resposta correta:</strong> ${l.opcoes[l.respostaCorreta]}</p>
            </div>
        `;
    }
    licoesCompletas++;
    document.getElementById('acertos').textContent = acertos;
    document.getElementById('licoes').textContent = licoesCompletas;
    document.getElementById('nextBtn').style.display = 'inline-block';
}

function proximaLicao() {
    licaoAtual++;
    mostrarLicao();
}

function finalizarCurso() {
    const total = licoes.length;
    const porcentagem = Math.round((acertos / total) * 100);
    const feedback = document.getElementById('feedback');
    const feedbackContent = document.getElementById('feedbackContent');
    feedback.style.display = 'block';
    feedback.className = 'feedback correto';
    feedbackContent.innerHTML = `
        <h3>🎉 Curso Concluído!</h3>
        <div class="explicacao">
            <p><strong>${porcentagem === 100 ? '🏆 PERFEITO! Você é um mestre do VBA!' : porcentagem >= 80 ? '🌟 Excelente! Você está muito bem!' : porcentagem >= 60 ? '📚 Bom! Continue praticando!' : '💪 Continue estudando! A prática leva à perfeição!'}</strong></p>
            <p>📊 <strong>Aproveitamento:</strong> ${porcentagem}%</p>
            <p>✅ <strong>Acertos:</strong> ${acertos} de ${total}</p>
            <p>📚 <strong>Lições completas:</strong> ${licoesCompletas}</p>
            <br>
            <p><strong>🏁 PARABÉNS POR COMPLETAR O CURSO!</strong></p>
            <p>Agora você sabe VBA do básico ao intermediário!</p>
            <br>
            <button onclick="reiniciarCurso()" class="btn-primary" style="width:100%">
                🔄 Recomeçar Curso
            </button>
        </div>
    `;
    document.getElementById('nextBtn').style.display = 'none';
    document.getElementById('quizArea').style.display = 'none';
}

function reiniciarCurso() {
    acertos = 0;
    licoesCompletas = 0;
    licaoAtual = 0;
    document.getElementById('quizArea').style.display = 'block';
    mostrarLicao();
}

document.addEventListener('DOMContentLoaded', function() {
    console.log('✅ Curso VBA carregado!');
    console.log('📚 Total de lições: ' + licoes.length);
});
