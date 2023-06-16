# Função InputBox

Exemplo 

Sub c()
Dim coluna As String
Dim linha As Integer
MsgBox "teste", vbApplicationModal
coluna = InputBox("teste")
linha = 7
Range("I7").Value = Range(coluna & linha).Value
End Sub


# Sintaxe
InputBox(prompt, [ title ], [ default ], [ xpos ], [ ypos ], [ helpfile, context ] )

A sintaxe da função InputBox possui estes argumentos nomeados:

prompt	Obrigatório. Expressão de cadeia de caracteres exibida como a mensagem na caixa de diálogo. O comprimento máximo de prompt é aproximadamente 1024 caracteres, dependendo da largura dos caracteres usados. Se o prompt consistir em mais de uma linha, você poderá separar as linhas usando um caractere de retorno de carruagem (Chr(13)), um caractere de linefeed (Chr(10)) ou uma combinação de caracteres return-linefeed ((Chr(13) & (Chr(10)) entre cada linha.

title	Opcional. Cadeia de caracteres exibida na barra de título da caixa de diálogo de expressão. Se você omitir title, o nome do aplicativo é colocado na barra de título.

default	Opcional. Expressão de cadeia de caracteres exibida na caixa de texto como a resposta padrão, caso nenhuma outra entrada seja fornecida. Se você omitir padrão, a caixa de texto será exibida vazia.

xpos	Opcional. Expressão numérica que especifica, em twips, a distância horizontal da borda esquerda da caixa de diálogo da borda esquerda da tela. Se xpos for omitido, a caixa de diálogo será centralizada horizontalmente.

ypos	Opcional. Expressão numérica que especifica, em twips, a distância vertical da borda superior da caixa de diálogo da parte superior da tela. Se ypos for omitido, a caixa de diálogo será posicionada verticalmente a um terço da altura da tela, de cima para baixo.

helpfile	Opcional. Expressão de cadeia de caracteres que identifica o arquivo de ajuda a usar para oferecer ajuda contextual para a caixa de diálogo. Se helpfile for fornecido, context também deve ser fornecido.

context	Opcional. Expressão numérica que é o número de contexto da Ajuda atribuído ao tópico da Ajuda apropriado pelo autor da Ajuda. Se context for fornecido, helpfile também deve ser fornecido.


