Attribute VB_Name = "Módulo1"
Sub Email()


    'função para abrir o Outlook.
    
Set objeto_outlook = CreateObject("Outlook.Application")

    'Quantidades de E-mails a serem enviados iniciando na linha 11 ate o valor presende na celula B6.

For linha = 11 To Cells(6, 2).Value

    'Função para se criar um novo E-mail.

    Set Email = objeto_outlook.createitem(0)
    
    'Função para abrir E-mail.
    
    Email.display
    
    'Função para quem os e-mails serão enviados.
    
    Email.To = Cells(linha, 3).Value
    Email.cc = Cells(2, 2).Value
    Email.bcc = Cells(3, 2).Value
    
    'Função para assunto do E-mail - Cada Linha gera um assunto diferente.
    
    Email.Subject = Cells(linha, 4).Value
    
    'Texto do corpo do E-mail.
    
    ' Dicas de Formatação.
    ' <br> quebra de paragrafo.
    ' <b> .Texto. </b> Negrito.
    ' <i> .Texto. </i> itálico.
    ' <u> .Texto. </u> Sublinhado.
    
    'Adição de Fotos e Imagens
    ' Recomenda-se utilizar imagens com no máximo 1600pix Site para converter(https://www.easy-resize.com/pt/)
    ' <"<img src='Local do arquivo.jpg'>" para se anexa imagens obrigatoriamente deve-se utilizar_
    ' a Macro "com anexo"
    ' Cells(Linha, 2).Value & ", " & Cells(4, 2) & "<br> <br>"
    '& "Prezado(a) " & Cells(Linha, 2).Value & ", " & Cells(6, 5) & "<br> <br>"
    
   Email.HTMLBody = "<BODY style=font-size:12pt;font-family:Calibri> " _
    & "Belo Horizonte, " & Cells(5, 3) & "." & "<br> <br>" _
    & "Prezado(a) " & Cells(linha, 2).Value & ", " & Cells(6, 5) & "<br> <br>" _
    & "Como é de conhecimento, à ABIH/MG está desenvolvendo a Cesta Competitiva de hotéis da Grande BH." & "<br> <br>" _
    & "Com o intuito de respeitar a Lei Geral de Proteção de Dados (LGPD) e assegurar que os dados compartilhados no grupo serão disponibilizados apenas com os membros do grupo, solicitamos a assinatura do termo de autorização de dados, em anexo." & "<br> <br>" _
    & "Atenciosamente." & Email.HTMLBody


    'Função para anexar um arquivo.
    
    'Se deseja remover a função "anexo" favor adicionar aspas ( ' ) unica na frente da função.
    Email.Attachments.Add ("C:\Users\Guilh\Desktop\Anexo\Pedido de Assinatura no Termo de Autorização de dado.pdf")
    
    
    Email.send
Next
    

End Sub

