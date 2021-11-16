Attribute VB_Name = "M�dulo1"
Sub Email()


    'fun��o para abrir o Outlook.
    
Set objeto_outlook = CreateObject("Outlook.Application")

    'Quantidades de E-mails a serem enviados iniciando na linha 11 ate o valor presende na celula B6.

For linha = 11 To Cells(6, 2).Value

    'Fun��o para se criar um novo E-mail.

    Set Email = objeto_outlook.createitem(0)
    
    'Fun��o para abrir E-mail.
    
    Email.display
    
    'Fun��o para quem os e-mails ser�o enviados.
    
    Email.To = Cells(linha, 3).Value
    Email.cc = Cells(2, 2).Value
    Email.bcc = Cells(3, 2).Value
    
    'Fun��o para assunto do E-mail - Cada Linha gera um assunto diferente.
    
    Email.Subject = Cells(linha, 4).Value
    
    'Texto do corpo do E-mail.
    
    ' Dicas de Formata��o.
    ' <br> quebra de paragrafo.
    ' <b> .Texto. </b> Negrito.
    ' <i> .Texto. </i> it�lico.
    ' <u> .Texto. </u> Sublinhado.
    
    'Adi��o de Fotos e Imagens
    ' Recomenda-se utilizar imagens com no m�ximo 1600pix Site para converter(https://www.easy-resize.com/pt/)
    ' <"<img src='Local do arquivo.jpg'>" para se anexa imagens obrigatoriamente deve-se utilizar_
    ' a Macro "com anexo"
    ' Cells(Linha, 2).Value & ", " & Cells(4, 2) & "<br> <br>"
    '& "Prezado(a) " & Cells(Linha, 2).Value & ", " & Cells(6, 5) & "<br> <br>"
    
   Email.HTMLBody = "<BODY style=font-size:12pt;font-family:Calibri> " _
    & "Belo Horizonte, " & Cells(5, 3) & "." & "<br> <br>" _
    & "Prezado(a) " & Cells(linha, 2).Value & ", " & Cells(6, 5) & "<br> <br>" _
    & "Como � de conhecimento, � ABIH/MG est� desenvolvendo a Cesta Competitiva de hot�is da Grande BH." & "<br> <br>" _
    & "Com o intuito de respeitar a Lei Geral de Prote��o de Dados (LGPD) e assegurar que os dados compartilhados no grupo ser�o disponibilizados apenas com os membros do grupo, solicitamos a assinatura do termo de autoriza��o de dados, em anexo." & "<br> <br>" _
    & "Atenciosamente." & Email.HTMLBody


    'Fun��o para anexar um arquivo.
    
    'Se deseja remover a fun��o "anexo" favor adicionar aspas ( ' ) unica na frente da fun��o.
    Email.Attachments.Add ("C:\Users\Guilh\Desktop\Anexo\Pedido de Assinatura no Termo de Autoriza��o de dado.pdf")
    
    
    Email.send
Next
    

End Sub

