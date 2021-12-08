Attribute VB_Name = "EnviarEmail"
Sub enviarEmail()

Dim olApp As Outlook.Application
Dim janelaDeEmail As Outlook.MailItem
Dim emailAdress As String, comCopia As String
Dim anexo As String
Dim assinatura As String


Set olApp = New Outlook.Application
Set janelaDeEmail = olApp.CreateItem(olMailItem)
emailAdress = InputBox("Ensira os endere�os de email a serem enviados, separe-os com com ponto-e-v�rgula (;).", "DESTINAT�RIO(s) PRINCIPAL(IS)")
comCopia = InputBox("Caso houver destinat�rios em c�pia, ensira os endere�os de e-mail no campo abaixo.", "DESTINAT�RIO(S) EM C�PIA")
anexo = ThisWorkbook.FullName

If emailAdress = "" Then
    MsgBox "Por favor, o campo de endere�o de email do destinat�rio principal � obrigat�rio", vbExclamation, "DESTINAT�RIO DE EMAIL OBRIGAT�RIO"
    Exit Sub
End If

With janelaDeEmail
    ''.Display
    .To = emailAdress
    .CC = comCopia
    .BCC = "eng.milton.soares@gmail.com"
    .Subject = "RELAT�RIO CONTAS A PAGAR"
    assinatura = .HTMLBody
    .HTMLBody = "<div align='center' style='padding: 24px; border: 2px solid #545454'><h1> EMAIL AUTOMATIZADO - TESTE </h1><BR><p style='color: #545454'>Esse � um teste de envio de email pelo <b>Excel</b></p></div>" & assinatura & " "
    .Attachments.Add anexo
    .Send
    
End With

MsgBox "Email enviado com sucesso", vbInformation, "SUCESO!"



End Sub
