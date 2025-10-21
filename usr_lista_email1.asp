<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp" -->

<%
' ====================================================================================
' Configurações do E-mail
' ====================================================================================
Dim EmailRemetente, EmailDestinatario, AssuntoEmail, NomeRemetente
EmailRemetente = "sendmail@gabnetweb.com.br" ' Remetente conforme seu exemplo
EmailDestinatario = "sendmail@gabnetweb.com.br, valterpb@hotmail.com" ' Destinatários conforme seu exemplo

' O nome do grupo que você deseja filtrar
Dim GroupToFilter
GroupToFilter = "TOCCA_ONZE"

' Assunto do E-mail, incluindo o grupo
AssuntoEmail = "Lista de Usuarios do Grupo " & UCase(GroupToFilter) & " - " & Date() & " - " & Time()

' ====================================================================================
' BLOC0 VBSCRIPT: Busca e formata a lista de usuários (texto simples)
' ====================================================================================
Dim objConn, rsUsers, strTextBody, strSQL

Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open StrConn

' SQL para buscar usuários do grupo "TOCCA_ONZE" em ordem alfabética
strSQL = "SELECT Usuarios.Usuario, Usuarios.Nome, Usuarios.Ativo, Grupo.Nome_Grupo, Usuarios.Email, Usuarios.Telefones " & _
         "FROM Grupo INNER JOIN (Usuarios INNER JOIN Usuario_Grupo ON Usuarios.UserId = Usuario_Grupo.UserId) ON Grupo.ID_Grupo = Usuario_Grupo.ID_Grupo " & _
         "WHERE Grupo.Nome_Grupo = '" & Replace(GroupToFilter, "'", "''") & "' " & _
         "ORDER BY Usuarios.Usuario ASC;"

Set rsUsers = objConn.Execute(strSQL)

' Inicia a construção do corpo de texto simples do e-mail
strTextBody = "Prezado(a) administrador(a)," & vbCrLf & vbCrLf & _
              "Abaixo está a lista atualizada dos usuários pertencentes ao grupo '" & GroupToFilter & "':" & vbCrLf & vbCrLf

If Not rsUsers.EOF Then
    cont = 0
    Do While Not rsUsers.EOF
        cont = cont + 1
        Dim statusTexto
        If CBool(rsUsers("Ativo")) Then
            statusTexto = "ATIVO"
        Else
            statusTexto = "INATIVO"
        End If
        nome = rsUsers("Nome")

        strTextBody = strTextBody & "----------------------------------------" & vbCrLf & _
                                   "Usuário: " & cont & "-" & UCase(rsUsers("Usuario")) & vbCrLf & _
                                   "Nome: " & nome & vbCrLf & _
                                   "Status: " & statusTexto & vbCrLf & _
                                   "Grupo: " & rsUsers("Nome_Grupo") & vbCrLf & _
                                   "Email: " & rsUsers("Email") & vbCrLf & _
                                   "Telefone(s): " & rsUsers("Telefones") & vbCrLf
        rsUsers.MoveNext
    Loop
    strTextBody = strTextBody & "----------------------------------------" & vbCrLf & vbCrLf
Else
    strTextBody = strTextBody & "Nenhum usuário encontrado no grupo '" & GroupToFilter & "'." & vbCrLf & vbCrLf
End If

strTextBody = strTextBody & "Atenciosamente," & vbCrLf & _
                           "A equipe do Sunny." & vbCrLf & vbCrLf & _
                           "© " & Year(Now()) & " Sunny. Todos os direitos reservados."

rsUsers.Close
Set rsUsers = Nothing
objConn.Close
Set objConn = Nothing

' ====================================================================================
' BLOC0 VBSCRIPT: Envio do E-mail usando CDONTS.NewMail
' ====================================================================================
Dim objMail, bMailSent

On Error Resume Next ' Habilita tratamento de erro para o envio de e-mail

' Verifica se não é uma requisição local (conforme seu exemplo)
If (Request.ServerVariables("remote_addr") <> "127.0.0.1") AND (Request.ServerVariables("remote_addr") <> "::1") Then
    Set objMail = Server.CreateObject("CDONTS.NewMail")

    objMail.From = EmailRemetente
    objMail.To   = EmailDestinatario
    objMail.Subject = AssuntoEmail
    objMail.Body = strTextBody ' Define o corpo do e-mail como texto simples
    objMail.MailFormat = 0     ' 0 = Texto Simples, 1 = HTML

    objMail.Send ' Tenta enviar o e-mail
    bMailSent = (Err.Number = 0) ' Verifica se não houve erro
    Set objMail = Nothing
Else
    ' Se a requisição for local, não tenta enviar o e-mail
    bMailSent = True ' Simula sucesso para não exibir erro em ambiente de desenvolvimento local
End If

If Not bMailSent Then
    ' Captura e exibe o erro se o envio falhar
    Response.Write "<!DOCTYPE html><html lang='pt-br'><head><meta charset='UTF-8'><title>Erro no Envio de E-mail</title><link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css'></head><body><div class='container mt-5'><div class='alert alert-danger'><h4 class='alert-heading'>Erro ao Enviar E-mail!</h4><p>Ocorreu um erro ao tentar enviar a lista de usuários para <strong>" & Server.HTMLEncode(EmailDestinatario) & "</strong>.</p><hr><p class='mb-0'><strong>Detalhes do Erro:</strong> " & Err.Description & " (Código: " & Err.Number & ")</p><p>Verifique as configurações do serviço SMTP no seu servidor IIS.</p><br><a href='usr_listar.asp' class='btn btn-danger'>Voltar para a Lista de Usuários</a></div></div></body></html>"
Else
    ' Exibe mensagem de sucesso
    Response.Write "<!DOCTYPE html><html lang='pt-br'><head><meta charset='UTF-8'><title>E-mail Enviado</title><link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css'></head><body><div class='container mt-5'><div class='alert alert-success'><h4 class='alert-heading'>Sucesso!</h4><p>A lista de usuários do grupo <strong>" & Server.HTMLEncode(GroupToFilter) & "</strong> foi enviada com sucesso para <strong>" & Server.HTMLEncode(EmailDestinatario) & "</strong>.</p><br><a href='usr_listar.asp' class='btn btn-success'>Voltar para a Lista de Usuários</a></div></div></body></html>"
End If

On Error GoTo 0 ' Desabilita tratamento de erro
%>