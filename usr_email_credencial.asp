<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->


<%
' --- Tratamento de Erros Global para Debug ---
On Error Resume Next ' Habilita tratamento de erros. Desabilite em producao ou use tratamento mais robusto.

' --- 1. Obter o UserID da URL ---
Dim UserID
UserID = Request.QueryString("UserID")

If UserID = "" Or Not IsNumeric(UserID) Then
    Response.Write "<p style='color: red;'>Erro: O parâmetro UserID é obrigatório e deve ser um número.</p>"
    Response.End ' Para a execução do script
End If

' --- 2. Conectar ao Banco de Dados e Buscar Credenciais ---
Dim objConn, rsUser
Set objConn = Server.CreateObject("ADODB.Connection")

' Tenta abrir a conexao
objConn.Open StrConn

If Err.Number <> 0 Then
    Response.Write "<p style='color: red;'>Erro ao conectar ao banco de dados: " & Err.Description & "</p>"
    ' Opcional: Logar o erro em um arquivo para depuracao mais detalhada.
    ' Call LogError("Erro na conexao: " & Err.Description)
    Set objConn = Nothing
    Err.Clear
    Response.End
End If

' Prepara e executa a consulta para buscar os dados do usuario
Dim sqlQuery

sqlQuery = "SELECT * FROM Usuarios WHERE UserID = " & UserID

Set rsUser = objConn.Execute(sqlQuery)

If Err.Number <> 0 Then
    Response.Write "<p style='color: red;'>Erro ao executar a consulta: " & Err.Description & "</p>"
    Set rsUser = Nothing
    objConn.Close
    Set objConn = Nothing
    Err.Clear
    Response.End
End If

' --- 3. Verificar se o Usuário Existe e Possui E-mail ---
If rsUser.EOF Then
    Response.Write "<p>Usuário com ID " & UserID & " não encontrado no sistema.</p>"
Else
    Dim userEmail, userName, userPassword
    userEmail = rsUser("Email")
    userName = LCase(rsUser("Usuario"))
    userPassword = rsUser("SenhaAtual") ' Pega a senha diretamente do DB

    If Trim(userEmail) = "" Then
        Response.Write "<p>O usuário '" & userName & "' (ID: " & UserID & ") não possui um e-mail cadastrado. Não foi possível enviar as credenciais.</p>"
    Else
        ' --- 4. Enviar o E-mail usando CDONTS.NewMail ---
        Dim objMail
        Set objMail = Server.CreateObject("CDONTS.NewMail")

        ' Remetente
        objMail.From = "sendmail@gabnetweb.com.br"

        ' Destinatário principal (o e-mail do usuario)
        objMail.To = userEmail

        ' Cópia para Valter Barreto, como no seu exemplo
        objMail.Cc = "valterpb@hotmail.com"

        ' Assunto do E-mail
        objMail.Subject = "SUNNYIMOB - Credenciais de Acesso" ' Assunto mais direto para as credenciais

        ' Corpo do E-mail (TEXTO SIMPLES)
        ' **MUDANÇA 1: MailFormat = 0 para texto puro**
        objMail.MailFormat = 0 
        ' **MUDANÇA 2: Usar vbCrLf para quebras de linha e remover tags HTML**
        objMail.Body = "Olá, você está recebendo a credencial para acessar o sistema SUNNYIMOB." & vbCrLf & vbCrLf & _
                       "Usuário: " & userName & vbCrLf & _
                       "Senha: " & userPassword & vbCrLf & vbCrLf & _
                       "Qualquer dúvida, entrar em contato com Valter Barreto - 81 98842-1455 (WhatsApp)." & vbCrLf & vbCrLf & _
                       "Atenciosamente," & vbCrLf & _
                       "Equipe SUNNYIMOB"
        
        ' Envia o e-mail
        objMail.Send

        If Err.Number <> 0 Then
            Response.Write "<p style='color: red;'>Erro ao enviar o e-mail: " & Err.Description & "</p>"
            Err.Clear
        Else
            Response.Write "<p>E-mail com credenciais enviado com sucesso para: <strong>" & userEmail & "</strong></p>"
            Response.Write "<p>Usuário: " & userName & "</strong></p>"
            Response.Write "<p>Pwr: " & userPassword & "</strong></p>"
            Response.Write "<p>Uma cópia também foi enviada para <strong>valterpb@hotmail.com</strong>.</p>"
        End If

        ' Limpar objeto CDONTS.NewMail
        Set objMail = Nothing
    End If
End If

' --- 5. Fechar Conexão com o Banco de Dados ---
If Not rsUser Is Nothing Then
    If rsUser.State = 1 Then rsUser.Close
    Set rsUser = Nothing
End If
If Not objConn Is Nothing Then
    If objConn.State = 1 Then objConn.Close
    Set objConn = Nothing
End If

On Error GoTo 0 ' Desabilita tratamento de erros (volta ao padrao)
%>