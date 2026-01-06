<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: DJTGQUCZJY          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/dbcrud.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	  <title>GabNet  - Login</title>	
	  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	  <meta http-equiv="X-UA-Compatible" content="IE=edge">  
	  <meta charset="utf-8">
	</head>
<body>

<%
' Verifique se já foi enviado um e-mail hoje

Dim hoje
hoje = Date
  
' Conecte-se ao banco de dados (exemplo usando o Access)
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open MM_dbcrud_STRING

Dim sql
sql = "SELECT * FROM LoginsDiarios WHERE Data = ?"

Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")

' Crie e configure um objeto de comando
Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandText = sql
cmd.CommandType = 1 ' AdCmdText (Consulta de texto)

' Adicione um parâmetro à consulta
cmd.Parameters.Append cmd.CreateParameter("DataParam", 7, 1, , hoje) ' 7 = adDate, 1 = adParamInput

' Associe o comando ao objeto de registro
Set rs = cmd.Execute
If rs.EOF Then
   vPrimeiro = True
Else
   vPrimeiro = False	
End if   
'If rs.EOF Then
If vPrimeiro Then	
    ' Nenhum e-mail foi enviado hoje, envie o e-mail e registre o login
    ' Envie o e-mail com o anexo log.txt para valterpb@gmail.com aqui
	
	User = Session("MM_Username")
	sql = "INSERT INTO LoginsDiarios (Usuario, Enviado) VALUES (?, ?)"
		
	' Crie e configure um novo comando para a inserção
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = conn
	cmd.CommandText = sql
	cmd.CommandType = 1 ' AdCmdText (Consulta de texto)
	
	' Adicione parâmetros à consulta de inserção
	cmd.Parameters.Append cmd.CreateParameter("Param1", 200, 1, 255, User) ' 200 = adVarChar, 1 = adParamInput
	cmd.Parameters.Append cmd.CreateParameter("Param2", 11, 1, 255, True) ' 11 = adBoolean, 1 = adParamInput
	
	' Execute a consulta de inserção
	cmd.Execute	
	' Feche a conexão com o banco de dados
	rs.Close
	conn.Close
	Set rs = Nothing
	Set conn = Nothing	
	'------------------------------------------------------------------------------------

	If Request.ServerVariables("SERVER_NAME") <> "localhost" Then
		On Error Resume Next
		
		set objMail = server.createobject("CDONTS.NewMail")
		
		If Err.Number <> 0 Then
			Response.Write "Erro ao criar o objeto de email: " & Err.Description
			Response.End
		End If
	
		objMail.From = "sendmail@gabnetweb.com.br"
		objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
		objMail.Subject = "SV-1o Login e Backup. Usuário: " & UCase(Session("MM_Username"))
		objMail.Body  = "Login: " & UCase(Session("MM_Username"))
		objMail.MailFormat = 0
		objMail.Attachfile "E:\ClientHome\gabnetweb.com.br\bdados\SunSales.mdb", "SunSales.mdb"		
		
		objMail.Send
	
		If Err.Number <> 0 Then
			Response.Write "Erro ao enviar o email: " & Err.Description
			'Response.End()
		Else
			Response.Write "Email enviado com sucesso."
			'Response.End()
		End If
	
		set objMail = Nothing
	End if


   ' Registre o login na tabela LoginsDiarios
Else
   'A partir do segundo login'
	If Request.ServerVariables("SERVER_NAME") <> "localhost" Then

		On Error Resume Next

		
		set objMail = server.createobject("CDONTS.NewMail")
		
		If Err.Number <> 0 Then
			Response.Write "Erro ao criar o objeto de email: " & Err.Description
			Response.End
		End If
	
		objMail.From = "sendmail@gabnetweb.com.br"
		objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
		objMail.Subject = "GN-Login: " & UCase(Session("MM_Username"))
		objMail.Body  = "Login: " & UCase(Session("MM_Username"))
		objMail.MailFormat = 0
		'objMail.Attachfile "E:\ClientHome\gabnetweb.com.br\bdados\gabnet2017.mdb", "gabnet2017.mdb"		
		
		objMail.Send

	' Feche a conexão com o banco de dados
   end if
	rs.Close
	conn.Close
	Set rs = Nothing
	Set conn = Nothing
End If


response.Redirect("mainmenu.asp")
%>
</body>
</html>
