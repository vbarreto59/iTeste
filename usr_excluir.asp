<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp" -->

<%


' Obtém o ID do usuário a ser excluído
Dim UserId
UserId = Request.QueryString("Id")

' Validação básica
If UserId = "" Or Not IsNumeric(UserId) Then
    Session("MsgErro") = "ID do usuário inválido!"
    Response.Redirect("usr_listar.asp")
End If

' Verifica se o usuário existe antes de excluir
Dim rsVerifica
Set rsVerifica = Server.CreateObject("ADODB.Recordset")
rsVerifica.Open "SELECT Usuario FROM Usuarios WHERE UserId = " & UserId, StrConn

If rsVerifica.EOF Then
    Session("MsgErro") = "Usuário não encontrado!"
    Response.Redirect("usr_listar.asp")
End If

Dim usuarioExcluido
usuarioExcluido = rsVerifica("Usuario")
rsVerifica.Close()
Set rsVerifica = Nothing

' Executa a exclusão
Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = StrConn
cmd.CommandText = "DELETE FROM Usuarios WHERE UserId = ?"
cmd.Parameters.Append cmd.CreateParameter("UserId", 5, 1, , UserId)  ' adDouble

On Error Resume Next
cmd.Execute

If Err.Number <> 0 Then
    Session("MsgErro") = "Erro ao excluir usuário: " & Err.Description
    Response.Redirect("usr_listar.asp")
End If

On Error GoTo 0

Response.Redirect("usr_listar.asp")
%>