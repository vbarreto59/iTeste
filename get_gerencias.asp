<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->

<%
Response.ContentType = "application/json"

Dim diretoriaId
diretoriaId = Request.QueryString("diretoriaId")

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT GerenciaID, NomeGerencia FROM Gerencias WHERE DiretoriaId = " & diretoriaId & " ORDER BY NomeGerencia", conn

Dim json
json = "["
Do While Not rs.EOF
    If json <> "[" Then json = json & ","
    json = json & "{""GerenciaID"":" & rs("GerenciaID") & ",""NomeGerencia"":""" & Replace(rs("NomeGerencia"), """", "\""") & """}"
    rs.MoveNext
Loop
json = json & "]"

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing

Response.Write json
%>