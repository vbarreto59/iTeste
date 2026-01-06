<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 18/12/2025               -->
<!-- CODIGO_ARQUIVO: ZHCNMYFDNX          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<%
Dim conn, connSales

Set conn = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")

conn.Open StrConn        ' Diretorias / Gerencias
connSales.Open StrConnSales  ' MetasGerencia / MetasDiretoria / MetaEmpresa
%>
