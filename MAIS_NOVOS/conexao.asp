<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 18/12/2025               -->
<!-- CODIGO_ARQUIVO: OCCHQKOKOZ          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<%
' ===================================================================================
' CONEXÃO DE BANCO DE DADOS DINÂMICA (ASP CLASSIC)
' 27 11 2025'
' Este script detecta o ambiente (Local ou Produção) para definir a Connection String.
' ===================================================================================

On Error Resume Next ' Habilita o tratamento de erros

' -----------------------------------------------------------
' 1. DEFINIÇÕES DE CAMINHO
' -----------------------------------------------------------

'/httpdocs/iTeste/db
Dim PRODUCTION_ABSOLUTE_PATH
PRODUCTION_ABSOLUTE_PATH = "E:/ClientHome/gabnetweb.com.br/httpdocs/iTeste/db/ImobVendas.mdb"

Dim LOCAL_VIRTUAL_PATH
LOCAL_VIRTUAL_PATH = "db/ImobVendas.mdb" ' Assumindo que /db/ImobVendas.mdb é o local no localhost

' Provedor OLEDB (Jet 4.0 para .mdb)
Dim PROVIDER
PROVIDER = "Provider=Microsoft.Jet.OLEDB.4.0;"

' -----------------------------------------------------------
' 2. DETECÇÃO DE AMBIENTE E DEFINIÇÃO DA CONNECTION STRING
' -----------------------------------------------------------

Dim serverName
serverName = LCase(Request.ServerVariables("SERVER_NAME")) ' Pega o nome do servidor em letras minúsculas

Dim connectionPath
Dim connectionString

' Lógica de Detecção:
If InStr(serverName, "gabnetweb.com.br") > 0 Then
   connectionPath = PRODUCTION_ABSOLUTE_PATH
Else
   connectionPath = Server.MapPath(LOCAL_VIRTUAL_PATH)
End If

' Monta a Connection String final
connectionString = PROVIDER & "Data Source=" & connectionPath & ";"

' -----------------------------------------------------------
' 3. ESTABELECIMENTO E TESTE DA CONEXÃO
' -----------------------------------------------------------

' Cria o objeto de conexão ADODB.
Dim StrConn
Set StrConn = Server.CreateObject("ADODB.Connection")

' Tenta abrir a conexão
StrConn.Open connectionString


' Verifica se houve erro na conexão
If Err.Number <> 0 Then
Response.Write("<h3>[ERRO DE CONEXÃO]</h3>")
Response.Write("<p>Não foi possível conectar ao banco de dados.</p>")
Response.Write("<p><strong>Connection String:</strong> " & connectionString & "</p>")
Response.Write("<p><strong>Ambiente Detectado:</strong> " & serverName & "</p>")
Response.Write("<p><strong>Caminho Utilizado:</strong> " & connectionPath & "</p>")
Response.Write("<p><strong>Erro VBScript:</strong> " & Err.Description & " (" & Err.Number & ")</p>")

' Limpeza de objetos e encerramento da execução
If Not StrConn Is Nothing Then StrConn.Close
   Set StrConn = Nothing
   Response.End()
End If

' Se chegou aqui, a conexão está OK!
'Response.Write "<!-- Conexão estabelecida com sucesso! -->" ' Comentário HTML para debug
'Response.Write "Conexão OK. Servidor: " & serverName ' Descomente para debug

On Error GoTo 0 ' Desabilita o tratamento de erros
connectionString = PROVIDER & "Data Source=" & connectionPath & ";"
StrConn = connectionString
%>