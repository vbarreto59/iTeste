<%
On Error Resume Next ' Habilita o tratamento de erros

' Constantes para strings de conexão
Const LOCAL_CONNECTION_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\inetpub\wwwroot\iTeste\db\ImobVendas.mdb;"
Const PRODUCTION_CONNECTION_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\ClientHome\gabnetweb.com.br\bdados\ImobVendas.mdb;"

' Define a string de conexão com base no ambiente
Dim StrConn, connectionString
Set StrConn = Server.CreateObject("ADODB.Connection")

If Request.ServerVariables("SERVER_NAME") = "localhost" Then
    connectionString = LOCAL_CONNECTION_STRING
Else
    connectionString = PRODUCTION_CONNECTION_STRING
End If

'response.Write connectionString
'response.end()

' Tenta abrir a conexão
StrConn.Open connectionString

' Verifica se houve erro na conexão
If Err.Number <> 0 Then
    Response.Write("Erro ao conectar ao banco de dados (conexao.asp). Por favor, tente novamente mais tarde. Erro: " & Err.Number )
    StrConn.Close ' Fechar a conexão em caso de erro
    Set StrConn = Nothing ' Limpar o objeto de conexão
    Response.End()
End If

' Aqui pode ir o código que interage com o banco de dados...

' Fechamento da conexão
If Not StrConn Is Nothing Then
    StrConn.Close
    Set StrConn = Nothing ' Limpar o objeto de conexão
End If

On Error GoTo 0 ' Desabilita o tratamento de erros
StrConn = connectionString
%>
