<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<%
' ====================================================================
' Script para Excluir Pagamento de Comissão em formato JSON
' ====================================================================

' ATENÇÃO CRÍTICA:
' 1. Certifique-se de que NÃO HÁ NENHUM CARACTERE (espaço, quebra de linha, BOM)
'    ANTES desta primeira linha (<%@LANGUAGE...).
' 2. Certifique-se de que o arquivo 'conexao.asp' TAMBÉM NÃO CONTÉM NENHUM CARACTERE
'    fora da definição de StrConn. Ele deve ser puramente VBScript.

' Ativa o buffer de resposta. Isso impede que o servidor envie qualquer coisa
' antes que o script termine de processar e o Response.End seja chamado.
Response.Buffer = True

' Limpa qualquer conteúdo que possa ter sido adicionado ao buffer antes deste ponto.
Response.Clear

' Define o tipo de conteúdo como JSON. Esta é a primeira coisa que o navegador deve ver.
Response.ContentType = "application/json"

' Variável para armazenar a mensagem de log/resultado
Dim jsonResponse

' Verifica se o método da requisição é POST.
If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    jsonResponse = "{""success"": false, ""message"": ""Método não permitido. Use POST para esta operação.""}"
    Response.Write jsonResponse
    Response.End ' Encerra a execução e envia o conteúdo do buffer
End If

' Obtém o ID do pagamento do formulário (enviado pelo JavaScript como ID_Pagamento).
Dim paymentId
paymentId = Request.Form("ID_Pagamento")

' Adiciona log sobre o ID recebido
Dim logMessage
logMessage = "Recebido ID_Pagamento: " & paymentId & ". "

' Valida o ID: verifica se não está vazio e se é um número.
If paymentId = "" Or Not IsNumeric(paymentId) Then
    jsonResponse = "{""success"": false, ""message"": """ & logMessage & "ID de pagamento inválido ou não fornecido.""}"
    Response.Write jsonResponse
    Response.End ' Encerra a execução e envia o conteúdo do buffer
End If

' Habilita o tratamento de erro para capturar problemas de banco de dados.
On Error Resume Next

Dim conn, sql
Set conn = Server.CreateObject("ADODB.Connection")

' Adiciona log antes de tentar abrir a conexão
logMessage = logMessage & "Tentando abrir conexão com o banco de dados. "

' Tenta abrir a conexão com o banco de dados.
conn.Open StrConnSales

' Verifica se ocorreu um erro ao abrir a conexão.
If Err.Number <> 0 Then
    jsonResponse = "{""success"": false, ""message"": """ & logMessage & "Erro ao conectar ao banco de dados: " & Replace(Err.Description, """", "'") & """}"
    Response.Write jsonResponse
    Response.End ' Encerra a execução e envia o conteúdo do buffer
End If

sql = "DELETE FROM PAGAMENTOS_COMISSOES WHERE ID_Pagamento = " & paymentId

' Adiciona log antes de executar a query
logMessage = logMessage & "Query SQL completa para exclusão: [" & sql & "]. "

' Executa a exclusão diretamente.
conn.Execute sql

' Verifica se ocorreu um erro durante a execução da exclusão.
If Err.Number <> 0 Then
    jsonResponse = "{""success"": false, ""message"": """ & logMessage & "Erro ao excluir pagamento: " & Replace(Err.Description, """", "'") & """}"
Else
    jsonResponse = "{""success"": true, ""message"": """ & logMessage & "Pagamento excluído com sucesso.""}"
End If

' Desabilita o tratamento de erro.
On Error GoTo 0

' Fecha a conexão com o banco de dados.
If IsObject(conn) Then
    conn.Close
    Set conn = Nothing
End If

' Escreve a resposta JSON final no buffer.
Response.Write jsonResponse

' Encerra a execução do script e envia o conteúdo do buffer.
' Esta é a linha mais CRÍTICA para garantir que nada mais seja enviado.
Response.End
%>
