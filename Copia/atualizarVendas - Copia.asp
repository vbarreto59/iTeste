<%
on error resume next

' Verifica se as connection strings já foram definidas.
' Se não, inclui os arquivos para defini-las.
If Len(StrConn) = 0 Then
    ' ATENÇÃO: A sintaxe de inclusão correta é esta, sem o ' na frente
    ' A linha abaixo inclui o arquivo "conexao.asp"
    <!--#include file="conexao.asp"-->
End If

If Len(StrConnSales) = 0 Then
    ' A linha abaixo inclui o arquivo "conSunSales.asp"
    <!--#include file="conSunSales.asp"-->
End If

' Conexão com o banco de dados principal
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

' Conexão com o banco de dados de Vendas
Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' O caminho completo para o banco de dados original é necessário para o comando IN
' Usando a StrConn, extraímos o caminho para dbSunnyPath.
Dim dbSunnyPath
dbSunnyPath = Split(StrConn, "Data Source=")(1)
dbSunnyPath = Left(dbSunnyPath, InStr(dbSunnyPath, ";") - 1)

' O caminho completo para o banco de dados de Vendas pode ser extraído de forma similar,
' mas não é necessário para esta lógica de JOIN.
' Dim dbSalesPath
' dbSalesPath = Split(StrConnSales, "Data Source=")(1)
' dbSalesPath = Left(dbSalesPath, InStr(dbSalesPath, ";") - 1)


' Comando de atualização de nomes de diretores na tabela Vendas
' Esta query também usa a cláusula IN para acessar Diretorias e Usuarios
' Foi ajustado para usar a sintaxe de [;DATABASE=...].
Dim sql3, affectedRows3
sql3 = "UPDATE (Vendas INNER JOIN [;DATABASE=" & dbSunnyPath & "].Diretorias " & _
       "ON Vendas.DiretoriaId = Diretorias.DiretoriaId) " & _
       "INNER JOIN [;DATABASE=" & dbSunnyPath & "].Usuarios " & _
       "ON Diretorias.UserId = Usuarios.UserId " & _
       "SET Vendas.UserIdDiretoria = Usuarios.UserId, " & _
       "Vendas.NomeDiretor = Usuarios.Nome;"
connSales.Execute sql3, affectedRows3

' Comando de atualização de nomes de gerentes na tabela Vendas
' Mesma lógica, usando a cláusula IN para Gerencias e Usuarios
' Foi ajustado para usar a sintaxe de [;DATABASE=...].
Dim sql4, affectedRows4
sql4 = "UPDATE (Vendas INNER JOIN [;DATABASE=" & dbSunnyPath & "].Gerencias " & _
       "ON Vendas.GerenteId = Gerencias.GerenciaId) " & _
       "INNER JOIN [;DATABASE=" & dbSunnyPath & "].Usuarios " & _
       "ON Gerencias.UserId = Usuarios.UserId " & _
       "SET Vendas.UserIdGerencia = Usuarios.UserId, " & _
       "Vendas.NomeGerente = Usuarios.Nome;"
connSales.Execute sql4, affectedRows4

' Os comandos sql5 comentados no seu código original estão aqui para referência:
' sql5 = "UPDATE Vendas SET Vendas.ValorComissao = [ValorUnidade]*([ComissaoPercentual]/100);"
' connSales.Execute sql5, affectedRows5

Response.Write "Atualização Concluída!"

' Fecha ambas as conexões
conn.Close
Set conn = Nothing

connSales.Close
Set connSales = Nothing
%>
<%

'======== Executar querys dentro do banco - 14 08 2025 =========='
Set rsAjustaData = Server.CreateObject ("ADODB.Command")
rsAjustaData.ActiveConnection = StrConnSales
rsAjustaData.CommandType = 4 'adCmdStoredProc

'Limpa a tabela de resumo ComissaoSaldo
rsAjustaData.CommandText = "qryDelComissao"
rsAjustaData.Execute()

'------------ corretores '
rsAjustaData.CommandText = "qryAddComisVendaCorretor"
rsAjustaData.Execute()

'------------ diretores '
'incluir as vendas com as respectivas comissões'
rsAjustaData.CommandText = "qryAddComisVendaDiretor"
rsAjustaData.Execute()

'------------ Gerentes '
'incluir as vendas com as respectivas comissões'
rsAjustaData.CommandText = "qryAddComisVendaGerente"
rsAjustaData.Execute()



' adiciona todas as comissoes pagas '
rsAjustaData.CommandText = "qryAddComisPaga"
rsAjustaData.Execute()

'Atualizar Nomes de Comissao a pagar'
' adiciona todas as comissoes pagas '
rsAjustaData.CommandText = "qryAtuNomes"
rsAjustaData.Execute()

'================================================================'
' Fecha ambas as conexões
conn.Close
Set conn = Nothing

connSales.Close
Set connSales = Nothing
Response.Write " Atulizado!"
%>
