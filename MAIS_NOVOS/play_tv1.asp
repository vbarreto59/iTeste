<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                  -->
<!-- Data: 04/12/2025                       -->
<!-- CODIGO_ARQUIVO: USSQQOOLZW             -->
<!-- OBS: Com seleção de mês e paletas      -->
<!-- ###################################### -->
<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="conSunSales.asp"-->

<%
' FUNÇÃO PARA POPULAR OS SELECTS DE FILTRO
Function GetUniqueValues(conn, fieldName, tableName)
    Dim dict, rs, sqlQuery
    Set dict = Server.CreateObject("Scripting.Dictionary")
    Set rs = Server.CreateObject("ADODB.Recordset")
    
    sqlQuery = "SELECT DISTINCT " & fieldName & " FROM " & tableName & " ORDER BY " & fieldName & ";"
    
    rs.Open sqlQuery, conn
    If Not rs.EOF Then
        Do While Not rs.EOF
            If Not IsNull(rs(fieldName)) Then
                dict.Add CStr(rs(fieldName)), 1
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    
    GetUniqueValues = dict.Keys
End Function

' FUNÇÃO PARA CONSTRUIR A CLÁUSULA WHERE
Function BuildWhereClause()
    Dim sqlWhere
    sqlWhere = " WHERE 1=1 AND Excluido = 0 AND Excluido IS NOT NULL"

    ' Manter outros filtros mas remover ano e mês do where global
    If Request.QueryString("diretoria") <> "" Then
        sqlWhere = sqlWhere & " AND Diretoria = '" & Replace(Request.QueryString("diretoria"), "'", "''") & "'"
    End If

    If Request.QueryString("gerencia") <> "" Then
        sqlWhere = sqlWhere & " AND Gerencia = '" & Replace(Request.QueryString("gerencia"), "'", "''") & "'"
    End If

    If Request.QueryString("corretor") <> "" Then
        sqlWhere = sqlWhere & " AND Corretor = '" & Replace(Request.QueryString("corretor"), "'", "''") & "'"
    End If

    If Request.QueryString("empreendimento") <> "" Then
        sqlWhere = sqlWhere & " AND NomeEmpreendimento = '" & Replace(Request.QueryString("empreendimento"), "'", "''") & "'"
    End If
    
    BuildWhereClause = sqlWhere
End Function

' FUNÇÃO PARA SERIALIZAR ARRAY EM JSON
Function JSON_Serialize(arr)
    Dim i, result, valor
    result = "["
    
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            ' Converter para string primeiro
            Dim strValor
            If IsNull(arr(i)) Then
                strValor = "0"
            Else
                strValor = CStr(arr(i))
            End If
            
            ' Se for numérico (pode conter vírgula como decimal)
            If IsNumeric(Replace(strValor, ",", ".")) Then
                ' Substituir vírgula por ponto para formato JSON
                valor = Replace(strValor, ",", ".")
                ' Remover qualquer formatação de milhar
                valor = Replace(valor, ".", "", 1, 1)
                result = result & valor
            ElseIf strValor = "" Then
                result = result & "0"
            Else
                ' Para strings, escapar aspas
                result = result & """" & Replace(strValor, """", "\""") & """"
            End If
            
            If i < UBound(arr) Then result = result & ","
        Next
    Else
        ' Se não for array, retornar array vazio
        result = "[]"
    End If
    
    result = result & "]"
    JSON_Serialize = result
End Function

' FUNÇÃO PARA FORMATAR VALOR PARA GRÁFICO
Function FormatForChart(valor)
    On Error Resume Next
    
    If IsNull(valor) Then
        FormatForChart = "0"
        Exit Function
    End If
    
    ' Tentar converter para número
    Dim numValor
    If IsNumeric(valor) Then
        tempValor = Replace(Valor, ".", "") ' Remover separador de milhar
        tempValor = Replace(tempValor, ",", ".") ' Converter vírgula decimal para ponto        
        numValor = tempValor
         'response.Write tempValor
        'Response.end         
    Else
        ' Remover formatação de moeda
        Dim tempValor
        tempValor = Trim(CStr(valor))
        tempValor = Replace(tempValor, "R$", "")
        tempValor = Replace(tempValor, "$", "")
        tempValor = Replace(tempValor, ".", "") ' Remover separador de milhar
        tempValor = Replace(tempValor, ",", ".") ' Converter vírgula decimal para ponto
        
        
        If IsNumeric(tempValor) Then
            numValor = CDbl(tempValor)
        Else
            numValor = 0
        End If
    End If
    
    ' Retornar como string sem formatação
    FormatForChart = numValor
    
    On Error GoTo 0
End Function

' =======================================================
' INÍCIO DO PROCESSAMENTO
' =======================================================

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open strConnSales

' Determinar ano e mês a serem exibidos
Dim anoExibicao, mesExibicao

If Request.QueryString("ano") <> "" Then
    anoExibicao = CInt(Request.QueryString("ano"))
Else
    anoExibicao = Year(Date())
End If

If Request.QueryString("mes") <> "" Then
    mesExibicao = CInt(Request.QueryString("mes"))
Else
    mesExibicao = Month(Date())
End If

' Verificar paleta selecionada
Dim paletaSelecionada
paletaSelecionada = Request.QueryString("paleta")
If paletaSelecionada = "" Then paletaSelecionada = "azul"

Dim whereClause
whereClause = BuildWhereClause()

' Usar as variáveis de exibição em todo o código
Dim anoAtual, mesAtual
anoAtual = anoExibicao
mesAtual = mesExibicao

' CALCULAR TICKET MÉDIO E QUANTIDADE DE UNIDADES
Dim ticketMedio, quantidadeUnidades, totalVendas, ticketMedioAno, quantidadeUnidadesAno, totalVendasAno
quantidadeUnidades = 0
totalVendas = 0

' Dados do mês atual
SQL = "SELECT COUNT(*) AS TotalUnidades, SUM(ValorUnidade) AS TotalVendas FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " AND MesVenda = " & mesAtual
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn

If Not rs.EOF Then
    If Not IsNull(rs("TotalUnidades")) Then
        quantidadeUnidades = rs("TotalUnidades")
    End If
    If Not IsNull(rs("TotalVendas")) Then
        totalVendas = rs("TotalVendas")
    End If
End If
rs.Close
Set rs = Nothing

If quantidadeUnidades > 0 And totalVendas > 0 Then
    ticketMedio = totalVendas / quantidadeUnidades
Else
    ticketMedio = 0
End If

' Dados do ano atual
SQL = "SELECT COUNT(*) AS TotalUnidades, SUM(ValorUnidade) AS TotalVendas FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn

If Not rs.EOF Then
    If Not IsNull(rs("TotalUnidades")) Then
        quantidadeUnidadesAno = rs("TotalUnidades")
    End If
    If Not IsNull(rs("TotalVendas")) Then
        totalVendasAno = rs("TotalVendas")
    End If
End If
rs.Close
Set rs = Nothing

If quantidadeUnidadesAno > 0 And totalVendasAno > 0 Then
    ticketMedioAno = totalVendasAno / quantidadeUnidadesAno
Else
    ticketMedioAno = 0
End If

' Calcular variação em relação ao mês anterior
Dim mesAnterior, totalVendasMesAnterior, variacaoMensal
mesAnterior = mesAtual - 1
If mesAnterior = 0 Then
    mesAnterior = 12
End If

SQL = "SELECT SUM(ValorUnidade) AS TotalVendas FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " AND MesVenda = " & mesAnterior
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn

If Not rs.EOF Then
    If Not IsNull(rs("TotalVendas")) Then
        totalVendasMesAnterior = rs("TotalVendas")
    Else
        totalVendasMesAnterior = 0
    End If
Else
    totalVendasMesAnterior = 0
End If
rs.Close
Set rs = Nothing

If totalVendasMesAnterior > 0 Then
    variacaoMensal = ((totalVendas - totalVendasMesAnterior) / totalVendasMesAnterior) * 100
Else
    variacaoMensal = 100
End If

' Calcular performance vs meta (exemplo com meta fictícia de 1.000.000)
Dim metaAnual, performanceAnual
metaAnual = 1000000 ' Valor de exemplo
performanceAnual = (totalVendasAno / metaAnual) * 100

Dim arrMesesNome(12)
arrMesesNome(1) = "Janeiro"
arrMesesNome(2) = "Fevereiro"
arrMesesNome(3) = "Março"
arrMesesNome(4) = "Abril"
arrMesesNome(5) = "Maio"
arrMesesNome(6) = "Junho"
arrMesesNome(7) = "Julho"
arrMesesNome(8) = "Agosto"
arrMesesNome(9) = "Setembro"
arrMesesNome(10) = "Outubro"
arrMesesNome(11) = "Novembro"
arrMesesNome(12) = "Dezembro"

Dim autoTime
autoTime = Request.QueryString("autotime")
If autoTime = "" Then autoTime = 10

' Preparar dados para os gráficos

' Gráfico de vendas anual por mês
Dim dadosVendasAnual(12), mesesAno(12), dadosVendasAnualChart(12)
For i = 1 to 12
    mesesAno(i-1) = Left(arrMesesNome(i), 3)
    SQL = "SELECT IIF(SUM(ValorUnidade) IS NULL, 0, SUM(ValorUnidade)) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " AND MesVenda = " & i


    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open SQL, conn
    
    If Not rs.EOF Then
        dadosVendasAnual(i-1) = rs("Total")
    Else
        dadosVendasAnual(i-1) = 0
    End If
    rs.Close
    Set rs = Nothing
    
    ' Formatar para gráfico
    dadosVendasAnualChart(i-1) = FormatForChart(dadosVendasAnual(i-1))
Next

' Verificar se há dados
Dim hasVendasData
hasVendasData = False
For i = 0 to 11
    If dadosVendasAnual(i) > 0 Then
        hasVendasData = True
        Exit For
    End If
Next

' Gráfico de diretorias
Dim diretoriasArray(), totaisDiretoriasArray()
ReDim diretoriasArray(-1)
ReDim totaisDiretoriasArray(-1)

SQL = "SELECT Diretoria, SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " GROUP BY Diretoria HAVING SUM(ValorUnidade) > 0 ORDER BY SUM(ValorUnidade) DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn

Dim diretoriasCount
diretoriasCount = 0

Do Until rs.EOF
    If Not IsNull(rs("Diretoria")) And Not IsNull(rs("Total")) Then
        If diretoriasCount = 0 Then
            ReDim diretoriasArray(0)
            ReDim totaisDiretoriasArray(0)
            diretoriasArray(0) = Trim(rs("Diretoria"))
            totaisDiretoriasArray(0) = FormatForChart(rs("Total"))
        Else
            ReDim Preserve diretoriasArray(diretoriasCount)
            ReDim Preserve totaisDiretoriasArray(diretoriasCount)
            diretoriasArray(diretoriasCount) = Trim(rs("Diretoria"))
            totaisDiretoriasArray(diretoriasCount) = FormatForChart(rs("Total"))
        End If
        diretoriasCount = diretoriasCount + 1
    End If
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

' Gráfico de gerências (top 10)
Dim gerenciasArray(), totaisGerenciasArray()
ReDim gerenciasArray(-1)
ReDim totaisGerenciasArray(-1)

SQL = "SELECT TOP 10 Gerencia, SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " AND Gerencia IS NOT NULL GROUP BY Gerencia HAVING SUM(ValorUnidade) > 0 ORDER BY SUM(ValorUnidade) DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn

Dim gerenciasCount
gerenciasCount = 0

Do Until rs.EOF
    If Not IsNull(rs("Gerencia")) And Not IsNull(rs("Total")) Then
        If gerenciasCount = 0 Then
            ReDim gerenciasArray(0)
            ReDim totaisGerenciasArray(0)
            gerenciasArray(0) = Trim(rs("Gerencia"))
            totaisGerenciasArray(0) = FormatForChart(rs("Total"))
        Else
            ReDim Preserve gerenciasArray(gerenciasCount)
            ReDim Preserve totaisGerenciasArray(gerenciasCount)
            gerenciasArray(gerenciasCount) = Trim(rs("Gerencia"))
            totaisGerenciasArray(gerenciasCount) = FormatForChart(rs("Total"))
        End If
        gerenciasCount = gerenciasCount + 1
    End If
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

' Gráfico de ticket médio mensal
Dim ticketMedioMensal(12), ticketMedioMensalChart(12)
For i = 1 to 12
    SQL = "SELECT COUNT(*) AS TotalUnidades, SUM(ValorUnidade) AS TotalVendas FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " AND MesVenda = " & i
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open SQL, conn
    
    If Not rs.EOF Then
        If Not IsNull(rs("TotalUnidades")) And Not IsNull(rs("TotalVendas")) Then
            If rs("TotalUnidades") > 0 And rs("TotalVendas") > 0 Then
                ticketMedioMensal(i-1) = rs("TotalVendas") / rs("TotalUnidades")
            Else
                ticketMedioMensal(i-1) = 0
            End If
        Else
            ticketMedioMensal(i-1) = 0
        End If
    Else
        ticketMedioMensal(i-1) = 0
    End If
    
    ' Preparar valor para gráfico
    ticketMedioMensalChart(i-1) = FormatForChart(ticketMedioMensal(i-1))
    
    rs.Close
    Set rs = Nothing
Next

' Gráfico de tipo de unidade
Dim tiposUnidadeArray(), vendasTiposUnidadeArray()
ReDim tiposUnidadeArray(4)
ReDim vendasTiposUnidadeArray(4)

tiposUnidadeArray(0) = "Apartamento"
tiposUnidadeArray(1) = "Casa"
tiposUnidadeArray(2) = "Sobrado"
tiposUnidadeArray(3) = "Terreno"
tiposUnidadeArray(4) = "Comercial"

' Valores de exemplo
vendasTiposUnidadeArray(0) = "450000"
vendasTiposUnidadeArray(1) = "320000"
vendasTiposUnidadeArray(2) = "280000"
vendasTiposUnidadeArray(3) = "150000"
vendasTiposUnidadeArray(4) = "80000"

' Preparar dados JSON para JavaScript
Dim diretoriasJSON, totaisDiretoriasJSON, gerenciasJSON, totaisGerenciasJSON
Dim mesesAnoJSON, vendasAnualJSON, ticketMedioJSON, tiposUnidadeJSON, vendasTiposUnidadeJSON

diretoriasJSON = JSON_Serialize(diretoriasArray)
totaisDiretoriasJSON = JSON_Serialize(totaisDiretoriasArray)
gerenciasJSON = JSON_Serialize(gerenciasArray)
totaisGerenciasJSON = JSON_Serialize(totaisGerenciasArray)
mesesAnoJSON = JSON_Serialize(mesesAno)

' Para vendas anuais, construir manualmente para garantir formato correto
vendasAnualJSON = "["
For i = 0 to 11
    vendasAnualJSON = vendasAnualJSON & dadosVendasAnualChart(i)
    If i < 11 Then vendasAnualJSON = vendasAnualJSON & ","
Next
vendasAnualJSON = vendasAnualJSON & "]"

ticketMedioJSON = JSON_Serialize(ticketMedioMensalChart)
tiposUnidadeJSON = JSON_Serialize(tiposUnidadeArray)
vendasTiposUnidadeJSON = JSON_Serialize(vendasTiposUnidadeArray)

' DEBUG: Verificar dados gerados
Response.Write "<!-- DEBUG INFO -->"
Response.Write "<!-- mesesAnoJSON: " & mesesAnoJSON & " -->"
Response.Write "<!-- vendasAnualJSON: " & vendasAnualJSON & " -->"
Response.Write "<!-- hasVendasData: " & hasVendasData & " -->"
Response.Write "<!-- anoAtual: " & anoAtual & " -->"

' NÃO FECHAR A CONEXÃO AQUI - vamos mantê-la aberta para as consultas dentro dos slides
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Dashboard de Vendas - Sala de Vendas</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            overflow-x: hidden;
            transition: background-color 0.5s ease;
        }
        
        /* Paleta Azul (Default) */
        .paleta-azul {
            background-color: #0a1929;
            color: #ffffff;
        }
        .paleta-azul .header {
            background: linear-gradient(135deg, #1a3a5f 0%, #0a1929 100%);
        }
        .paleta-azul .card {
            background-color: #1a3a5f;
        }
        .paleta-azul .card-header {
            background: linear-gradient(135deg, #2a4a7a 0%, #1a3a5f 100%);
            border-bottom: 1px solid #2a4a7a;
        }
        .paleta-azul .metric-value,
        .paleta-azul .metric-icon {
            color: #a0c4ff;
        }
        .paleta-azul .list-group-item {
            background-color: #1a3a5f;
            border: 1px solid #2a4a7a;
        }
        
        /* Paleta Verde */
        .paleta-verde {
            background-color: #0c2b1a;
            color: #ffffff;
        }
        .paleta-verde .header {
            background: linear-gradient(135deg, #1e4d2c 0%, #0c2b1a 100%);
        }
        .paleta-verde .card {
            background-color: #1e4d2c;
        }
        .paleta-verde .card-header {
            background: linear-gradient(135deg, #2e5d3c 0%, #1e4d2c 100%);
            border-bottom: 1px solid #2e5d3c;
        }
        .paleta-verde .metric-value,
        .paleta-verde .metric-icon {
            color: #86efac;
        }
        .paleta-verde .list-group-item {
            background-color: #1e4d2c;
            border: 1px solid #2e5d3c;
        }
        
        /* Paleta Roxa */
        .paleta-roxa {
            background-color: #1a0c2b;
            color: #ffffff;
        }
        .paleta-roxa .header {
            background: linear-gradient(135deg, #2a1c3b 0%, #1a0c2b 100%);
        }
        .paleta-roxa .card {
            background-color: #2a1c3b;
        }
        .paleta-roxa .card-header {
            background: linear-gradient(135deg, #3a2c4b 0%, #2a1c3b 100%);
            border-bottom: 1px solid #3a2c4b;
        }
        .paleta-roxa .metric-value,
        .paleta-roxa .metric-icon {
            color: #d8b4fe;
        }
        .paleta-roxa .list-group-item {
            background-color: #2a1c3b;
            border: 1px solid #3a2c4b;
        }
        
        /* NOVA: Paleta Laranja */
        .paleta-laranja {
            background-color: #2c1908;
            color: #ffffff;
        }
        .paleta-laranja .header {
            background: linear-gradient(135deg, #5a2c0d 0%, #2c1908 100%);
        }
        .paleta-laranja .card {
            background-color: #5a2c0d;
        }
        .paleta-laranja .card-header {
            background: linear-gradient(135deg, #7a4c1d 0%, #5a2c0d 100%);
            border-bottom: 1px solid #7a4c1d;
        }
        .paleta-laranja .metric-value,
        .paleta-laranja .metric-icon {
            color: #fbbf24;
        }
        .paleta-laranja .list-group-item {
            background-color: #5a2c0d;
            border: 1px solid #7a4c1d;
        }
        
        /* NOVA: Paleta Bordô */
        .paleta-bordo {
            background-color: #2c0819;
            color: #ffffff;
        }
        .paleta-bordo .header {
            background: linear-gradient(135deg, #5a0d2c 0%, #2c0819 100%);
        }
        .paleta-bordo .card {
            background-color: #5a0d2c;
        }
        .paleta-bordo .card-header {
            background: linear-gradient(135deg, #7a1d4c 0%, #5a0d2c 100%);
            border-bottom: 1px solid #7a1d4c;
        }
        .paleta-bordo .metric-value,
        .paleta-bordo .metric-icon {
            color: #f472b6;
        }
        .paleta-bordo .list-group-item {
            background-color: #5a0d2c;
            border: 1px solid #7a1d4c;
        }
        
        .dashboard-container {
            padding: 20px;
        }
        .header {
            text-align: center;
            margin-bottom: 20px;
            padding: 10px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
        }
        .header h1 {
            color: #ffffff;
            font-weight: 700;
            margin: 0;
            font-size: 2.5rem;
        }
        .header h2 {
            color: inherit;
            font-weight: 400;
            margin: 0;
            font-size: 1.5rem;
        }
        .paleta-azul .header h2 {
            color: #a0c4ff;
        }
        .paleta-verde .header h2 {
            color: #86efac;
        }
        .paleta-roxa .header h2 {
            color: #d8b4fe;
        }
        .paleta-laranja .header h2 {
            color: #fbbf24;
        }
        .paleta-bordo .header h2 {
            color: #f472b6;
        }
        .card {
            border: none;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            margin-bottom: 20px;
            transition: transform 0.3s ease;
        }
        .card:hover {
            transform: translateY(-5px);
        }
        .card-header {
            color: white;
            border-radius: 10px 10px 0 0 !important;
            font-weight: 600;
            padding: 15px 20px;
        }
        .metric-card {
            text-align: center;
            padding: 25px 15px;
            height: 100%;
        }
        .metric-value {
            font-size: 2.5rem;
            font-weight: bold;
            margin: 10px 0;
        }
        .metric-label {
            font-size: 1rem;
            margin-bottom: 0;
        }
        .paleta-azul .metric-label {
            color: #c0d6ff;
        }
        .paleta-verde .metric-label {
            color: #bbf7d0;
        }
        .paleta-roxa .metric-label {
            color: #e9d5ff;
        }
        .paleta-laranja .metric-label {
            color: #fde68a;
        }
        .paleta-bordo .metric-label {
            color: #fbcfe8;
        }
        .metric-icon {
            font-size: 2.5rem;
            margin-bottom: 15px;
        }
        .list-group-item {
            color: #ffffff;
            padding: 15px 20px;
        }
        .badge {
            font-weight: 600;
            padding: 8px 12px;
            border-radius: 10px;
        }
        .controls {
            position: fixed;
            bottom: 20px;
            right: 20px;
            z-index: 1000;
            border-radius: 10px;
            padding: 15px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
        }
        .paleta-azul .controls {
            background-color: rgba(26, 58, 95, 0.9);
        }
        .paleta-verde .controls {
            background-color: rgba(30, 77, 44, 0.9);
        }
        .paleta-roxa .controls {
            background-color: rgba(42, 28, 59, 0.9);
        }
        .paleta-laranja .controls {
            background-color: rgba(90, 44, 13, 0.9);
        }
        .paleta-bordo .controls {
            background-color: rgba(90, 13, 44, 0.9);
        }
        .slide {
            display: none;
        }
        .slide.active {
            display: block;
        }
        .comparison-chart {
            height: 300px;
        }
        .venda-item {
            border-left: 4px solid;
            padding-left: 15px;
            margin-bottom: 10px;
        }
        .paleta-azul .venda-item {
            border-left-color: #a0c4ff;
        }
        .paleta-verde .venda-item {
            border-left-color: #86efac;
        }
        .paleta-roxa .venda-item {
            border-left-color: #d8b4fe;
        }
        .paleta-laranja .venda-item {
            border-left-color: #fbbf24;
        }
        .paleta-bordo .venda-item {
            border-left-color: #f472b6;
        }
        .venda-info {
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .venda-details {
            font-size: 0.9rem;
        }
        .paleta-azul .venda-details {
            color: #c0d6ff;
        }
        .paleta-verde .venda-details {
            color: #bbf7d0;
        }
        .paleta-roxa .venda-details {
            color: #e9d5ff;
        }
        .paleta-laranja .venda-details {
            color: #fde68a;
        }
        .paleta-bordo .venda-details {
            color: #fbcfe8;
        }
        .chart-container {
            position: relative;
            height: 100%;
            min-height: 300px;
        }
        .progress {
            height: 10px;
            margin-bottom: 10px;
        }
        .paleta-azul .progress-bar {
            background-color: #a0c4ff;
        }
        .paleta-verde .progress-bar {
            background-color: #86efac;
        }
        .paleta-roxa .progress-bar {
            background-color: #d8b4fe;
        }
        .paleta-laranja .progress-bar {
            background-color: #fbbf24;
        }
        .paleta-bordo .progress-bar {
            background-color: #f472b6;
        }
        .countdown-timer {
            position: fixed;
            top: 20px;
            left: 20px;
            color: white;
            padding: 10px 15px;
            border-radius: 5px;
            font-size: 1rem;
            font-weight: bold;
            z-index: 999;
            display: none;
        }
        .paleta-azul .countdown-timer {
            background-color: rgba(26, 58, 95, 0.9);
        }
        .paleta-verde .countdown-timer {
            background-color: rgba(30, 77, 44, 0.9);
        }
        .paleta-roxa .countdown-timer {
            background-color: rgba(42, 28, 59, 0.9);
        }
        .paleta-laranja .countdown-timer {
            background-color: rgba(90, 44, 13, 0.9);
        }
        .paleta-bordo .countdown-timer {
            background-color: rgba(90, 13, 44, 0.9);
        }
        .info-box {
            border-radius: 10px;
            padding: 15px;
            margin-bottom: 15px;
            border-left: 4px solid;
        }
        .paleta-azul .info-box {
            background-color: #1a3a5f;
            border-left-color: #a0c4ff;
        }
        .paleta-verde .info-box {
            background-color: #1e4d2c;
            border-left-color: #86efac;
        }
        .paleta-roxa .info-box {
            background-color: #2a1c3b;
            border-left-color: #d8b4fe;
        }
        .paleta-laranja .info-box {
            background-color: #5a2c0d;
            border-left-color: #fbbf24;
        }
        .paleta-bordo .info-box {
            background-color: #5a0d2c;
            border-left-color: #f472b6;
        }
        .info-title {
            font-size: 1rem;
            margin-bottom: 5px;
        }
        .paleta-azul .info-title {
            color: #c0d6ff;
        }
        .paleta-verde .info-title {
            color: #bbf7d0;
        }
        .paleta-roxa .info-title {
            color: #e9d5ff;
        }
        .paleta-laranja .info-title {
            color: #fde68a;
        }
        .paleta-bordo .info-title {
            color: #fbcfe8;
        }
        .info-value {
            font-size: 1.5rem;
            font-weight: bold;
            color: #ffffff;
        }
        .trend-up {
            color: #4ade80;
        }
        .trend-down {
            color: #f87171;
        }
        .slide-indicator {
            position: fixed;
            bottom: 20px;
            left: 20px;
            border-radius: 10px;
            padding: 10px 15px;
            z-index: 999;
        }
        .paleta-azul .slide-indicator {
            background-color: rgba(26, 58, 95, 0.9);
        }
        .paleta-verde .slide-indicator {
            background-color: rgba(30, 77, 44, 0.9);
        }
        .paleta-roxa .slide-indicator {
            background-color: rgba(42, 28, 59, 0.9);
        }
        .paleta-laranja .slide-indicator {
            background-color: rgba(90, 44, 13, 0.9);
        }
        .paleta-bordo .slide-indicator {
            background-color: rgba(90, 13, 44, 0.9);
        }
        .slide-dot {
            display: inline-block;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            background-color: rgba(255, 255, 255, 0.3);
            margin: 0 5px;
            cursor: pointer;
        }
        .slide-dot.active {
            background-color: inherit;
        }
        .paleta-azul .slide-dot.active {
            background-color: #a0c4ff;
        }
        .paleta-verde .slide-dot.active {
            background-color: #86efac;
        }
        .paleta-roxa .slide-dot.active {
            background-color: #d8b4fe;
        }
        .paleta-laranja .slide-dot.active {
            background-color: #fbbf24;
        }
        .paleta-bordo .slide-dot.active {
            background-color: #f472b6;
        }
        .filter-active {
            animation: pulse 2s infinite;
            border: 2px solid;
        }
        .paleta-azul .filter-active {
            border-color: #a0c4ff;
        }
        .paleta-verde .filter-active {
            border-color: #86efac;
        }
        .paleta-roxa .filter-active {
            border-color: #d8b4fe;
        }
        .paleta-laranja .filter-active {
            border-color: #fbbf24;
        }
        .paleta-bordo .filter-active {
            border-color: #f472b6;
        }
        @keyframes pulse {
            0% { box-shadow: 0 0 0 0 rgba(160, 196, 255, 0.7); }
            70% { box-shadow: 0 0 0 10px rgba(160, 196, 255, 0); }
            100% { box-shadow: 0 0 0 0 rgba(160, 196, 255, 0); }
        }
        .paleta-selector {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 999;
            background-color: rgba(0,0,0,0.7);
            border-radius: 10px;
            padding: 10px;
        }
        .paleta-btn {
            width: 30px;
            height: 30px;
            border-radius: 50%;
            border: 2px solid white;
            cursor: pointer;
            margin: 0 5px;
            display: inline-block;
        }
        .paleta-btn.active {
            transform: scale(1.2);
            box-shadow: 0 0 10px rgba(255,255,255,0.8);
        }
        .paleta-btn-azul {
            background: linear-gradient(135deg, #1a3a5f 0%, #0a1929 100%);
        }
        .paleta-btn-verde {
            background: linear-gradient(135deg, #1e4d2c 0%, #0c2b1a 100%);
        }
        .paleta-btn-roxa {
            background: linear-gradient(135deg, #2a1c3b 0%, #1a0c2b 100%);
        }
        .paleta-btn-laranja {
            background: linear-gradient(135deg, #5a2c0d 0%, #2c1908 100%);
        }
        .paleta-btn-bordo {
            background: linear-gradient(135deg, #5a0d2c 0%, #2c0819 100%);
        }
        .filtro-badge {
            font-size: 0.8rem;
        }
        .form-select, .form-control {
            background-color: rgba(255,255,255,0.1);
            border: 1px solid rgba(255,255,255,0.3);
            color: white;
        }
        .form-select:focus, .form-control:focus {
            background-color: rgba(255,255,255,0.2);
            border-color: rgba(255,255,255,0.5);
            color: white;
            box-shadow: 0 0 0 0.25rem rgba(255,255,255,0.25);
        }
        .form-select option {
            background-color: #333;
            color: white;
        }
        .btn-group .btn {
            border-color: rgba(255,255,255,0.3);
        }
        .no-data-message {
            text-align: center;
            padding: 50px 20px;
            color: rgba(255,255,255,0.6);
        }
        .no-data-message i {
            font-size: 3rem;
            margin-bottom: 20px;
            opacity: 0.3;
        }
    </style>
</head>
<body class="paleta-<%=paletaSelecionada%>" id="bodyElement">

<div class="countdown-timer" id="countdown-timer">
    <i class="fas fa-clock"></i> Próxima visualização em: <span id="seconds-left">0</span>s
</div>

<div class="paleta-selector" id="paleta-selector">
    <span class="paleta-btn paleta-btn-azul <% If paletaSelecionada = "azul" Then Response.Write "active" %>" 
          data-paleta="azul" 
          title="Paleta Azul">
    </span>
    <span class="paleta-btn paleta-btn-verde <% If paletaSelecionada = "verde" Then Response.Write "active" %>" 
          data-paleta="verde" 
          title="Paleta Verde">
    </span>
    <span class="paleta-btn paleta-btn-roxa <% If paletaSelecionada = "roxa" Then Response.Write "active" %>" 
          data-paleta="roxa" 
          title="Paleta Roxa">
    </span>
    <span class="paleta-btn paleta-btn-laranja <% If paletaSelecionada = "laranja" Then Response.Write "active" %>" 
          data-paleta="laranja" 
          title="Paleta Laranja">
    </span>
    <span class="paleta-btn paleta-btn-bordo <% If paletaSelecionada = "bordo" Then Response.Write "active" %>" 
          data-paleta="bordo" 
          title="Paleta Bordô">
    </span>
</div>

<div class="slide-indicator" id="slide-indicator">
    <span class="slide-dot active" data-slide="1"></span>
    <span class="slide-dot" data-slide="2"></span>
    <span class="slide-dot" data-slide="3"></span>
    <span class="slide-dot" data-slide="4"></span>
    <span class="slide-dot" data-slide="5"></span>
</div>

<div class="dashboard-container">
    <div class="header" id="dashboardHeader">
        <h1>Dashboard de Vendas - Sala de Vendas</h1>
        <h2 id="mesAtualTexto">
            <%
            If Request.QueryString("mes") <> "" Then
                Response.Write arrMesesNome(mesExibicao) & " de " & anoExibicao
            ElseIf Request.QueryString("ano") <> "" Then
                Response.Write "Ano " & anoExibicao & " - Visão Anual"
            Else
                Response.Write arrMesesNome(mesAtual) & " de " & anoAtual
            End If
            %>
        </h2>
        
        <!-- Formulário de Filtro -->
        <div class="row justify-content-center mt-3">
            <div class="col-md-8">
                <form method="GET" action="" id="filterForm" class="row g-2">
                    <!-- Manter outros filtros existentes -->
                    <input type="hidden" name="diretoria" value="<%=Server.HTMLEncode(Request.QueryString("diretoria"))%>">
                    <input type="hidden" name="gerencia" value="<%=Server.HTMLEncode(Request.QueryString("gerencia"))%>">
                    <input type="hidden" name="corretor" value="<%=Server.HTMLEncode(Request.QueryString("corretor"))%>">
                    <input type="hidden" name="empreendimento" value="<%=Server.HTMLEncode(Request.QueryString("empreendimento"))%>">
                    <input type="hidden" name="paleta" id="paletaInput" value="<%=paletaSelecionada%>">
                    
                    <div class="col-md-3">
                        <select name="ano" class="form-select form-select-sm" id="anoSelect">
                            <%
                            ' Obter anos disponíveis
                            SQL = "SELECT DISTINCT AnoVenda FROM Vendas WHERE Excluido = 0 AND Excluido IS NOT NULL ORDER BY AnoVenda DESC"
                            Set rsAnos = Server.CreateObject("ADODB.Recordset")
                            rsAnos.Open SQL, conn
                            
                            anoSelecionado = Request.QueryString("ano")
                            If anoSelecionado = "" Then anoSelecionado = anoAtual
                            
                            Do Until rsAnos.EOF
                                anoOpt = rsAnos("AnoVenda")
                                %>
                                <option value="<%=anoOpt%>" <% If CStr(anoOpt) = CStr(anoSelecionado) Then Response.Write "selected" %>><%=anoOpt%></option>
                                <%
                                rsAnos.MoveNext
                            Loop
                            rsAnos.Close
                            Set rsAnos = Nothing
                            %>
                        </select>
                    </div>
                    <div class="col-md-4">
                        <select name="mes" class="form-select form-select-sm" id="mesSelect">
                            <option value="">Todos os meses</option>
                            <%
                            mesSelecionado = Request.QueryString("mes")
                            For i = 1 to 12
                                %>
                                <option value="<%=i%>" <% If CStr(i) = CStr(mesSelecionado) Then Response.Write "selected" %>><%=arrMesesNome(i)%></option>
                                <%
                            Next
                            %>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <select name="autotime" class="form-select form-select-sm" id="autoTimeSelectForm">
                            <option value="5" <% If CStr(autoTime) = "5" Then Response.Write "selected" %>>5 segundos</option>
                            <option value="10" <% If CStr(autoTime) = "10" Then Response.Write "selected" %>>10 segundos</option>
                            <option value="15" <% If CStr(autoTime) = "15" Then Response.Write "selected" %>>15 segundos</option>
                            <option value="20" <% If CStr(autoTime) = "20" Then Response.Write "selected" %>>20 segundos</option>
                            <option value="25" <% If CStr(autoTime) = "25" Then Response.Write "selected" %>>25 segundos</option>
                            <option value="30" <% If CStr(autoTime) = "30" Then Response.Write "selected" %>>30 segundos</option>
                        </select>
                    </div>
                    <div class="col-md-2">
                        <div class="d-grid gap-1">
                            <button type="submit" class="btn btn-sm btn-primary">
                                <i class="fas fa-filter"></i> Aplicar
                            </button>
                            <% 
                            If Request.QueryString("mes") <> "" OR Request.QueryString("ano") <> "" OR Request.QueryString("autotime") <> "" Then 
                            %>
                            <a href="?paleta=<%=paletaSelecionada%>" class="btn btn-sm btn-secondary">
                                <i class="fas fa-times"></i> Limpar
                            </a>
                            <% End If %>
                        </div>
                    </div>
                </form>
                
                <% 
                If Request.QueryString("ano") <> "" OR Request.QueryString("mes") <> "" Then 
                %>
                <div class="mt-2 text-center">
                    <span class="badge filtro-badge bg-info">
                        <i class="fas fa-filter"></i> 
                        Filtro Ativo: 
                        <%
                        If Request.QueryString("ano") <> "" Then
                            Response.Write "Ano " & Request.QueryString("ano")
                        End If
                        If Request.QueryString("mes") <> "" Then
                            If Request.QueryString("ano") <> "" Then Response.Write " - "
                            Response.Write arrMesesNome(CInt(Request.QueryString("mes")))
                        End If
                        %>
                        <% If Request.QueryString("autotime") <> "" Then %>
                         | Intervalo: <%=Request.QueryString("autotime")%>s
                        <% End If %>
                    </span>
                </div>
                <% End If %>
            </div>
        </div>
    </div>

    <!-- Slide 1: Visão Geral do Mês -->
    <div class="slide active" id="slide1">
        <div class="row">
            <div class="col-md-3">
                <div class="card metric-card">
                    <i class="fas fa-money-bill-wave metric-icon"></i>
                    <div class="metric-value">R$ <%=FormatNumber(totalVendas, 2)%></div>
                    <p class="metric-label">Vendas do Mês</p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card metric-card">
                    <i class="fas fa-cube metric-icon"></i>
                    <div class="metric-value"><%=FormatNumber(quantidadeUnidades, 0)%></div>
                    <p class="metric-label">Unidades Vendidas</p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card metric-card">
                    <i class="fas fa-ticket-alt metric-icon"></i>
                    <div class="metric-value">R$ <%=FormatNumber(ticketMedio, 2)%></div>
                    <p class="metric-label">Ticket Médio Mensal</p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card metric-card">
                    <i class="fas fa-chart-line metric-icon"></i>
                    <div class="metric-value">
                        <% If variacaoMensal >= 0 Then %>
                            <span class="trend-up">+<%=FormatNumber(variacaoMensal, 1)%>%</span>
                        <% Else %>
                            <span class="trend-down"><%=FormatNumber(variacaoMensal, 1)%>%</span>
                        <% End If %>
                    </div>
                    <p class="metric-label">Variação vs Mês Anterior</p>
                </div>
            </div>
        </div>
        
        <div class="row mt-4">
            <div class="col-md-8">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-chart-bar"></i> Vendas do Ano por Mês - <%=anoAtual%></h5>
                    </div>
                    <div class="card-body">
                        <% If hasVendasData Then %>
                        <div class="chart-container">
                            <canvas id="graficoVendasAnual"></canvas>
                        </div>
                        <% Else %>
                        <div class="no-data-message">
                            <i class="fas fa-chart-bar fa-3x"></i>
                            <h4>Não há dados de vendas para <%=anoAtual%></h4>
                            <p>Selecione outro ano ou verifique os filtros aplicados.</p>
                        </div>
                        <% End If %>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-trophy"></i> Top 5 Corretores - <%=arrMesesNome(mesAtual)%></h5>
                    </div>
                    <div class="card-body">
                        <%
                        SQL = "SELECT TOP 5 Corretor, SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " AND MesVenda = " & mesAtual & " GROUP BY Corretor ORDER BY SUM(ValorUnidade) DESC"
                        Set rsSlide1 = Server.CreateObject("ADODB.Recordset")
                        rsSlide1.Open SQL, conn
                        
                        If Not rsSlide1.EOF Then
                            Do Until rsSlide1.EOF
                                Response.Write "<div class='info-box'>"
                                Response.Write "<div class='info-title'>" & rsSlide1("Corretor") & "</div>"
                                Response.Write "<div class='info-value'>R$ " & FormatNumber(rsSlide1("Total"), 2) & "</div>"
                                Response.Write "</div>"
                                rsSlide1.MoveNext
                            Loop
                        Else
                            Response.Write "<div class='no-data-message'>"
                            Response.Write "<i class='fas fa-chart-bar'></i>"
                            Response.Write "<p>Nenhum dado disponível</p>"
                            Response.Write "</div>"
                        End If
                        rsSlide1.Close
                        Set rsSlide1 = Nothing
                        %>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Slide 2: Comparativo de Diretorias -->
    <div class="slide" id="slide2">
        <div class="row">
            <div class="col-md-8">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-chart-pie"></i> Comparativo de Diretorias - <%=anoAtual%></h5>
                    </div>
                    <div class="card-body">
                        <%
                        If IsArray(diretoriasArray) And UBound(diretoriasArray) >= 0 And diretoriasArray(0) <> "" Then
                        %>
                        <div class="chart-container">
                            <canvas id="graficoDiretorias"></canvas>
                        </div>
                        <%
                        Else
                        %>
                        <div class="no-data-message">
                            <i class="fas fa-chart-pie"></i>
                            <p>Não há dados de diretorias para exibir no período selecionado</p>
                        </div>
                        <%
                        End If
                        %>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-list-ol"></i> Ranking de Diretorias</h5>
                    </div>
                    <div class="card-body">
                        <%
                        If IsArray(diretoriasArray) And UBound(diretoriasArray) >= 0 And diretoriasArray(0) <> "" Then
                            SQL = "SELECT Diretoria, SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " GROUP BY Diretoria HAVING SUM(ValorUnidade) > 0 ORDER BY SUM(ValorUnidade) DESC"
                            Set rsSlide2 = Server.CreateObject("ADODB.Recordset")
                            rsSlide2.Open SQL, conn
                            
                            contador = 1
                            If Not rsSlide2.EOF Then
                                Do Until rsSlide2.EOF
                                    Response.Write "<div class='info-box'>"
                                    Response.Write "<div class='info-title'>#" & contador & " " & rsSlide2("Diretoria") & "</div>"
                                    Response.Write "<div class='info-value'>R$ " & FormatNumber(rsSlide2("Total"), 2) & "</div>"
                                    Response.Write "</div>"
                                    contador = contador + 1
                                    rsSlide2.MoveNext
                                Loop
                            End If
                            rsSlide2.Close
                            Set rsSlide2 = Nothing
                        Else
                            Response.Write "<div class='no-data-message'>"
                            Response.Write "<i class='fas fa-list-ol'></i>"
                            Response.Write "<p>Nenhum dado disponível</p>"
                            Response.Write "</div>"
                        End If
                        %>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Slide 3: Comparativo de Gerências -->
    <div class="slide" id="slide3">
        <div class="row">
            <div class="col-md-8">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-chart-bar"></i> Comparativo de Gerências - <%=anoAtual%></h5>
                    </div>
                    <div class="card-body">
                        <%
                        If IsArray(gerenciasArray) And UBound(gerenciasArray) >= 0 And gerenciasArray(0) <> "" Then
                        %>
                        <div class="chart-container">
                            <canvas id="graficoGerencias"></canvas>
                        </div>
                        <%
                        Else
                        %>
                        <div class="no-data-message">
                            <i class="fas fa-chart-bar"></i>
                            <p>Não há dados de gerências para exibir no período selecionado</p>
                        </div>
                        <%
                        End If
                        %>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-list-ol"></i> Top 10 Gerências</h5>
                    </div>
                    <div class="card-body">
                        <%
                        If IsArray(gerenciasArray) And UBound(gerenciasArray) >= 0 And gerenciasArray(0) <> "" Then
                            SQL = "SELECT TOP 10 Gerencia, SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " GROUP BY Gerencia ORDER BY SUM(ValorUnidade) DESC"
                            Set rsSlide3 = Server.CreateObject("ADODB.Recordset")
                            rsSlide3.Open SQL, conn
                            
                            contador = 1
                            If Not rsSlide3.EOF Then
                                Do Until rsSlide3.EOF
                                    Response.Write "<div class='info-box'>"
                                    Response.Write "<div class='info-title'>#" & contador & " " & rsSlide3("Gerencia") & "</div>"
                                    Response.Write "<div class='info-value'>R$ " & FormatNumber(rsSlide3("Total"), 2) & "</div>"
                                    Response.Write "</div>"
                                    contador = contador + 1
                                    rsSlide3.MoveNext
                                Loop
                            End If
                            rsSlide3.Close
                            Set rsSlide3 = Nothing
                        Else
                            Response.Write "<div class='no-data-message'>"
                            Response.Write "<i class='fas fa-list-ol'></i>"
                            Response.Write "<p>Nenhum dado disponível</p>"
                            Response.Write "</div>"
                        End If
                        %>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Slide 4: Ticket Médio e Eficiência -->
    <div class="slide" id="slide4">
        <div class="row">
            <div class="col-md-4">
                <div class="card metric-card">
                    <i class="fas fa-ticket-alt metric-icon"></i>
                    <div class="metric-value">R$ <%=FormatNumber(ticketMedioAno, 2)%></div>
                    <p class="metric-label">Ticket Médio Anual</p>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card metric-card">
                    <i class="fas fa-money-bill-wave metric-icon"></i>
                    <div class="metric-value">R$ <%=FormatNumber(totalVendasAno, 2)%></div>
                    <p class="metric-label">Vendas do Ano</p>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card metric-card">
                    <i class="fas fa-cube metric-icon"></i>
                    <div class="metric-value"><%=FormatNumber(quantidadeUnidadesAno, 0)%></div>
                    <p class="metric-label">Unidades Vendidas no Ano</p>
                </div>
            </div>
        </div>
        
        <div class="row mt-4">
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-chart-line"></i> Evolução do Ticket Médio Mensal - <%=anoAtual%></h5>
                    </div>
                    <div class="card-body">
                        <div class="chart-container">
                            <canvas id="graficoTicketMedio"></canvas>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-building"></i> Top 5 Empreendimentos - <%=anoAtual%></h5>
                    </div>
                    <div class="card-body">
                        <%
                        SQL = "SELECT TOP 5 NomeEmpreendimento, SUM(ValorUnidade) AS Total FROM Vendas " & whereClause & " AND AnoVenda = " & anoAtual & " GROUP BY NomeEmpreendimento ORDER BY SUM(ValorUnidade) DESC"
                        Set rsSlide4 = Server.CreateObject("ADODB.Recordset")
                        rsSlide4.Open SQL, conn
                        
                        If Not rsSlide4.EOF Then
                            Do Until rsSlide4.EOF
                                Response.Write "<div class='info-box'>"
                                Response.Write "<div class='info-title'>" & rsSlide4("NomeEmpreendimento") & "</div>"
                                Response.Write "<div class='info-value'>R$ " & FormatNumber(rsSlide4("Total"), 2) & "</div>"
                                Response.Write "</div>"
                                rsSlide4.MoveNext
                            Loop
                        Else
                            Response.Write "<div class='no-data-message'>"
                            Response.Write "<i class='fas fa-building'></i>"
                            Response.Write "<p>Nenhum dado disponível</p>"
                            Response.Write "</div>"
                        End If
                        rsSlide4.Close
                        Set rsSlide4 = Nothing
                        %>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Slide 5: Últimas Vendas e Performance -->
    <div class="slide" id="slide5">
        <div class="row">
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-clock"></i> Últimas 10 Vendas</h5>
                    </div>
                    <div class="card-body">
                        <%
                        SQL = "SELECT TOP 10 Corretor, ValorUnidade, NomeEmpreendimento, Gerencia, DiaVenda, MesVenda, AnoVenda FROM Vendas " & whereClause & " ORDER BY AnoVenda DESC, MesVenda DESC, DiaVenda DESC, ID DESC"
                        Set rsSlide5 = Server.CreateObject("ADODB.Recordset")
                        rsSlide5.Open SQL, conn
                        
                        If Not rsSlide5.EOF Then
                            Do While Not rsSlide5.EOF
                                %>
                                <div class="venda-item">
                                    <div class="venda-info">
                                        <strong><%=rsSlide5("Corretor")%></strong>
                                        <span class="badge bg-primary">R$ <%=FormatNumber(rsSlide5("ValorUnidade"), 2)%></span>
                                    </div>
                                    <div class="venda-details">
                                        <small>
                                            <i class="fas fa-building"></i> <%=rsSlide5("NomeEmpreendimento")%> | 
                                            <i class="fas fa-user-tie"></i> <%=rsSlide5("Gerencia")%> | 
                                            <i class="fas fa-calendar"></i> <%=rsSlide5("DiaVenda")%>/<%=rsSlide5("MesVenda")%>/<%=rsSlide5("AnoVenda")%>
                                        </small>
                                    </div>
                                </div>
                                <%
                                rsSlide5.MoveNext
                            Loop
                        Else
                            Response.Write "<div class='no-data-message'>"
                            Response.Write "<i class='fas fa-clock'></i>"
                            Response.Write "<p>Nenhuma venda encontrada</p>"
                            Response.Write "</div>"
                        End If
                        
                        rsSlide5.Close
                        Set rsSlide5 = Nothing
                        %>
                    </div>
                </div>
            </div>
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-tachometer-alt"></i> Metas e Performance - <%=anoAtual%></h5>
                    </div>
                    <div class="card-body">
                        <div class="info-box">
                            <div class="info-title">Meta Anual</div>
                            <div class="info-value">R$ <%=FormatNumber(metaAnual, 2)%></div>
                        </div>
                        <div class="info-box">
                            <div class="info-title">Vendas Realizadas</div>
                            <div class="info-value">R$ <%=FormatNumber(totalVendasAno, 2)%></div>
                        </div>
                        <div class="info-box">
                            <div class="info-title">Performance</div>
                            <div class="info-value"><%=FormatNumber(performanceAnual, 1)%>%</div>
                        </div>
                        <div class="progress mt-3">
                            <div class="progress-bar" role="progressbar" style="width: <%=performanceAnual%>%;" aria-valuenow="<%=performanceAnual%>" aria-valuemin="0" aria-valuemax="100"></div>
                        </div>
                    </div>
                </div>
                
                <div class="card mt-4">
                    <div class="card-header">
                        <h5 class="mb-0"><i class="fas fa-chart-area"></i> Vendas por Tipo de Unidade</h5>
                    </div>
                    <div class="card-body">
                        <div class="chart-container">
                            <canvas id="graficoTipoUnidade"></canvas>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<div class="controls">
    <div class="btn-group" role="group">
        <button type="button" class="btn btn-primary" id="prevSlide">
            <i class="fas fa-chevron-left"></i>
        </button>
        <button type="button" class="btn btn-success" id="playPause">
            <i class="fas fa-play" id="playIcon"></i>
        </button>
        <button type="button" class="btn btn-primary" id="nextSlide">
            <i class="fas fa-chevron-right"></i>
        </button>
    </div>
    <div class="mt-2">
        <select class="form-select form-select-sm" id="autoTimeSelect">
            <option value="5" <% If CStr(autoTime) = "5" Then Response.Write "selected" %>>5s</option>
            <option value="10" <% If CStr(autoTime) = "10" Then Response.Write "selected" %>>10s</option>
            <option value="15" <% If CStr(autoTime) = "15" Then Response.Write "selected" %>>15s</option>
            <option value="20" <% If CStr(autoTime) = "20" Then Response.Write "selected" %>>20s</option>
            <option value="25" <% If CStr(autoTime) = "25" Then Response.Write "selected" %>>25s</option>
            <option value="30" <% If CStr(autoTime) = "30" Then Response.Write "selected" %>>30s</option>
        </select>
    </div>
</div>

<script>
    // Configuração do slideshow
    let currentSlide = 1;
    const totalSlides = 5;
    let autoPlay = true;
    let slideInterval;
    let countdownInterval;
    const countdownTimer = document.getElementById('countdown-timer');
    const secondsLeftSpan = document.getElementById('seconds-left');
    const playPauseBtn = document.getElementById('playPause');
    const playIcon = document.getElementById('playIcon');
    const autoTimeSelect = document.getElementById('autoTimeSelect');
    let slideDuration = parseInt('<%=autoTime%>') || 10;

    // Paleta de cores vibrantes para gráficos
    const vibrantColors = [
        '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF', 
        '#FF9F40', '#FF6384', '#C9CBCF', '#36A2EB', '#FFCE56',
        '#4BC0C0', '#9966FF', '#FF9F40', '#FF6384', '#36A2EB'
    ];

    // Função para obter cores baseadas no tema
    function getChartColors() {
        const bodyClass = document.body.className;
        
        if (bodyClass.includes('azul')) {
            return {
                primary: '#a0c4ff',
                background: '#a0c4ff',
                border: '#a0c4ff',
                vibrant: vibrantColors
            };
        } else if (bodyClass.includes('verde')) {
            return {
                primary: '#86efac',
                background: '#86efac',
                border: '#86efac',
                vibrant: [
                    '#86efac', '#4ade80', '#22c55e',
                    '#16a34a', '#15803d', '#166534',
                    '#14532d', '#052e16', '#86efac'
                ]
            };
        } else if (bodyClass.includes('roxa')) {
            return {
                primary: '#d8b4fe',
                background: '#d8b4fe',
                border: '#d8b4fe',
                vibrant: vibrantColors
            };
        } else if (bodyClass.includes('laranja')) {
            return {
                primary: '#fbbf24',
                background: '#fbbf24',
                border: '#fbbf24',
                vibrant: [
                    '#fbbf24', '#f59e0b', '#d97706',
                    '#b45309', '#92400e', '#78350f',
                    '#fbbf24', '#f59e0b', '#d97706'
                ]
            };
        } else if (bodyClass.includes('bordo')) {
            return {
                primary: '#f472b6',
                background: '#f472b6',
                border: '#f472b6',
                vibrant: [
                    '#f472b6', '#ec4899', '#db2777',
                    '#be185d', '#9d174d', '#831843',
                    '#f472b6', '#ec4899', '#db2777'
                ]
            };
        } else {
            return {
                primary: '#a0c4ff',
                background: '#a0c4ff',
                border: '#a0c4ff',
                vibrant: vibrantColors
            };
        }
    }

    // Função para mostrar slide específico
    function showSlide(slideNumber) {
        // Validar número do slide
        if (slideNumber < 1) slideNumber = totalSlides;
        if (slideNumber > totalSlides) slideNumber = 1;
        
        // Esconder todos os slides
        document.querySelectorAll('.slide').forEach(slide => {
            slide.classList.remove('active');
        });
        
        // Mostrar o slide atual
        const slideElement = document.getElementById('slide' + slideNumber);
        if (slideElement) {
            slideElement.classList.add('active');
        }
        
        currentSlide = slideNumber;
        
        // Atualizar indicadores
        document.querySelectorAll('.slide-dot').forEach((dot, index) => {
            if (index + 1 === slideNumber) {
                dot.classList.add('active');
            } else {
                dot.classList.remove('active');
            }
        });
        
        // Reiniciar o contador
        resetCountdown();
    }

    // Função para próximo slide
    function nextSlide() {
        let next = currentSlide + 1;
        if (next > totalSlides) next = 1;
        showSlide(next);
    }

    // Função para slide anterior
    function prevSlide() {
        let prev = currentSlide - 1;
        if (prev < 1) prev = totalSlides;
        showSlide(prev);
    }

    // Função para iniciar slideshow automático
    function startAutoPlay() {
        if (autoPlay) {
            clearInterval(slideInterval);
            slideInterval = setInterval(nextSlide, slideDuration * 1000);
            countdownTimer.style.display = 'block';
            playIcon.classList.remove('fa-play');
            playIcon.classList.add('fa-pause');
            resetCountdown();
        }
    }

    // Função para parar slideshow
    function stopAutoPlay() {
        clearInterval(slideInterval);
        clearInterval(countdownInterval);
        countdownTimer.style.display = 'none';
        playIcon.classList.remove('fa-pause');
        playIcon.classList.add('fa-play');
    }

    // Função para alternar play/pause
    function toggleAutoPlay() {
        autoPlay = !autoPlay;
        if (autoPlay) {
            startAutoPlay();
        } else {
            stopAutoPlay();
        }
    }

    // Função para reiniciar contador
    function resetCountdown() {
        if (autoPlay) {
            clearInterval(countdownInterval);
            let secondsLeft = slideDuration;
            secondsLeftSpan.textContent = secondsLeft;
            
            countdownInterval = setInterval(() => {
                secondsLeft--;
                secondsLeftSpan.textContent = secondsLeft;
                
                if (secondsLeft <= 0) {
                    clearInterval(countdownInterval);
                }
            }, 1000);
        }
    }

    // Função para atualizar duração do slide
    function updateSlideDuration() {
        slideDuration = parseInt(autoTimeSelect.value);
        if (autoPlay) {
            startAutoPlay();
        }
        document.getElementById('autoTimeSelectForm').value = slideDuration;
    }

    // Função para mudar paleta de cores
    function changePaleta(paleta) {
        document.body.className = 'paleta-' + paleta;
        document.getElementById('paletaInput').value = paleta;
        
        document.querySelectorAll('.paleta-btn').forEach(btn => {
            btn.classList.remove('active');
            if (btn.dataset.paleta === paleta) {
                btn.classList.add('active');
            }
        });
        
        updateChartColors();
    }

    // Função para atualizar cores dos gráficos
    function updateChartColors() {
        const colors = getChartColors();
        
        if (window.vendasAnualChart) {
            window.vendasAnualChart.data.datasets[0].backgroundColor = colors.background + '80';
            window.vendasAnualChart.data.datasets[0].borderColor = colors.border;
            window.vendasAnualChart.update('none');
        }
        
        if (window.diretoriasChart) {
            const diretoriasCount = window.diretoriasChart.data.labels.length;
            const newColors = colors.vibrant.slice(0, diretoriasCount);
            window.diretoriasChart.data.datasets[0].backgroundColor = newColors;
            window.diretoriasChart.update('none');
        }
        
        if (window.gerenciasChart) {
            window.gerenciasChart.data.datasets[0].backgroundColor = colors.background + '80';
            window.gerenciasChart.data.datasets[0].borderColor = colors.border;
            window.gerenciasChart.update('none');
        }
        
        if (window.ticketMedioChart) {
            window.ticketMedioChart.data.datasets[0].backgroundColor = colors.background + '40';
            window.ticketMedioChart.data.datasets[0].borderColor = colors.border;
            window.ticketMedioChart.update('none');
        }
        
        if (window.tipoUnidadeChart) {
            const tipoCount = window.tipoUnidadeChart.data.labels.length;
            const newColors = colors.vibrant.slice(0, tipoCount);
            window.tipoUnidadeChart.data.datasets[0].backgroundColor = newColors;
            window.tipoUnidadeChart.update('none');
        }
    }

    // Inicializar controles do slideshow
    function initSlideshowControls() {
        document.getElementById('nextSlide').addEventListener('click', nextSlide);
        document.getElementById('prevSlide').addEventListener('click', prevSlide);
        document.getElementById('playPause').addEventListener('click', toggleAutoPlay);
        
        autoTimeSelect.addEventListener('change', updateSlideDuration);
        
        document.querySelectorAll('.slide-dot').forEach(dot => {
            dot.addEventListener('click', function() {
                const slideNumber = parseInt(this.getAttribute('data-slide'));
                showSlide(slideNumber);
            });
        });
        
        document.querySelectorAll('.paleta-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                const paleta = this.dataset.paleta;
                changePaleta(paleta);
                document.getElementById('filterForm').submit();
            });
        });
        
        document.getElementById('anoSelect').addEventListener('change', function() {
            this.form.submit();
        });

        document.getElementById('mesSelect').addEventListener('change', function() {
            this.form.submit();
        });
        
        const formSelect = document.getElementById('autoTimeSelectForm');
        if (formSelect) {
            formSelect.addEventListener('change', function() {
                autoTimeSelect.value = this.value;
                updateSlideDuration();
            });
        }
        
        const urlParams = new URLSearchParams(window.location.search);
        if (urlParams.has('ano') || urlParams.has('mes')) {
            const header = document.getElementById('dashboardHeader');
            if (header) {
                header.classList.add('filter-active');
                
                setTimeout(() => {
                    header.classList.remove('filter-active');
                }, 3000);
            }
        }
        
        setTimeout(() => {
            startAutoPlay();
        }, 1000);
    }
</script>

<script>
    // ################# Inicializar gráficos
    function initCharts() {
        const diretoriasData = <%=diretoriasJSON%> || [];
        const totaisDiretoriasData = <%=totaisDiretoriasJSON%> || [];
        const gerenciasData = <%=gerenciasJSON%> || [];
        const totaisGerenciasData = <%=totaisGerenciasJSON%> || [];
        const mesesAnoData = <%=mesesAnoJSON%> || [];
        const vendasAnualData = <%=vendasAnualJSON%> || [];
        const ticketMedioData = <%=ticketMedioJSON%> || [];
        const tiposUnidadeData = <%=tiposUnidadeJSON%> || [];
        const vendasTiposUnidadeData = <%=vendasTiposUnidadeJSON%> || [];

        // Gráfico de vendas anual
        if (document.getElementById('graficoVendasAnual')) {
            const ctxVendasAnual = document.getElementById('graficoVendasAnual').getContext('2d');
            const colors = getChartColors();
            
            const vendasData = Array.isArray(vendasAnualData) ? 
                vendasAnualData.map(v => {
                    const num = parseFloat(v);
                    return isNaN(num) ? 0 : num;
                }) : 
                new Array(12).fill(0);
            
            window.vendasAnualChart = new Chart(ctxVendasAnual, {
                type: 'bar',
                data: {
                    labels: mesesAnoData,
                    datasets: [{
                        label: 'Vendas Mensais',
                        data: vendasData,
                        backgroundColor: colors.background + '80',
                        borderColor: colors.border,
                        borderWidth: 2,
                        borderRadius: 4
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            labels: {
                                color: '#ffffff',
                                font: { size: 14 }
                            }
                        },
                        tooltip: {
                            backgroundColor: 'rgba(0,0,0,0.8)',
                            titleColor: '#ffffff',
                            bodyColor: '#ffffff',
                            callbacks: {
                                label: function(context) {
                                    return 'R$ ' + context.parsed.y.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
                                }
                            }
                        }
                    },
                    scales: {
                        x: {
                            grid: { color: 'rgba(255,255,255,0.1)' },
                            ticks: {
                                color: '#ffffff',
                                font: { size: 12 }
                            }
                        },
                        y: {
                            beginAtZero: true,
                            grid: { color: 'rgba(255,255,255,0.1)' },
                            ticks: {
                                color: '#ffffff',
                                font: { size: 12 },
                                callback: function(value) {
                                    return 'R$ ' + value.toLocaleString('pt-BR', {minimumFractionDigits: 0, maximumFractionDigits: 0});
                                }
                            }
                        }
                    }
                }
            });
        }

        // Gráfico de diretorias
        if (document.getElementById('graficoDiretorias') && Array.isArray(diretoriasData) && diretoriasData.length > 0) {
            const ctxDiretorias = document.getElementById('graficoDiretorias').getContext('2d');
            const colors = getChartColors();
            
            const pieColors = colors.vibrant.slice(0, diretoriasData.length);
            const diretoriasValues = totaisDiretoriasData.map(v => parseFloat(v) || 0);
            
            window.diretoriasChart = new Chart(ctxDiretorias, {
                type: 'pie',
                data: {
                    labels: diretoriasData,
                    datasets: [{
                        data: diretoriasValues,
                        backgroundColor: pieColors,
                        borderColor: '#ffffff',
                        borderWidth: 2
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'right',
                            labels: {
                                color: '#ffffff',
                                font: { size: 12 },
                                padding: 15
                            }
                        },
                        tooltip: {
                            backgroundColor: 'rgba(0,0,0,0.8)',
                            titleColor: '#ffffff',
                            bodyColor: '#ffffff',
                            callbacks: {
                                label: function(context) {
                                    const label = context.label || '';
                                    const value = context.parsed;
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const percentage = total > 0 ? Math.round((value / total) * 100) : 0;
                                    return `${label}: R$ ${value.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})} (${percentage}%)`;
                                }
                            }
                        }
                    }
                }
            });
        }

        // Gráfico de gerências
        if (document.getElementById('graficoGerencias') && Array.isArray(gerenciasData) && gerenciasData.length > 0) {
            const ctxGerencias = document.getElementById('graficoGerencias').getContext('2d');
            const colors = getChartColors();
            
            const gerenciasValues = totaisGerenciasData.map(v => parseFloat(v) || 0);
            
            window.gerenciasChart = new Chart(ctxGerencias, {
                type: 'bar',
                data: {
                    labels: gerenciasData,
                    datasets: [{
                        label: 'Vendas por Gerência',
                        data: gerenciasValues,
                        backgroundColor: colors.background + '80',
                        borderColor: colors.border,
                        borderWidth: 2,
                        borderRadius: 4
                    }]
                },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            backgroundColor: 'rgba(0,0,0,0.8)',
                            titleColor: '#ffffff',
                            bodyColor: '#ffffff',
                            callbacks: {
                                label: function(context) {
                                    return 'R$ ' + context.parsed.x.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
                                }
                            }
                        }
                    },
                    scales: {
                        x: {
                            beginAtZero: true,
                            grid: { color: 'rgba(255,255,255,0.1)' },
                            ticks: {
                                color: '#ffffff',
                                font: { size: 12 },
                                callback: function(value) {
                                    return 'R$ ' + value.toLocaleString('pt-BR', {minimumFractionDigits: 0, maximumFractionDigits: 0});
                                }
                            }
                        },
                        y: {
                            grid: { color: 'rgba(255,255,255,0.1)' },
                            ticks: {
                                color: '#ffffff',
                                font: { size: 12 }
                            }
                        }
                    }
                }
            });
        }

        // Gráfico de ticket médio
        if (document.getElementById('graficoTicketMedio')) {
            const ctxTicketMedio = document.getElementById('graficoTicketMedio').getContext('2d');
            const colors = getChartColors();
            
            window.ticketMedioChart = new Chart(ctxTicketMedio, {
                type: 'line',
                data: {
                    labels: mesesAnoData,
                    datasets: [{
                        label: 'Ticket Médio Mensal',
                        data: ticketMedioData.map(v => parseFloat(v) || 0),
                        backgroundColor: colors.background + '40',
                        borderColor: colors.border,
                        borderWidth: 3,
                        tension: 0.3,
                        fill: true
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            labels: {
                                color: '#ffffff',
                                font: { size: 14 }
                            }
                        },
                        tooltip: {
                            backgroundColor: 'rgba(0,0,0,0.8)',
                            titleColor: '#ffffff',
                            bodyColor: '#ffffff',
                            callbacks: {
                                label: function(context) {
                                    return 'R$ ' + context.parsed.y.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
                                }
                            }
                        }
                    },
                    scales: {
                        x: {
                            grid: { color: 'rgba(255,255,255,0.1)' },
                            ticks: {
                                color: '#ffffff',
                                font: { size: 12 }
                            }
                        },
                        y: {
                            beginAtZero: true,
                            grid: { color: 'rgba(255,255,255,0.1)' },
                            ticks: {
                                color: '#ffffff',
                                font: { size: 12 },
                                callback: function(value) {
                                    return 'R$ ' + value.toLocaleString('pt-BR', {minimumFractionDigits: 0, maximumFractionDigits: 0});
                                }
                            }
                        }
                    }
                }
            });
        }

        // Gráfico de tipo de unidade
        if (document.getElementById('graficoTipoUnidade') && Array.isArray(tiposUnidadeData) && tiposUnidadeData.length > 0) {
            const ctxTipoUnidade = document.getElementById('graficoTipoUnidade').getContext('2d');
            const colors = getChartColors();
            
            const tiposUnidadeValues = vendasTiposUnidadeData.map(v => parseFloat(v) || 0);
            const tipoColors = colors.vibrant.slice(0, tiposUnidadeData.length);
            
            window.tipoUnidadeChart = new Chart(ctxTipoUnidade, {
                type: 'doughnut',
                data: {
                    labels: tiposUnidadeData,
                    datasets: [{
                        data: tiposUnidadeValues,
                        backgroundColor: tipoColors,
                        borderColor: '#ffffff',
                        borderWidth: 2
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                color: '#ffffff',
                                font: { size: 12 },
                                padding: 15
                            }
                        },
                        tooltip: {
                            backgroundColor: 'rgba(0,0,0,0.8)',
                            titleColor: '#ffffff',
                            bodyColor: '#ffffff',
                            callbacks: {
                                label: function(context) {
                                    const label = context.label || '';
                                    const value = context.parsed;
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const percentage = total > 0 ? Math.round((value / total) * 100) : 0;
                                    return `${label}: R$ ${value.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})} (${percentage}%)`;
                                }
                            }
                        }
                    },
                    cutout: '50%'
                }
            });
        }
    }

    // Inicializar quando o DOM estiver carregado
    document.addEventListener('DOMContentLoaded', function() {
        initSlideshowControls();
        initCharts();
    });

    // Fallback caso o DOM já esteja carregado
    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        setTimeout(() => {
            initSlideshowControls();
            initCharts();
        }, 100);
    }
</script>
<!-- fim initChart -->
</body>
</html>

<%
' FECHAR A CONEXÃO APENAS NO FINAL DO ARQUIVO
If IsObject(conn) Then
    If conn.State = 1 Then
        conn.Close
    End If
    Set conn = Nothing
End If
%>