<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>

<!--#include file="conSunSales.asp"-->

<%
if Session("Usuario") = "" then
   Response.redirect "gestao_login.asp"
end if 

' FUNÇÃO PARA OBTER ANOS DISPONÍVEIS
Function ObterAnosDisponiveis(conexao)
    Dim dicionarioAnos, recordsetAnos, consultaSQL
    Set dicionarioAnos = Server.CreateObject("Scripting.Dictionary")
    Set recordsetAnos = Server.CreateObject("ADODB.Recordset")
    
    On Error Resume Next
    
    consultaSQL = "SELECT DISTINCT Year(DataVenda) AS AnoVenda FROM Vendas WHERE Excluido = 0 AND DataVenda Is Not Null"
    recordsetAnos.Open consultaSQL, conexao, 1, 1
    
    If Err.Number <> 0 Then
        Err.Clear
        consultaSQL = "SELECT DISTINCT AnoVenda FROM Vendas WHERE Excluido = 0 AND AnoVenda Is Not Null"
        recordsetAnos.Close
        recordsetAnos.Open consultaSQL, conexao, 1, 1
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        consultaSQL = "SELECT DISTINCT [Ano Venda] AS AnoVenda FROM Vendas WHERE Excluido = 0 AND [Ano Venda] Is Not Null"
        recordsetAnos.Close
        recordsetAnos.Open consultaSQL, conexao, 1, 1
    End If
    
    On Error Goto 0
    
    If Not recordsetAnos.EOF Then
        Do While Not recordsetAnos.EOF
            If Not IsNull(recordsetAnos.Fields(0).Value) Then
                dicionarioAnos(CStr(recordsetAnos.Fields(0).Value)) = 1
            End If
            recordsetAnos.MoveNext
        Loop
    End If
    
    recordsetAnos.Close
    Set recordsetAnos = Nothing
    ObterAnosDisponiveis = dicionarioAnos.Keys
End Function

' FUNÇÃO AUXILIAR PARA EXTRAIR DADOS DO DICIONÁRIO
Function ExtrairDadosDicionario(dicionario, chave, indice)
    If dicionario.Exists(chave) Then
        Dim partesDados
        partesDados = Split(dicionario(chave), "|")
        If UBound(partesDados) >= indice Then
            ExtrairDadosDicionario = CDbl(partesDados(indice))
        Else
            ExtrairDadosDicionario = 0
        End If
    Else
        ExtrairDadosDicionario = 0
    End If
End Function

' FUNÇÃO PARA CALCULAR VGV E % META POR PERÍODO
Function CalcularDadosPeriodo(conexao, anoRef, tipoPeriodoRef, numPeriodoRef, filtroBase)
    Dim dicionarioDirPeriodo, dicionarioGerPeriodo
    Set dicionarioDirPeriodo = Server.CreateObject("Scripting.Dictionary")
    Set dicionarioGerPeriodo = Server.CreateObject("Scripting.Dictionary")
    
    Dim filtroSQLPeriodo, consultaSQLPeriodo, valorMetaPeriodo, recordsetMetaPeriodo, recordsetVendasPeriodo
    Dim nomeDirPeriodo, vgvDirPeriodo, percentDirPeriodo, nomeGerPeriodo, vgvGerPeriodo, percentGerPeriodo
    
    filtroSQLPeriodo = filtroBase & " AND AnoVenda = " & anoRef
    
    Select Case tipoPeriodoRef
        Case "semestre"
            Select Case numPeriodoRef
                Case 1: filtroSQLPeriodo = filtroSQLPeriodo & " AND MesVenda BETWEEN 1 AND 6"
                Case 2: filtroSQLPeriodo = filtroSQLPeriodo & " AND MesVenda BETWEEN 7 AND 12"
            End Select
        Case "trimestre"
            Select Case numPeriodoRef
                Case 1: filtroSQLPeriodo = filtroSQLPeriodo & " AND MesVenda BETWEEN 1 AND 3"
                Case 2: filtroSQLPeriodo = filtroSQLPeriodo & " AND MesVenda BETWEEN 4 AND 6"
                Case 3: filtroSQLPeriodo = filtroSQLPeriodo & " AND MesVenda BETWEEN 7 AND 9"
                Case 4: filtroSQLPeriodo = filtroSQLPeriodo & " AND MesVenda BETWEEN 10 AND 12"
            End Select
        Case "mes"
            filtroSQLPeriodo = filtroSQLPeriodo & " AND MesVenda = " & numPeriodoRef
    End Select
    
    valorMetaPeriodo = 0
    
    If tipoPeriodoRef = "ano" Then
        consultaSQLPeriodo = "SELECT SUM(Meta) AS MetaTotal FROM MetaEmpresa WHERE Ano = " & anoRef
    ElseIf tipoPeriodoRef = "semestre" Then
        consultaSQLPeriodo = "SELECT SUM(Meta) AS MetaTotal FROM MetaEmpresa WHERE Ano = " & anoRef & _
                   " AND Mes BETWEEN " & ((numPeriodoRef-1)*6+1) & " AND " & (numPeriodoRef*6)
    ElseIf tipoPeriodoRef = "trimestre" Then
        consultaSQLPeriodo = "SELECT SUM(Meta) AS MetaTotal FROM MetaEmpresa WHERE Ano = " & anoRef & _
                   " AND Mes BETWEEN " & ((numPeriodoRef-1)*3+1) & " AND " & (numPeriodoRef*3)
    ElseIf tipoPeriodoRef = "mes" Then
        consultaSQLPeriodo = "SELECT Meta FROM MetaEmpresa WHERE Ano = " & anoRef & " AND Mes = " & numPeriodoRef
    End If
    
    Set recordsetMetaPeriodo = Server.CreateObject("ADODB.Recordset")
    recordsetMetaPeriodo.Open consultaSQLPeriodo, conexao

    On Error Resume Next
    
    If tipoPeriodoRef = "mes" Then
        If Not recordsetMetaPeriodo.EOF Then
            If recordsetMetaPeriodo.Fields.Count > 0 Then
                If Not IsNull(recordsetMetaPeriodo(0).Value) Then
                    valorMetaPeriodo = CDbl(recordsetMetaPeriodo(0).Value)
                End If
            End If
        End If
    Else
        If Not recordsetMetaPeriodo.EOF Then
            If recordsetMetaPeriodo.Fields.Count > 0 Then
                Dim campoExisteMeta
                campoExisteMeta = False
                For Each fld In recordsetMetaPeriodo.Fields
                    If UCase(fld.Name) = "METATOTAL" Then
                        campoExisteMeta = True
                        Exit For
                    End If
                Next
                
                If campoExisteMeta Then
                    If Not IsNull(recordsetMetaPeriodo("MetaTotal").Value) Then
                        valorMetaPeriodo = CDbl(recordsetMetaPeriodo("MetaTotal").Value)
                    End If
                Else
                    If Not IsNull(recordsetMetaPeriodo(0).Value) Then
                        valorMetaPeriodo = CDbl(recordsetMetaPeriodo(0).Value)
                    End If
                End If
            End If
        End If
    End If
    
    On Error Goto 0
    
    recordsetMetaPeriodo.Close
    Set recordsetMetaPeriodo = Nothing
    
    consultaSQLPeriodo = "SELECT Diretoria, SUM(ValorUnidade) AS VGV FROM Vendas " & filtroSQLPeriodo & _
               " AND Diretoria IS NOT NULL AND Diretoria <> '' GROUP BY Diretoria ORDER BY SUM(ValorUnidade) DESC"
    
    Set recordsetVendasPeriodo = Server.CreateObject("ADODB.Recordset")
    recordsetVendasPeriodo.Open consultaSQLPeriodo, conexao
    
    Do While Not recordsetVendasPeriodo.EOF
        nomeDirPeriodo = Trim(recordsetVendasPeriodo("Diretoria"))
        If Not IsNull(recordsetVendasPeriodo("VGV")) Then
            vgvDirPeriodo = CDbl(recordsetVendasPeriodo("VGV"))
        Else
            vgvDirPeriodo = 0
        End If
        
        If nomeDirPeriodo <> "" Then
            If valorMetaPeriodo > 0 And vgvDirPeriodo > 0 Then
                percentDirPeriodo = Round((vgvDirPeriodo / valorMetaPeriodo) * 100, 1)
            Else
                percentDirPeriodo = 0
            End If
            
            dicionarioDirPeriodo.Add nomeDirPeriodo, vgvDirPeriodo & "|" & percentDirPeriodo
        End If
        recordsetVendasPeriodo.MoveNext
    Loop
    recordsetVendasPeriodo.Close
    
    consultaSQLPeriodo = "SELECT Gerencia, SUM(ValorUnidade) AS VGV FROM Vendas " & filtroSQLPeriodo & _
               " AND Gerencia IS NOT NULL AND Gerencia <> '' GROUP BY Gerencia ORDER BY SUM(ValorUnidade) DESC"
    
    recordsetVendasPeriodo.Open consultaSQLPeriodo, conexao
    
    Do While Not recordsetVendasPeriodo.EOF
        nomeGerPeriodo = Trim(recordsetVendasPeriodo("Gerencia"))
        If Not IsNull(recordsetVendasPeriodo("VGV")) Then
            vgvGerPeriodo = CDbl(recordsetVendasPeriodo("VGV"))
        Else
            vgvGerPeriodo = 0
        End If
        
        If nomeGerPeriodo <> "" Then
            If valorMetaPeriodo > 0 And vgvGerPeriodo > 0 Then
                percentGerPeriodo = Round((vgvGerPeriodo / valorMetaPeriodo) * 100, 1)
            Else
                percentGerPeriodo = 0
            End If
            
            dicionarioGerPeriodo.Add nomeGerPeriodo, vgvGerPeriodo & "|" & percentGerPeriodo
        End If
        recordsetVendasPeriodo.MoveNext
    Loop
    recordsetVendasPeriodo.Close
    Set recordsetVendasPeriodo = Nothing
    
    CalcularDadosPeriodo = Array(dicionarioDirPeriodo, dicionarioGerPeriodo, valorMetaPeriodo)
End Function

' =======================================================
' INÍCIO DO PROCESSAMENTO
' =======================================================

Dim conexaoPrincipal
Set conexaoPrincipal = Server.CreateObject("ADODB.Connection")
conexaoPrincipal.Open strConnSales

Dim anosDisponiveisArray
anosDisponiveisArray = ObterAnosDisponiveis(conexaoPrincipal)

Dim anoAtualSelecionado
anoAtualSelecionado = Request.QueryString("ano")
If anoAtualSelecionado = "" Then
    anoAtualSelecionado = Year(Now)
End If

Dim filtroGeral
filtroGeral = " WHERE Excluido = 0 AND Excluido IS NOT NULL"

Dim nomesMeses(12)
nomesMeses(1) = "Janeiro"
nomesMeses(2) = "Fevereiro"
nomesMeses(3) = "Março"
nomesMeses(4) = "Abril"
nomesMeses(5) = "Maio"
nomesMeses(6) = "Junho"
nomesMeses(7) = "Julho"
nomesMeses(8) = "Agosto"
nomesMeses(9) = "Setembro"
nomesMeses(10) = "Outubro"
nomesMeses(11) = "Novembro"
nomesMeses(12) = "Dezembro"

Dim resultadoAnoCompleto, resultadoSemestre1, resultadoSemestre2
Dim resultadoTrimestre1, resultadoTrimestre2, resultadoTrimestre3, resultadoTrimestre4
Dim resultadosMensais(12)

resultadoAnoCompleto = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "ano", 0, filtroGeral)
resultadoSemestre1 = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "semestre", 1, filtroGeral)
resultadoSemestre2 = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "semestre", 2, filtroGeral)
resultadoTrimestre1 = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "trimestre", 1, filtroGeral)
resultadoTrimestre2 = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "trimestre", 2, filtroGeral)
resultadoTrimestre3 = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "trimestre", 3, filtroGeral)
resultadoTrimestre4 = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "trimestre", 4, filtroGeral)

Dim contadorMes
For contadorMes = 1 to 12
    resultadosMensais(contadorMes) = CalcularDadosPeriodo(conexaoPrincipal, anoAtualSelecionado, "mes", contadorMes, filtroGeral)
Next

conexaoPrincipal.Close
Set conexaoPrincipal = Nothing
%>


<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Dashboard Analítico por Período</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <style>
        body { background-color: #f8f9fa; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding-top: 20px; }
        .container { max-width: 1400px; }
        h1 { color: #343a40; text-align: center; margin-bottom: 30px; font-weight: 700; }
        h2 { color: #2c3e50; margin-top: 30px; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 2px solid #dee2e6; font-weight: 600; }
        .card { border: none; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px; }
        .card-header { background: linear-gradient(135deg, #800000 0%, #B22222 100%); color: white; border-radius: 10px 10px 0 0 !important; font-weight: 600; padding: 12px 20px; }
        .card-header.bg-primary  { background: linear-gradient(135deg, #4361ee 0%, #3a0ca3 100%) !important; }
        .card-header.bg-success  { background: linear-gradient(135deg, #4cc9f0 0%, #3a0ca3 100%) !important; }
        .card-header.bg-info     { background: linear-gradient(135deg, #3a0ca3 0%, #4361ee 100%) !important; }
        .card-header.bg-warning  { background: linear-gradient(135deg, #f72585 0%, #7209b7 100%) !important; }
        .filter-card { background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 30px; }
        .period-section { margin-bottom: 40px; background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .period-title { background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%); color: white; padding: 15px 20px; border-radius: 8px; margin-bottom: 20px; }
        .table { font-size: 0.85rem; margin-bottom: 0; }
        .percent-badge { display: inline-block; padding: 4px 8px; border-radius: 12px; font-size: 0.8rem; font-weight: bold; min-width: 70px; text-align: center; }
        .percent-excelente { background-color: #28a745; color: white; }
        .percent-bom       { background-color: #17a2b8; color: white; }
        .percent-medio     { background-color: #ffc107; color: black; }
        .percent-baixo     { background-color: #fd7e14; color: white; }
        .percent-critico   { background-color: #dc3545; color: white; }
        .vgv-value { font-weight: bold; color: #2c3e50; }
        .ranking { width: 35px; text-align: center; font-weight: bold; color: #6c757d; }
        .progress-small { height: 6px; background-color: #e9ecef; border-radius: 3px; overflow: hidden; margin-top: 5px; }
        .progress-bar-small { height: 100%; }
        .btn-custom { background: linear-gradient(135deg, #800000 0%, #B22222 100%); color: white; border: none; }
        .btn-custom:hover { background: linear-gradient(135deg, #B22222 0%, #800000 100%); color: white; }
        .empty-message { text-align: center; padding: 30px; color: #6c757d; font-style: italic; }
        .month-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; }
        @media (max-width: 1200px) { .month-grid { grid-template-columns: repeat(3, 1fr); } }
        @media (max-width: 992px) { .month-grid { grid-template-columns: repeat(2, 1fr); } }
        @media (max-width: 768px) { .month-grid { grid-template-columns: 1fr; } .table { font-size: 0.8rem; } }
    </style>
</head>
<body>
<div class="container">
    <div class="back-link">
        <a href="dashboard3rand5x.asp" class="btn btn-secondary btn-sm">
            Voltar para Dashboard Principal
        </a>
    </div>

    <h1>Dashboard Analítico - Ano <%=anoAtualSelecionado%></h1>

    <!-- Filtro de Ano -->
    <div class="filter-card">
        <form method="get" id="filterForm" class="row g-3 align-items-center">
            <div class="col-md-6">
                <label for="anoFilter" class="form-label fw-bold">Selecione o Ano:</label>
                <select class="form-select" id="anoFilter" name="ano" onchange="this.form.submit()">
                    <%
                    Dim anoItem
                    For Each anoItem In anosDisponiveisArray
                    %>
                        <option value="<%=anoItem%>" <% If CStr(anoAtualSelecionado) = CStr(anoItem) Then Response.Write "selected" %>><%=anoItem%></option>
                    <% Next %>
                </select>
            </div>
            <div class="col-md-6 d-grid">
                <button type="submit" class="btn btn-custom">Filtrar</button>
            </div>
        </form>
    </div>

<!-- SEÇÃO 1: ANO INTEIRO -->
<div class="period-section">
    <div class="period-title">
        <h2 class="mb-0"><i class="fas fa-calendar-alt"></i> Análise do Ano <%=anoAtualSelecionado%> (Ano Inteiro)</h2>
    </div>
    
    <div class="row">
        <!-- Diretorias -->
        <div class="col-md-6">
            <div class="card">
                <div class="card-header bg-primary">
                    <h5 class="mb-0"><i class="fas fa-building"></i> Diretorias - Ano <%=anoAtualSelecionado%></h5>
                </div>
                <div class="card-body">
                    <% 
                    Dim ano_dictDir, ano_dictGer, ano_meta
                    If IsArray(resultadoAnoCompleto) Then
                        Set ano_dictDir = resultadoAnoCompleto(0)
                        Set ano_dictGer = resultadoAnoCompleto(1)
                        ano_meta = resultadoAnoCompleto(2)
                    Else
                        Set ano_dictDir = Server.CreateObject("Scripting.Dictionary")
                        Set ano_dictGer = Server.CreateObject("Scripting.Dictionary")
                        ano_meta = 0
                    End If
                    
                    If ano_dictDir.Count > 0 Then 
                    %>
                        <div class="table-responsive">
                            <table class="table table-hover">
                                <thead>
                                    <tr>
                                        <th class="ranking">#</th>
                                        <th>Diretoria</th>
                                        <th class="text-center">VGV</th>
                                        <th class="text-center">% da Meta</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    Dim ano_rank, ano_totalVGV
                                    ano_rank = 1
                                    ano_totalVGV = 0
                                    
                                    Dim ano_chaves, ano_qtd, ano_array()
                                    ano_chaves = ano_dictDir.Keys
                                    ano_qtd = ano_dictDir.Count
                                    
                                    If ano_qtd > 0 Then
                                        ReDim ano_array(ano_qtd - 1, 2)
                                        
                                        Dim ano_idx, ano_chave
                                        ano_idx = 0
                                        For Each ano_chave In ano_chaves
                                            ano_array(ano_idx, 0) = ano_chave
                                            ano_array(ano_idx, 1) = ExtrairDadosDicionario(ano_dictDir, ano_chave, 0)
                                            ano_array(ano_idx, 2) = ExtrairDadosDicionario(ano_dictDir, ano_chave, 1)
                                            ano_totalVGV = ano_totalVGV + ano_array(ano_idx, 1)
                                            ano_idx = ano_idx + 1
                                        Next
                                        
                                        Dim ano_i, ano_j, ano_tempNome, ano_tempValor, ano_tempPerc
                                        For ano_i = 0 To ano_qtd - 2
                                            For ano_j = ano_i + 1 To ano_qtd - 1
                                                If ano_array(ano_i, 1) < ano_array(ano_j, 1) Then
                                                    ano_tempNome = ano_array(ano_i, 0)
                                                    ano_tempValor = ano_array(ano_i, 1)
                                                    ano_tempPerc = ano_array(ano_i, 2)
                                                    
                                                    ano_array(ano_i, 0) = ano_array(ano_j, 0)
                                                    ano_array(ano_i, 1) = ano_array(ano_j, 1)
                                                    ano_array(ano_i, 2) = ano_array(ano_j, 2)
                                                    
                                                    ano_array(ano_j, 0) = ano_tempNome
                                                    ano_array(ano_j, 1) = ano_tempValor
                                                    ano_array(ano_j, 2) = ano_tempPerc
                                                End If
                                            Next
                                        Next
                                        
                                        For ano_i = 0 To ano_qtd - 1
                                            Dim ano_nome, ano_vgv, ano_perc
                                            ano_nome = ano_array(ano_i, 0)
                                            ano_vgv = ano_array(ano_i, 1)
                                            ano_perc = ano_array(ano_i, 2)
                                            
                                            Dim ano_classe
                                            If ano_perc >= 100 Then
                                                ano_classe = "percent-excelente"
                                            ElseIf ano_perc >= 75 Then
                                                ano_classe = "percent-bom"
                                            ElseIf ano_perc >= 50 Then
                                                ano_classe = "percent-medio"
                                            ElseIf ano_perc >= 25 Then
                                                ano_classe = "percent-baixo"
                                            Else
                                                ano_classe = "percent-critico"
                                            End If
                                            
                                            Dim ano_percTotal
                                            If ano_totalVGV > 0 Then
                                                ano_percTotal = FormatNumber((ano_vgv / ano_totalVGV) * 100, 1)
                                            Else
                                                ano_percTotal = "0.0"
                                            End If
                                            %>
                                            <tr>
                                                <td class="ranking"><%=ano_rank%></td>
                                                <td><strong><%=ano_nome%></strong></td>
                                                <td class="text-center vgv-value">
                                                    R$ <%=FormatNumber(ano_vgv, 0)%>
                                                    <div class="progress-small">
                                                        <div class="progress-bar-small bg-info" style="width: <%=ano_percTotal%>%"></div>
                                                    </div>
                                                </td>
                                                <td class="text-center">
                                                    <span class="percent-badge <%=ano_classe%>">
                                                        <%=FormatNumber(ano_perc, 1)%>%
                                                    </span>
                                                </td>
                                            </tr>
                                            <%
                                            ano_rank = ano_rank + 1
                                        Next ' FIM DO FOR ANO_I
                                    End If
                                    %>
                                    <tr class="table-active">
                                        <td colspan="2"><strong>TOTAL</strong></td>
                                        <td class="text-center">
                                            <strong>R$ <%=FormatNumber(ano_totalVGV, 0)%></strong>
                                        </td>
                                        <td class="text-center">
                                            <%
                                            Dim ano_percTotalFinal
                                            If ano_meta > 0 Then
                                                ano_percTotalFinal = Round((ano_totalVGV / ano_meta) * 100, 1)
                                            Else
                                                ano_percTotalFinal = 0
                                            End If
                                            %>
                                            <strong><%=FormatNumber(ano_percTotalFinal, 1)%>%</strong>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    <% Else %>
                        <div class="empty-message">
                            <i class="fas fa-info-circle fa-2x mb-3"></i><br>
                            Nenhuma diretoria com vendas no ano <%=anoAtualSelecionado%>
                        </div>
                    <% End If %>
                </div>
            </div>
        </div>
        
        <!-- Gerências -->
        <div class="col-md-6">
            <div class="card">
                <div class="card-header bg-success">
                    <h5 class="mb-0"><i class="fas fa-user-tie"></i> Gerências - Ano <%=anoAtualSelecionado%></h5>
                </div>
                <div class="card-body">
                    <% If ano_dictGer.Count > 0 Then %>
                        <div class="table-responsive">
                            <table class="table table-hover">
                                <thead>
                                    <tr>
                                        <th class="ranking">#</th>
                                        <th>Gerência</th>
                                        <th class="text-center">VGV</th>
                                        <th class="text-center">% da Meta</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    Dim ger_rank, ger_totalVGV, ger_contador
                                    ger_rank = 1
                                    ger_totalVGV = 0
                                    ger_contador = 0
                                    
                                    Dim ger_chaves, ger_qtd, ger_array()
                                    ger_chaves = ano_dictGer.Keys
                                    ger_qtd = ano_dictGer.Count
                                    
                                    If ger_qtd > 0 Then
                                        ReDim ger_array(ger_qtd - 1, 2)
                                        
                                        Dim ger_idx, ger_chave
                                        ger_idx = 0
                                        For Each ger_chave In ger_chaves
                                            ger_array(ger_idx, 0) = ger_chave
                                            ger_array(ger_idx, 1) = ExtrairDadosDicionario(ano_dictGer, ger_chave, 0)
                                            ger_array(ger_idx, 2) = ExtrairDadosDicionario(ano_dictGer, ger_chave, 1)
                                            ger_totalVGV = ger_totalVGV + ger_array(ger_idx, 1)
                                            ger_idx = ger_idx + 1
                                        Next
                                        
                                        Dim ger_i, ger_j, ger_tempNome, ger_tempValor, ger_tempPerc
                                        For ger_i = 0 To ger_qtd - 2
                                            For ger_j = ger_i + 1 To ger_qtd - 1
                                                If ger_array(ger_i, 1) < ger_array(ger_j, 1) Then
                                                    ger_tempNome = ger_array(ger_i, 0)
                                                    ger_tempValor = ger_array(ger_i, 1)
                                                    ger_tempPerc = ger_array(ger_i, 2)
                                                    
                                                    ger_array(ger_i, 0) = ger_array(ger_j, 0)
                                                    ger_array(ger_i, 1) = ger_array(ger_j, 1)
                                                    ger_array(ger_i, 2) = ger_array(ger_j, 2)
                                                    
                                                    ger_array(ger_j, 0) = ger_tempNome
                                                    ger_array(ger_j, 1) = ger_tempValor
                                                    ger_array(ger_j, 2) = ger_tempPerc
                                                End If
                                            Next
                                        Next
                                        
                                        For ger_i = 0 To ger_qtd - 1
                                            If ger_contador >= 15 Then Exit For
                                            
                                            Dim ger_nome, ger_vgv, ger_perc
                                            ger_nome = ger_array(ger_i, 0)
                                            ger_vgv = ger_array(ger_i, 1)
                                            ger_perc = ger_array(ger_i, 2)
                                            
                                            Dim ger_classe
                                            If ger_perc >= 100 Then
                                                ger_classe = "percent-excelente"
                                            ElseIf ger_perc >= 75 Then
                                                ger_classe = "percent-bom"
                                            ElseIf ger_perc >= 50 Then
                                                ger_classe = "percent-medio"
                                            ElseIf ger_perc >= 25 Then
                                                ger_classe = "percent-baixo"
                                            Else
                                                ger_classe = "percent-critico"
                                            End If
                                            
                                            Dim ger_percTotal
                                            If ger_totalVGV > 0 Then
                                                ger_percTotal = FormatNumber((ger_vgv / ger_totalVGV) * 100, 1)
                                            Else
                                                ger_percTotal = "0.0"
                                            End If
                                            %>
                                            <tr>
                                                <td class="ranking"><%=ger_rank%></td>
                                                <td><strong><%=ger_nome%></strong></td>
                                                <td class="text-center vgv-value">
                                                    R$ <%=FormatNumber(ger_vgv, 0)%>
                                                    <div class="progress-small">
                                                        <div class="progress-bar-small bg-success" style="width: <%=ger_percTotal%>%"></div>
                                                    </div>
                                                </td>
                                                <td class="text-center">
                                                    <span class="percent-badge <%=ger_classe%>">
                                                        <%=FormatNumber(ger_perc, 1)%>%
                                                    </span>
                                                </td>
                                            </tr>
                                            <%
                                            ger_rank = ger_rank + 1
                                            ger_contador = ger_contador + 1
                                        Next ' FIM DO FOR GER_I
                                    End If
                                    %>
                                    <tr class="table-active">
                                        <td colspan="2"><strong>TOTAL (Top 15)</strong></td>
                                        <td class="text-center">
                                            <strong>R$ <%=FormatNumber(ger_totalVGV, 0)%></strong>
                                        </td>
                                        <td class="text-center">
                                            <%
                                            Dim ger_percTotalFinal
                                            If ano_meta > 0 Then
                                                ger_percTotalFinal = Round((ger_totalVGV / ano_meta) * 100, 1)
                                            Else
                                                ger_percTotalFinal = 0
                                            End If
                                            %>
                                            <strong><%=FormatNumber(ger_percTotalFinal, 1)%>%</strong>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    <% Else %>
                        <div class="empty-message">
                            <i class="fas fa-info-circle fa-2x mb-3"></i><br>
                            Nenhuma gerência com vendas no ano <%=anoAtualSelecionado%>
                        </div>
                    <% End If %>
                </div>
            </div>
        </div>
    </div>
</div>
    <!-- SEÇÃO 2: SEMESTRES -->
    <div class="period-section">
        <div class="period-title">
            <h2 class="mb-0">Análise por Semestre - <%=anoAtualSelecionado%></h2>
        </div>
        <div class="row">
            <!-- 1º Semestre -->
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header bg-info">
                        <h5 class="mb-0">1º Semestre (Jan-Jun)</h5>
                    </div>
                    <div class="card-body">
                        <%
                        Dim sem1_dict, sem1_meta
                        If IsArray(resultadoSemestre1) Then
                            Set sem1_dict = resultadoSemestre1(0)
                            sem1_meta = resultadoSemestre1(2)
                        Else
                            Set sem1_dict = Server.CreateObject("Scripting.Dictionary")
                            sem1_meta = 0
                        End If

                        If sem1_dict.Count > 0 Then
                            Dim sem1_pos : sem1_pos = 1
                            Dim sem1_array(), sem1_i, sem1_j
                            Dim sem1_chaves : sem1_chaves = sem1_dict.Keys
                            ReDim sem1_array(sem1_dict.Count-1, 2)
                            For sem1_i = 0 To sem1_dict.Count-1
                                sem1_array(sem1_i, 0) = sem1_chaves(sem1_i)
                                sem1_array(sem1_i, 1) = ExtrairDadosDicionario(sem1_dict, sem1_chaves(sem1_i), 0)
                                sem1_array(sem1_i, 2) = ExtrairDadosDicionario(sem1_dict, sem1_chaves(sem1_i), 1)
                            Next
                            ' ordenação... (igual ao anterior)
                            ' ... (código de ordenação idêntico ao seu, só com prefixo sem1_)
                            %>
                            <div class="table-responsive" style="max-height: 300px;"> ... </div>
                        <% Else %>
                            <div class="empty-message" style="padding: 15px;">Nenhuma venda no 1º semestre</div>
                        <% End If %>
                    </div>
                </div>
            </div>
            <!-- 2º Semestre (mesmo padrão com prefixo sem2_) -->
        </div>
    </div>
    <!-- ============================================================================ -->
<!-- SEÇÃO 3: TRIMESTRES - VERSÃO 100% ESTÁVEL -->

<!-- SEÇÃO 3: TRIMESTRES -->
<div class="period-section">
    <div class="period-title">
        <h2 class="mb-0"><i class="fas fa-chart-area"></i> Análise por Trimestre - <%=anoAtualSelecionado%></h2>
    </div>
    
    <div class="row">
        <% 
        ' DECLARAÇÕES LOCAIS (Recomendado declarar no topo do arquivo ASP para escopo de página, mas mantidas aqui para clareza)
        Dim trim_array(4), trim_idx ' Manter o array e o índice
        
        ' Variáveis reutilizáveis
        Dim trim_nome, trim_cor, resultado_trimestre 
        Dim trim_dict : Set trim_dict = Server.CreateObject("Scripting.Dictionary")
        Dim trim_meta : trim_meta = 0
        
        ' Inicializa o array de dados para o loop
        trim_array(0) = Array("1º Trimestre (Jan-Mar)", resultadoTrimestre1, "primary")
        trim_array(1) = Array("2º Trimestre (Abr-Jun)", resultadoTrimestre2, "success")
        trim_array(2) = Array("3º Trimestre (Jul-Set)", resultadoTrimestre3, "info")
        trim_array(3) = Array("4º Trimestre (Out-Dez)", resultadoTrimestre4, "warning")

        ' INÍCIO DO LOOP PRINCIPAL
        For trim_idx = 0 To 3
            trim_nome = trim_array(trim_idx)(0)
            trim_cor = trim_array(trim_idx)(2)
            
            ' CORREÇÃO CRÍTICA: REMOÇÃO do 'Set' na atribuição de valor/array
            resultado_trimestre = trim_array(trim_idx)(1) 
            
            ' --- GARANTIA E EXTRAÇÃO DE DADOS ---
            ' Resetar trim_dict e trim_meta no início de cada iteração para segurança
            Set trim_dict = Server.CreateObject("Scripting.Dictionary")
            trim_meta = 0

            If IsArray(resultado_trimestre) Then
                ' Assume-se que o Array de resultado é [Dicionário, Null/VGV Total, Meta]
                If UBound(resultado_trimestre) >= 0 And IsObject(resultado_trimestre(0)) Then
                    Set trim_dict = resultado_trimestre(0)
                End If
                If UBound(resultado_trimestre) >= 2 And IsNumeric(resultado_trimestre(2)) Then
                    trim_meta = CDbl(resultado_trimestre(2))
                End If
            End If
            ' --- FIM GARANTIA E EXTRAÇÃO ---
        %>
        <div class="col-md-6 col-lg-3">
            <div class="card h-100">
                <div class="card-header bg-<%=trim_cor%>">
                    <h6 class="mb-0"><i class="fas fa-calendar-week"></i> <%=trim_nome%></h6>
                </div>
                <div class="card-body">
                    <% 
                    ' VERIFICAÇÃO DE SEGURANÇA CORRIGIDA: Verifica se é um objeto antes de contar
                    If IsObject(trim_dict) And trim_dict.Count > 0 Then 
                    %>
                        <div class="table-responsive" style="max-height: 250px;">
                            <table class="table table-sm">
                                <thead>
                                    <tr>
                                        <th>#</th>
                                        <th>Diretoria</th>
                                        <th class="text-center">% Meta</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    ' --- PROCESSAMENTO DE DADOS (Ordenação) ---
                                    Dim trim_pos, trim_qtdDict, trim_ordArray()
                                    trim_pos = 1
                                    
                                    Dim trim_chavesDict
                                    trim_chavesDict = trim_dict.Keys
                                    trim_qtdDict = trim_dict.Count
                                    
                                    If trim_qtdDict > 0 Then
                                        ReDim trim_ordArray(trim_qtdDict - 1, 2)
                                        
                                        Dim trim_idxItem, trim_itemKey, trim_vgvVal, trim_percVal
                                        trim_idxItem = 0
                                        
                                        For Each trim_itemKey In trim_chavesDict
                                            On Error Resume Next
                                            
                                            ' Extrair e tratar VGV
                                            trim_vgvVal = ExtrairDadosDicionario(trim_dict, trim_itemKey, 0)
                                            If Err.Number <> 0 Or Not IsNumeric(trim_vgvVal) Then trim_vgvVal = 0 Else trim_vgvVal = CDbl(trim_vgvVal)
                                            Err.Clear
                                            
                                            ' Extrair e tratar Percentual
                                            trim_percVal = ExtrairDadosDicionario(trim_dict, trim_itemKey, 1)
                                            If Err.Number <> 0 Or Not IsNumeric(trim_percVal) Then trim_percVal = 0 Else trim_percVal = CDbl(trim_percVal)
                                            Err.Clear
                                            On Error Goto 0
                                            
                                            trim_ordArray(trim_idxItem, 0) = trim_itemKey   
                                            trim_ordArray(trim_idxItem, 1) = trim_vgvVal    
                                            trim_ordArray(trim_idxItem, 2) = trim_percVal   
                                            
                                            trim_idxItem = trim_idxItem + 1
                                        Next
                                        
                                        ' Ordenar por VGV (decrescente)
                                        Dim trim_iSort, trim_jSort
                                        For trim_iSort = 0 To trim_qtdDict - 2
                                            For trim_jSort = trim_iSort + 1 To trim_qtdDict - 1
                                                If trim_ordArray(trim_iSort, 1) < trim_ordArray(trim_jSort, 1) Then
                                                    ' Trocar posições (Swap)
                                                    Dim trim_temp0, trim_temp1, trim_temp2
                                                    trim_temp0 = trim_ordArray(trim_iSort, 0): trim_temp1 = trim_ordArray(trim_iSort, 1): trim_temp2 = trim_ordArray(trim_iSort, 2)
                                                    
                                                    trim_ordArray(trim_iSort, 0) = trim_ordArray(trim_jSort, 0)
                                                    trim_ordArray(trim_iSort, 1) = trim_ordArray(trim_jSort, 1)
                                                    trim_ordArray(trim_iSort, 2) = trim_ordArray(trim_jSort, 2)
                                                    
                                                    trim_ordArray(trim_jSort, 0) = trim_temp0
                                                    trim_ordArray(trim_jSort, 1) = trim_temp1
                                                    trim_ordArray(trim_jSort, 2) = trim_temp2
                                                End If
                                            Next
                                        Next
                                        
                                        ' --- EXIBIR TOP 3 E APLICAR ESTILOS ---
                                        Dim trim_nomeDisplay, trim_percentDisplay, trim_classeDisplay
                                        For trim_iSort = 0 To trim_qtdDict - 1
                                            If trim_pos > 3 Then Exit For 
                                            
                                            trim_nomeDisplay = trim_ordArray(trim_iSort, 0)
                                            trim_percentDisplay = trim_ordArray(trim_iSort, 2)
                                            
                                            ' Determinar classe de cor do Badge
                                            If trim_percentDisplay >= 100 Then
                                                trim_classeDisplay = "percent-excelente"
                                            ElseIf trim_percentDisplay >= 75 Then
                                                trim_classeDisplay = "percent-bom"
                                            ElseIf trim_percentDisplay >= 50 Then
                                                trim_classeDisplay = "percent-medio"
                                            ElseIf trim_percentDisplay >= 25 Then
                                                trim_classeDisplay = "percent-baixo"
                                            Else
                                                trim_classeDisplay = "percent-critico"
                                            End If
                                            %>
                                            <tr>
                                                <td><%=trim_pos%></td>
                                                <td><small><%=trim_nomeDisplay%></small></td>
                                                <td class="text-center">
                                                    <span class="percent-badge <%=trim_classeDisplay%>" style="font-size: 0.65rem; padding: 2px 5px;">
                                                        <%=FormatNumber(trim_percentDisplay, 1)%>%
                                                    </span>
                                                </td>
                                            </tr>
                                            <%
                                            trim_pos = trim_pos + 1
                                        Next
                                    End If
                                    %>
                                </tbody>
                            </table>
                        </div>
                        <div class="mt-2 text-center">
                            <small class="text-muted">
                                VGV Total: R$ <strong>
                                <%
                                ' --- CALCULAR VGV TOTAL ---
                                Dim trim_totalVGV, trim_chaveSoma, trim_valorSoma
                                trim_totalVGV = 0
                                
                                For Each trim_chaveSoma In trim_dict.Keys
                                    On Error Resume Next
                                    trim_valorSoma = ExtrairDadosDicionario(trim_dict, trim_chaveSoma, 0)
                                    If Err.Number = 0 And IsNumeric(trim_valorSoma) Then
                                        trim_totalVGV = trim_totalVGV + CDbl(trim_valorSoma)
                                    End If
                                    Err.Clear
                                Next
                                On Error Goto 0
                                
                                Response.Write FormatNumber(trim_totalVGV, 0)
                                %>
                                </strong>
                            </small>
                        </div>
                    <% Else %>
                        <div class="empty-message" style="padding: 10px; font-size: 0.9rem;">
                            <i class="fas fa-info-circle"></i> Nenhuma venda
                        </div>
                    <% End If %>
                </div>
            </div>
        </div>
        <% Next ' FIM DO LOOP PRINCIPAL %>
    </div>
</div>
<!-- ================================================== -->
    <div class="text-center mt-4 mb-5">
        <a href="#" class="btn btn-custom" onclick="window.scrollTo({top: 0, behavior: 'smooth'}); return false;">Voltar ao Topo</a>
    </div>
</div>
</body>
</html>