<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: BBKVHDODGV          -->
<!-- OBS: Painel Bolsa de Valores - CORRETORES -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' Adicionar Option Explicit no início para forçar declaração de variáveis

%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<%
if Session("Usuario") = "" then
   Response.redirect "gestao_login.asp"
end if 
%>

<%
' ===============================================
' CONFIGURAÇÕES INICIAIS
' ===============================================
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"

Dim conn, connSales
Set conn = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")
conn.Open StrConn
connSales.Open StrConnSales

' ===============================================
' FUNÇÃO SIMPLES PARA VALOR
' ===============================================
Function GetNum(val)
    On Error Resume Next
    If IsNull(val) Or Trim(val & "") = "" Then
        GetNum = 0
    ElseIf IsNumeric(val) Then
        GetNum = CDbl(val)
    Else
        GetNum = 0
    End If
    On Error GoTo 0
End Function

' ===============================================
' FILTROS
' ===============================================
Dim anoFiltro, mesFiltro, corretorFiltro, gerenciaFiltro
anoFiltro = Request.QueryString("ano")
mesFiltro = Request.QueryString("mes")
corretorFiltro = Request.QueryString("corretor")
gerenciaFiltro = Request.QueryString("gerencia")

If anoFiltro = "" Or Not IsNumeric(anoFiltro) Then 
    anoFiltro = Year(Date())
Else
    anoFiltro = CInt(anoFiltro)
End If

' ===============================================
' BUSCA DADOS DE CORRETORES DA TABELA VENDAS
' ===============================================
Dim sql, rs
sql = "SELECT "
sql = sql & "CorretorId, "  ' ID do corretor
sql = sql & "Corretor, "    ' Nome do corretor
sql = sql & "Gerencia, "
sql = sql & "MesVenda, "
sql = sql & "ValorUnidade "
sql = sql & "FROM Vendas "
sql = sql & "WHERE Excluido = 0 AND AnoVenda = " & anoFiltro
sql = sql & " AND ValorUnidade IS NOT NULL"

If corretorFiltro <> "" Then
    sql = sql & " AND CorretorId = '" & Replace(corretorFiltro, "'", "''") & "'"
End If

If gerenciaFiltro <> "" Then
    sql = sql & " AND Gerencia = '" & Replace(gerenciaFiltro, "'", "''") & "'"
End If

If mesFiltro <> "" And IsNumeric(mesFiltro) Then
    sql = sql & " AND MesVenda = " & CInt(mesFiltro)
End If

sql = sql & " ORDER BY CorretorId, MesVenda"
'response.Write sql
'response.end 

Set rs = connSales.Execute(sql)

' ===============================================
' ARRAYS PARA ARMAZENAR DADOS DOS CORRETORES
' ===============================================
Dim corretorCodigos(), corretorNomes(), corretorGerencias()
Dim corretorTotais(), corretorVariacoes()
Dim corretorMeses1(), corretorMeses2(), corretorMeses3()
Dim corretorMeses4(), corretorMeses5(), corretorMeses6()
Dim corretorMeses7(), corretorMeses8(), corretorMeses9()
Dim corretorMeses10(), corretorMeses11(), corretorMeses12()

Dim count, totalGeral, altas, baixas
count = 0
totalGeral = 0
altas = 0
baixas = 0

Dim currentCorretor, lastCorretor
Dim meses(12), totalCorretor
Dim lastVal1, lastVal2, lastMes1, lastMes2
Dim corretorNome, corretorGerencia

If Not rs.EOF Then
    lastCorretor = ""
    
    Do While Not rs.EOF
        Dim cCodigo, cNome, cGerencia, cMes, cValor
        cCodigo = Trim(rs("CorretorId"))
        If Not IsNull(rs("Corretor")) Then
            cNome = Trim(rs("Corretor"))
        Else
            cNome = cCodigo
        End If
        If Not IsNull(rs("Gerencia")) Then
            cGerencia = Trim(rs("Gerencia"))
        Else
            cGerencia = "Sem Gerência"
        End If
        cMes = GetNum(rs("MesVenda"))
        cValor = GetNum(rs("ValorUnidade"))
        
        ' Novo corretor?
        If lastCorretor <> cCodigo Then
            ' Processa corretor anterior se existir
            If lastCorretor <> "" Then
                ' Calcula total
                totalCorretor = 0
                For i = 1 To 12
                    totalCorretor = totalCorretor + meses(i)
                Next
                
                ' Encontra últimos valores
                lastVal1 = 0
                lastVal2 = 0
                lastMes1 = 0
                lastMes2 = 0
                
                For i = 12 To 1 Step -1
                    If meses(i) > 0 Then
                        If lastMes1 = 0 Then
                            lastMes1 = i
                            lastVal1 = meses(i)
                        ElseIf lastMes2 = 0 Then
                            lastMes2 = i
                            lastVal2 = meses(i)
                            Exit For
                        End If
                    End If
                Next
                
                ' Calcula variação
                Dim variacao
                variacao = 0
                If lastVal2 > 0 Then
                    variacao = ((lastVal1 - lastVal2) / lastVal2) * 100
                ElseIf lastVal1 > 0 Then
                    variacao = 100
                ElseIf lastVal1 = 0 And lastVal2 > 0 Then
                    variacao = -100
                End If
                
                ' Armazena nos arrays
                ReDim Preserve corretorCodigos(count)
                ReDim Preserve corretorNomes(count)
                ReDim Preserve corretorGerencias(count)
                ReDim Preserve corretorTotais(count)
                ReDim Preserve corretorVariacoes(count)
                
                ReDim Preserve corretorMeses1(count)
                ReDim Preserve corretorMeses2(count)
                ReDim Preserve corretorMeses3(count)
                ReDim Preserve corretorMeses4(count)
                ReDim Preserve corretorMeses5(count)
                ReDim Preserve corretorMeses6(count)
                ReDim Preserve corretorMeses7(count)
                ReDim Preserve corretorMeses8(count)
                ReDim Preserve corretorMeses9(count)
                ReDim Preserve corretorMeses10(count)
                ReDim Preserve corretorMeses11(count)
                ReDim Preserve corretorMeses12(count)
                
                corretorCodigos(count) = lastCorretor
                corretorNomes(count) = corretorNome
                corretorGerencias(count) = corretorGerencia
                corretorTotais(count) = totalCorretor
                corretorVariacoes(count) = variacao
                
                corretorMeses1(count) = meses(1)
                corretorMeses2(count) = meses(2)
                corretorMeses3(count) = meses(3)
                corretorMeses4(count) = meses(4)
                corretorMeses5(count) = meses(5)
                corretorMeses6(count) = meses(6)
                corretorMeses7(count) = meses(7)
                corretorMeses8(count) = meses(8)
                corretorMeses9(count) = meses(9)
                corretorMeses10(count) = meses(10)
                corretorMeses11(count) = meses(11)
                corretorMeses12(count) = meses(12)
                
                totalGeral = totalGeral + totalCorretor
                
                If variacao > 0 Then altas = altas + 1
                If variacao < 0 Then baixas = baixas + 1
                
                count = count + 1
            End If
            
            ' Prepara novo corretor
            lastCorretor = cCodigo
            corretorNome = cNome
            corretorGerencia = cGerencia
            For i = 1 To 12
                meses(i) = 0
            Next
        End If
        
        ' Acumula valor
        If cMes >= 1 And cMes <= 12 Then
            meses(cMes) = meses(cMes) + cValor
        End If
        
        rs.MoveNext
    Loop
    
    ' Processa último corretor
    If lastCorretor <> "" Then
        totalCorretor = 0
        For i = 1 To 12
            totalCorretor = totalCorretor + meses(i)
        Next
        
        lastVal1 = 0
        lastVal2 = 0
        lastMes1 = 0
        lastMes2 = 0
        
        For i = 12 To 1 Step -1
            If meses(i) > 0 Then
                If lastMes1 = 0 Then
                    lastMes1 = i
                    lastVal1 = meses(i)
                ElseIf lastMes2 = 0 Then
                    lastMes2 = i
                    lastVal2 = meses(i)
                    Exit For
                End If
            End If
        Next
        
        variacao = 0
        If lastVal2 > 0 Then
            variacao = ((lastVal1 - lastVal2) / lastVal2) * 100
        ElseIf lastVal1 > 0 Then
            variacao = 100
        ElseIf lastVal1 = 0 And lastVal2 > 0 Then
            variacao = -100
        End If
        
        ReDim Preserve corretorCodigos(count)
        ReDim Preserve corretorNomes(count)
        ReDim Preserve corretorGerencias(count)
        ReDim Preserve corretorTotais(count)
        ReDim Preserve corretorVariacoes(count)
        
        ReDim Preserve corretorMeses1(count)
        ReDim Preserve corretorMeses2(count)
        ReDim Preserve corretorMeses3(count)
        ReDim Preserve corretorMeses4(count)
        ReDim Preserve corretorMeses5(count)
        ReDim Preserve corretorMeses6(count)
        ReDim Preserve corretorMeses7(count)
        ReDim Preserve corretorMeses8(count)
        ReDim Preserve corretorMeses9(count)
        ReDim Preserve corretorMeses10(count)
        ReDim Preserve corretorMeses11(count)
        ReDim Preserve corretorMeses12(count)
        
        corretorCodigos(count) = lastCorretor
        corretorNomes(count) = corretorNome
        corretorGerencias(count) = corretorGerencia
        corretorTotais(count) = totalCorretor
        corretorVariacoes(count) = variacao
                
        corretorMeses1(count) = meses(1)
        corretorMeses2(count) = meses(2)
        corretorMeses3(count) = meses(3)
        corretorMeses4(count) = meses(4)
        corretorMeses5(count) = meses(5)
        corretorMeses6(count) = meses(6)
        corretorMeses7(count) = meses(7)
        corretorMeses8(count) = meses(8)
        corretorMeses9(count) = meses(9)
        corretorMeses10(count) = meses(10)
        corretorMeses11(count) = meses(11)
        corretorMeses12(count) = meses(12)
        
        totalGeral = totalGeral + totalCorretor
        
        If variacao > 0 Then altas = altas + 1
        If variacao < 0 Then baixas = baixas + 1
        
        count = count + 1
    End If
End If

rs.Close
Set rs = Nothing

' Busca lista de gerencias para o filtro
Dim rsGerencias
Set rsGerencias = connSales.Execute("SELECT DISTINCT Gerencia FROM Vendas WHERE Excluido = 0 AND Gerencia IS NOT NULL AND Gerencia <> '' ORDER BY Gerencia")

' Busca lista de corretores para o filtro - agora da tabela Vendas
Dim rsCorretores
Set rsCorretores = connSales.Execute("SELECT DISTINCT CorretorId, Corretor FROM Vendas WHERE Excluido = 0  ORDER BY Corretor")

%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>Painel Bolsa - Corretores</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<style>
    body { 
        background: #0a0e17; 
        color: #e0e0e0; 
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        padding-bottom: 20px;
    }
    
    .stock-header {
        background: linear-gradient(135deg, #1a1f2e 0%, #0a0e17 100%);
        border-bottom: 1px solid #2a3142;
        padding: 15px 0;
        margin-bottom: 20px;
    }
    
    .filter-panel {
        background: #1a1f2e;
        border-radius: 10px;
        border: 1px solid #2a3142;
        padding: 20px;
        margin-bottom: 25px;
    }
    
    .btn-stock {
        background: linear-gradient(135deg, #3a7bd5 0%, #00d2ff 100%);
        border: none;
        color: white;
        font-weight: 600;
        padding: 10px 25px;
        border-radius: 8px;
        transition: all 0.3s;
    }
    
    .btn-stock:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(58, 123, 213, 0.4);
        color: white;
    }
    
    .form-control-stock {
        background: #121826;
        border: 1px solid #2a3142;
        color: #e0e0e0;
        border-radius: 8px;
    }
    
    .market-status {
        background: #121826;
        border-radius: 20px;
        padding: 8px 15px;
        font-size: 0.9rem;
        border: 1px solid #2a3142;
    }
    
    .total-market {
        font-size: 1.8rem;
        font-weight: 700;
        color: #00d2ff;
        margin: 10px 0;
    }
    
    /* Estilos para a tabela */
    .stock-table {
        background: #1a1f2e;
        border-radius: 10px;
        border: 1px solid #2a3142;
        overflow: hidden;
        margin-bottom: 30px;
    }
    
    .table-header {
        background: #121826;
        padding: 15px 20px;
        border-bottom: 1px solid #2a3142;
    }
    
    .table-container {
        overflow-x: auto;
    }
    
    .table-dark-custom {
        background: #1a1f2e;
        color: #e0e0e0;
        margin-bottom: 0;
    }
    
    .table-dark-custom th {
        background: #121826;
        border-color: #2a3142;
        font-weight: 600;
        padding: 12px 15px;
        white-space: nowrap;
    }
    
    .table-dark-custom td {
        border-color: #2a3142;
        padding: 12px 15px;
        vertical-align: middle;
    }
    
    .table-dark-custom tbody tr:hover {
        background: rgba(58, 123, 213, 0.1);
    }
    
    .corretor-cell {
        font-weight: 600;
        color: #ffffff;
        min-width: 200px;
        position: sticky;
        left: 0;
        background: #1a1f2e;
        z-index: 1;
    }
    
    .gerencia-cell {
        min-width: 150px;
        color: #8a94a6;
    }
    
    .month-cell {
        text-align: center;
        min-width: 90px;
    }
    
    .valor-cell {
        font-weight: 600;
        font-size: 0.9rem;
    }
    
    .valor-up {
        color: #00c853;
        background: rgba(0, 200, 83, 0.1);
    }
    
    .valor-down {
        color: #ff3d00;
        background: rgba(255, 61, 0, 0.1);
    }
    
    .valor-neutral {
        color: #9e9e9e;
    }
    
    .trend-icon {
        margin-left: 3px;
        font-size: 0.75rem;
    }
    
    .total-cell {
        font-weight: 700;
        color: #00d2ff;
        background: rgba(0, 210, 255, 0.1);
        text-align: right;
        min-width: 140px;
    }
    
    .variacao-cell {
        text-align: center;
        min-width: 110px;
    }
    
    .change-up {
        color: #00c853;
        background: rgba(0, 200, 83, 0.15);
        padding: 3px 10px;
        border-radius: 20px;
        font-weight: 600;
        display: inline-block;
        font-size: 0.9rem;
    }
    
    .change-down {
        color: #ff3d00;
        background: rgba(255, 61, 0, 0.15);
        padding: 3px 10px;
        border-radius: 20px;
        font-weight: 600;
        display: inline-block;
        font-size: 0.9rem;
    }
    
    .change-neutral {
        color: #9e9e9e;
        background: rgba(158, 158, 158, 0.15);
        padding: 3px 10px;
        border-radius: 20px;
        font-weight: 600;
        display: inline-block;
        font-size: 0.9rem;
    }
    
    .empty-cell {
        color: #8a94a6;
        font-style: italic;
    }
    
    .table-total-row {
        background: #121826 !important;
        font-weight: 700;
        border-top: 2px solid #2a3142;
    }
    
    .table-total-row td {
        border-top: 2px solid #2a3142;
    }
    
    .simbolo-col {
        color: #3a7bd5;
        font-weight: 700;
        font-size: 1rem;
        width: 60px;
    }
    
    .participacao-bar {
        width: 80px;
        height: 6px;
        background: #121826;
        border-radius: 3px;
        overflow: hidden;
        display: inline-block;
        margin-right: 8px;
    }
    
    .participacao-fill {
        height: 100%;
        background: linear-gradient(90deg, #3a7bd5 0%, #00d2ff 100%);
        border-radius: 3px;
    }
    
    .participacao-cell {
        min-width: 100px;
        font-size: 0.9rem;
    }
    
    /* Cards de resumo */
    .resumo-card {
        background: #1a1f2e;
        border-radius: 10px;
        border: 1px solid #2a3142;
        padding: 15px;
        margin-bottom: 20px;
        height: 100%;
    }
    
    .resumo-title {
        font-size: 0.85rem;
        color: #8a94a6;
        margin-bottom: 8px;
    }
    
    .resumo-value {
        font-size: 1.4rem;
        font-weight: 700;
        color: #00d2ff;
        margin-bottom: 5px;
    }
    
    /* Estilos para o ranking */
    .ranking-badge {
        display: inline-block;
        width: 24px;
        height: 24px;
        background: #3a7bd5;
        color: white;
        border-radius: 50%;
        text-align: center;
        line-height: 24px;
        font-weight: 600;
        font-size: 0.8rem;
        margin-right: 8px;
    }
    
    .ranking-1 { background: linear-gradient(135deg, #ffd700 0%, #ffaa00 100%); color: #000; }
    .ranking-2 { background: linear-gradient(135deg, #c0c0c0 0%, #a0a0a0 100%); }
    .ranking-3 { background: linear-gradient(135deg, #cd7f32 0%, #a6692e 100%); }
    
    .ranking-cell {
        width: 60px;
        text-align: center;
    }
    
    /* Estilos para filtros */
    .mes-buttons {
        display: flex;
        flex-wrap: wrap;
        gap: 5px;
        margin-top: 5px;
    }
    
    .mes-btn {
        padding: 4px 10px;
        font-size: 0.8rem;
        border-radius: 4px;
        border: 1px solid #2a3142;
        background: #121826;
        color: #8a94a6;
        transition: all 0.2s;
    }
    
    .mes-btn:hover, .mes-btn.active {
        background: #3a7bd5;
        color: white;
        border-color: #3a7bd5;
    }
</style>
</head>
<body>

<!-- HEADER -->
<div class="stock-header">
    <div class="container">
        <div class="row align-items-center">
            <div class="col-md-6">
                <h1 class="mb-0">
                    <i class="fas fa-user-tie text-primary me-2"></i>
                    PAINEL BOLSA <span class="text-info">CORRETORES</span>
                </h1>
                <p class="text-muted mb-0">Ranking e desempenho mensal dos corretores</p>
            </div>
            <div class="col-md-6 text-end">
                <div class="market-status d-inline-block">
                    <span class="me-2">Mercado:</span>
                    <span class="text-success">
                        <i class="fas fa-arrow-up me-1"></i>ABERTO
                    </span>
                    <span class="ms-3">
                        <i class="far fa-calendar me-1"></i><%= anoFiltro %>
                    </span>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- FILTROS -->
<div class="container">
    <div class="filter-panel">
        <form method="get" class="row g-3 align-items-end">
            <div class="col-md-2">
                <label class="form-label">Ano</label>
                <select name="ano" class="form-select form-control-stock" onchange="this.form.submit()">
                    <% 
                    Dim rsAnos
                    Set rsAnos = connSales.Execute("SELECT DISTINCT AnoVenda FROM Vendas WHERE Excluido = 0 ORDER BY AnoVenda DESC")
                    
                    Do While Not rsAnos.EOF
                        If Not IsNull(rsAnos("AnoVenda")) Then
                            Dim anoVal
                            anoVal = CStr(rsAnos("AnoVenda"))
                            Response.Write "<option value=""" & anoVal & """"
                            If CStr(anoFiltro) = anoVal Then Response.Write " selected"
                            Response.Write ">" & anoVal & "</option>"
                        End If
                        rsAnos.MoveNext
                    Loop
                    
                    rsAnos.Close
                    Set rsAnos = Nothing
                    %>
                </select>
            </div>
            
            <div class="col-md-3">
                <label class="form-label">Mês</label>
                <select name="mes" class="form-select form-control-stock" onchange="this.form.submit()">
                    <option value="">Todos os meses</option>
                    <option value="1" <% If mesFiltro = "1" Then Response.Write "selected" %>>Janeiro</option>
                    <option value="2" <% If mesFiltro = "2" Then Response.Write "selected" %>>Fevereiro</option>
                    <option value="3" <% If mesFiltro = "3" Then Response.Write "selected" %>>Março</option>
                    <option value="4" <% If mesFiltro = "4" Then Response.Write "selected" %>>Abril</option>
                    <option value="5" <% If mesFiltro = "5" Then Response.Write "selected" %>>Maio</option>
                    <option value="6" <% If mesFiltro = "6" Then Response.Write "selected" %>>Junho</option>
                    <option value="7" <% If mesFiltro = "7" Then Response.Write "selected" %>>Julho</option>
                    <option value="8" <% If mesFiltro = "8" Then Response.Write "selected" %>>Agosto</option>
                    <option value="9" <% If mesFiltro = "9" Then Response.Write "selected" %>>Setembro</option>
                    <option value="10" <% If mesFiltro = "10" Then Response.Write "selected" %>>Outubro</option>
                    <option value="11" <% If mesFiltro = "11" Then Response.Write "selected" %>>Novembro</option>
                    <option value="12" <% If mesFiltro = "12" Then Response.Write "selected" %>>Dezembro</option>
                </select>
            </div>
            
            <div class="col-md-3">
                <label class="form-label">Gerência</label>
                <select name="gerencia" class="form-select form-control-stock" onchange="this.form.submit()">
                    <option value="">Todas as Gerências</option>
                    <%
                    Do While Not rsGerencias.EOF
                        If Not IsNull(rsGerencias("Gerencia")) Then
                            Dim gerenciaVal
                            gerenciaVal = Trim(rsGerencias("Gerencia"))
                            Response.Write "<option value=""" & gerenciaVal & """"
                            If CStr(gerenciaFiltro) = CStr(gerenciaVal) Then Response.Write " selected"
                            Response.Write ">" & gerenciaVal & "</option>"
                        End If
                        rsGerencias.MoveNext
                    Loop
                    rsGerencias.Close
                    Set rsGerencias = Nothing
                    %>
                </select>
            </div>
            
            <div class="col-md-3">
                <label class="form-label">Corretor</label>
                <select name="corretor" class="form-select form-control-stock" onchange="this.form.submit()">
                    <option value="">Todos os Corretores</option>
                    <%
                    Do While Not rsCorretores.EOF
                        If Not IsNull(rsCorretores("CorretorId")) Then
                            Dim corretorCod, corretorNomeFiltro
                            corretorCod = Trim(rsCorretores("CorretorId"))
                            If Not IsNull(rsCorretores("Corretor")) Then
                                corretorNomeFiltro = Trim(rsCorretores("Corretor"))
                            Else
                                corretorNomeFiltro = corretorCod
                            End If
                            Response.Write "<option value=""" & corretorCod & """"
                            If CStr(corretorFiltro) = CStr(corretorCod) Then Response.Write " selected"
                            Response.Write ">" & corretorNomeFiltro & " (" & corretorCod & ")" %></option>
                        <%
                        End If
                        rsCorretores.MoveNext
                    Loop
                    rsCorretores.Close
                    Set rsCorretores = Nothing
                    %>
                </select>
            </div>
            
            <div class="col-md-1 text-end">
                <button type="submit" class="btn btn-stock">
                    <i class="fas fa-sync-alt"></i>
                </button>
                <a href="?ano=<%= anoFiltro %>" class="btn btn-outline-light ms-2">
                    <i class="fas fa-times"></i>
                </a>
            </div>
        </form>
        
        <!-- Botões rápidos de meses -->
<!-- Botões rápidos de meses -->
<div class="row mt-3">
    <div class="col-12">
        <label class="form-label mb-2">Meses rápidos:</label>
        <div class="mes-buttons">
            <%
            Dim mesAtual
            mesAtual = Month(Date())
            Dim mesesRapidos
            mesesRapidos = Array("JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ")
            
            For m = 1 To 12
                Dim mesClass
                mesClass = "mes-btn"
                If CStr(m) = mesFiltro Then mesClass = mesClass & " active"
            %>
            <a href="?ano=<%= anoFiltro %>&mes=<%= m %><% If gerenciaFiltro <> "" Then Response.Write "&gerencia=" & Server.URLEncode(gerenciaFiltro) %><% If corretorFiltro <> "" Then Response.Write "&corretor=" & Server.URLEncode(corretorFiltro) %>" 
               class="<%= mesClass %>">
                <%= mesesRapidos(m-1) %>
            </a>
            <% Next %>
            <%
            ' Corrigir o botão "TODOS" sem usar IIf
            Dim btnTodosClass
            btnTodosClass = "mes-btn"
            If mesFiltro = "" Then
                btnTodosClass = btnTodosClass & " active"
            End If
            %>
            <a href="?ano=<%= anoFiltro %><% If gerenciaFiltro <> "" Then Response.Write "&gerencia=" & Server.URLEncode(gerenciaFiltro) %><% If corretorFiltro <> "" Then Response.Write "&corretor=" & Server.URLEncode(corretorFiltro) %>" 
               class="<%= btnTodosClass %>">
                TODOS
            </a>
        </div>
    </div>
</div>
        
    </div>
</div>

<!-- RESUMO DO MERCADO -->
<div class="container mb-4">
    <div class="row">
        <div class="col-md-3">
            <div class="resumo-card text-center">
                <div class="resumo-title">Total Geral</div>
                <div class="resumo-value">
                    R$ 
                    <%
                    If totalGeral > 0 Then
                        Response.Write FormatNumber(totalGeral, 2)
                    Else
                        Response.Write "0,00"
                    End If
                    %>
                </div>
                <div class="text-muted"><%= anoFiltro %></div>
            </div>
        </div>
        
        <div class="col-md-2">
            <div class="resumo-card text-center">
                <div class="resumo-title">Corretores</div>
                <div class="resumo-value">
                    <%= count %>
                </div>
                <div class="text-muted">Ativos</div>
            </div>
        </div>
        
        <div class="col-md-2">
            <div class="resumo-card text-center">
                <div class="resumo-title">Em Alta</div>
                <div class="resumo-value text-success">
                    <%= altas %>
                </div>
                <div class="text-success small">
                    <i class="fas fa-arrow-up me-1"></i>Crescimento
                </div>
            </div>
        </div>
        
        <div class="col-md-2">
            <div class="resumo-card text-center">
                <div class="resumo-title">Em Baixa</div>
                <div class="resumo-value text-danger">
                    <%= baixas %>
                </div>
                <div class="text-danger small">
                    <i class="fas fa-arrow-down me-1"></i>Queda
                </div>
            </div>
        </div>
        
        <div class="col-md-3">
            <div class="resumo-card text-center">
                <div class="resumo-title">Média por Corretor</div>
                <div class="resumo-value">
                    R$ 
                    <%
                    If count > 0 And totalGeral > 0 Then
                        Response.Write FormatNumber(totalGeral / count, 2)
                    Else
                        Response.Write "0,00"
                    End If
                    %>
                </div>
                <div class="text-muted">Média anual</div>
            </div>
        </div>
    </div>
</div>

<!-- TABELA PRINCIPAL -->
<div class="container">
    <div class="stock-table">
        <div class="table-header">
            <div class="d-flex justify-content-between align-items-center">
                <h5 class="mb-0">
                    <i class="fas fa-chart-line me-2"></i>
                    Ranking de Corretores - <%= anoFiltro %>
                </h5>
                <div class="text-muted small">
                    <i class="fas fa-info-circle me-1"></i>
                    <span>Verde ↑ = Crescimento | Vermelho ↓ = Queda</span>
                </div>
            </div>
        </div>
        
        <div class="table-container">
            <table class="table table-dark-custom">
                <thead>
                    <tr>
                        <th class="ranking-cell">#</th>
                        <th class="corretor-cell">Corretor</th>
                        <th class="gerencia-cell">Gerência</th>
                        <th class="month-cell">JAN</th>
                        <th class="month-cell">FEV</th>
                        <th class="month-cell">MAR</th>
                        <th class="month-cell">ABR</th>
                        <th class="month-cell">MAI</th>
                        <th class="month-cell">JUN</th>
                        <th class="month-cell">JUL</th>
                        <th class="month-cell">AGO</th>
                        <th class="month-cell">SET</th>
                        <th class="month-cell">OUT</th>
                        <th class="month-cell">NOV</th>
                        <th class="month-cell">DEZ</th>
                        <th class="total-cell">TOTAL</th>
                        <th class="variacao-cell">VARIAÇÃO</th>
                        <th class="participacao-cell">PARTICIPAÇÃO</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    If count > 0 Then
                        ' Declarar variáveis temporárias
                        Dim tempCodigo, tempNome, tempGerencia, tempTotal, tempVar
                        Dim tempM1, tempM2, tempM3, tempM4, tempM5, tempM6
                        Dim tempM7, tempM8, tempM9, tempM10, tempM11, tempM12
                        Dim i, j, m
                        
                        ' Ordena por total usando GetNum para garantir valores numéricos
                        For i = 0 To count - 2
                            For j = i + 1 To count - 1
                                If GetNum(corretorTotais(i)) < GetNum(corretorTotais(j)) Then
                                    ' Troca códigos
                                    tempCodigo = corretorCodigos(i)
                                    corretorCodigos(i) = corretorCodigos(j)
                                    corretorCodigos(j) = tempCodigo
                                    
                                    ' Troca nomes
                                    tempNome = corretorNomes(i)
                                    corretorNomes(i) = corretorNomes(j)
                                    corretorNomes(j) = tempNome
                                    
                                    ' Troca gerencias
                                    tempGerencia = corretorGerencias(i)
                                    corretorGerencias(i) = corretorGerencias(j)
                                    corretorGerencias(j) = tempGerencia
                                    
                                    ' Troca totais
                                    tempTotal = corretorTotais(i)
                                    corretorTotais(i) = corretorTotais(j)
                                    corretorTotais(j) = tempTotal
                                    
                                    ' Troca variações
                                    tempVar = corretorVariacoes(i)
                                    corretorVariacoes(i) = corretorVariacoes(j)
                                    corretorVariacoes(j) = tempVar
                                    
                                    ' Troca meses
                                    tempM1 = corretorMeses1(i)
                                    corretorMeses1(i) = corretorMeses1(j)
                                    corretorMeses1(j) = tempM1
                                    
                                    tempM2 = corretorMeses2(i)
                                    corretorMeses2(i) = corretorMeses2(j)
                                    corretorMeses2(j) = tempM2
                                    
                                    tempM3 = corretorMeses3(i)
                                    corretorMeses3(i) = corretorMeses3(j)
                                    corretorMeses3(j) = tempM3
                                    
                                    tempM4 = corretorMeses4(i)
                                    corretorMeses4(i) = corretorMeses4(j)
                                    corretorMeses4(j) = tempM4
                                    
                                    tempM5 = corretorMeses5(i)
                                    corretorMeses5(i) = corretorMeses5(j)
                                    corretorMeses5(j) = tempM5
                                    
                                    tempM6 = corretorMeses6(i)
                                    corretorMeses6(i) = corretorMeses6(j)
                                    corretorMeses6(j) = tempM6
                                    
                                    tempM7 = corretorMeses7(i)
                                    corretorMeses7(i) = corretorMeses7(j)
                                    corretorMeses7(j) = tempM7
                                    
                                    tempM8 = corretorMeses8(i)
                                    corretorMeses8(i) = corretorMeses8(j)
                                    corretorMeses8(j) = tempM8
                                    
                                    tempM9 = corretorMeses9(i)
                                    corretorMeses9(i) = corretorMeses9(j)
                                    corretorMeses9(j) = tempM9
                                    
                                    tempM10 = corretorMeses10(i)
                                    corretorMeses10(i) = corretorMeses10(j)
                                    corretorMeses10(j) = tempM10
                                    
                                    tempM11 = corretorMeses11(i)
                                    corretorMeses11(i) = corretorMeses11(j)
                                    corretorMeses11(j) = tempM11
                                    
                                    tempM12 = corretorMeses12(i)
                                    corretorMeses12(i) = corretorMeses12(j)
                                    corretorMeses12(j) = tempM12
                                End If
                            Next
                        Next
                        
                        ' Exibe cada corretor
                        For i = 0 To count - 1
                            'Dim cCodigo, cNome, cGerencia, cTotal, cVariacao
                            Dim changeClass, changeIcon, simbolo, palavras, palavra
                            Dim participacao, totalFormatado, variacaoFormatada, participacaoFormatada
                            Dim rankingClass
                            
                            cCodigo = corretorCodigos(i)
                            cNome = corretorNomes(i)
                            cGerencia = corretorGerencias(i)
                            cTotal = corretorTotais(i)
                            cVariacao = corretorVariacoes(i)
                            
                            ' Determina classe do ranking
                            If i = 0 Then
                                rankingClass = "ranking-1"
                            ElseIf i = 1 Then
                                rankingClass = "ranking-2"
                            ElseIf i = 2 Then
                                rankingClass = "ranking-3"
                            Else
                                rankingClass = ""
                            End If
                            
                            ' Determina estilo da variação
                            If GetNum(cVariacao) > 0 Then
                                changeClass = "change-up"
                                changeIcon = "fas fa-arrow-up"
                            ElseIf GetNum(cVariacao) < 0 Then
                                changeClass = "change-down"
                                changeIcon = "fas fa-arrow-down"
                            Else
                                changeClass = "change-neutral"
                                changeIcon = "fas fa-minus"
                            End If
                            
                            ' Gera símbolo do corretor (iniciais)
                            simbolo = ""
                            palavras = Split(cNome, " ")
                            Dim palavraCount
                            palavraCount = 0
                            For Each palavra In palavras
                                If Len(palavra) > 0 Then
                                    simbolo = simbolo & UCase(Left(palavra, 1))
                                    palavraCount = palavraCount + 1
                                    If palavraCount >= 2 Then Exit For
                                End If
                            Next
                            If Len(simbolo) = 0 Then
                                simbolo = Left(cCodigo, 2)
                            End If
                            
                            ' Participação
                            If GetNum(totalGeral) > 0 Then
                                participacao = (GetNum(cTotal) / GetNum(totalGeral)) * 100
                            Else
                                participacao = 0
                            End If
                            
                            ' Formata números
                            totalFormatado = FormatNumber(GetNum(cTotal), 2)
                            variacaoFormatada = FormatNumber(Abs(GetNum(cVariacao)), 2)
                            participacaoFormatada = FormatNumber(GetNum(participacao), 2)
                            
                            ' Declara e preenche array de valores mensais
                            Dim mesesValores
                            ReDim mesesValores(12)
                            
                            mesesValores(1) = GetNum(corretorMeses1(i))
                            mesesValores(2) = GetNum(corretorMeses2(i))
                            mesesValores(3) = GetNum(corretorMeses3(i))
                            mesesValores(4) = GetNum(corretorMeses4(i))
                            mesesValores(5) = GetNum(corretorMeses5(i))
                            mesesValores(6) = GetNum(corretorMeses6(i))
                            mesesValores(7) = GetNum(corretorMeses7(i))
                            mesesValores(8) = GetNum(corretorMeses8(i))
                            mesesValores(9) = GetNum(corretorMeses9(i))
                            mesesValores(10) = GetNum(corretorMeses10(i))
                            mesesValores(11) = GetNum(corretorMeses11(i))
                            mesesValores(12) = GetNum(corretorMeses12(i))
                    %>
                    <tr>
                        <td class="ranking-cell">
                            <% If rankingClass <> "" Then %>
                            <span class="ranking-badge <%= rankingClass %>"><%= i+1 %></span>
                            <% Else %>
                            <span class="text-muted"><%= i+1 %></span>
                            <% End If %>
                        </td>
                        
                        <td class="corretor-cell">
                            <div class="d-flex align-items-center">
                                <div class="simbolo-col <%= rankingClass %> d-inline-block me-2">
                                    <%= simbolo %>
                                </div>
                                <div>
                                    <div class="fw-bold"><%= Left(cNome, 20) %></div>
                                    <div class="small text-muted"><%= cCodigo %></div>
                                </div>
                            </div>
                        </td>
                        
                        <td class="gerencia-cell">
                            <%= Left(cGerencia, 15) %>
                        </td>
                        
                        <%
                        ' Exibe valores mensais com cores e setas
                        For m = 1 To 12
                            Dim valorMes, valorAnterior, valorClass, valorExibir
                            Dim trendIcon
                            
                            valorMes = mesesValores(m)
                            
                            ' Calcula variação em relação ao mês anterior
                            valorAnterior = 0
                            If m > 1 Then
                                valorAnterior = mesesValores(m-1)
                            End If
                            
                            ' Determina classe de cor e ícone
                            If valorMes = 0 Then
                                valorClass = "valor-neutral"
                                trendIcon = ""
                            Else
                                If m = 1 Then
                                    ' Primeiro mês: apenas valor positivo
                                    valorClass = "valor-up"
                                    trendIcon = ""
                                Else
                                    ' Compara com mês anterior
                                    If valorAnterior = 0 Then
                                        ' Cresceu de zero para algum valor
                                        valorClass = "valor-up"
                                        trendIcon = "<i class='fas fa-arrow-up trend-icon'></i>"
                                    ElseIf valorMes > valorAnterior Then
                                        ' Cresceu em relação ao anterior
                                        valorClass = "valor-up"
                                        trendIcon = "<i class='fas fa-arrow-up trend-icon'></i>"
                                    ElseIf valorMes < valorAnterior Then
                                        ' Caiu em relação ao anterior
                                        valorClass = "valor-down"
                                        trendIcon = "<i class='fas fa-arrow-down trend-icon'></i>"
                                    Else
                                        ' Igual ao anterior
                                        valorClass = "valor-neutral"
                                        trendIcon = "<i class='fas fa-minus trend-icon'></i>"
                                    End If
                                End If
                            End If
                            
                            ' Formata valor para exibição
                            If valorMes > 0 Then
                                If valorMes >= 1000000 Then
                                    valorExibir = FormatNumber(valorMes / 1000000, 1) & "M"
                                ElseIf valorMes >= 1000 Then
                                    valorExibir = FormatNumber(valorMes / 1000, 0) & "k"
                                Else
                                    valorExibir = FormatNumber(valorMes, 0)
                                End If
                            Else
                                valorExibir = "-"
                                valorClass = "empty-cell"
                            End If
                        %>
                        <td class="month-cell valor-cell <%= valorClass %>">
                            <%= valorExibir %>
                            <% If valorMes > 0 And trendIcon <> "" Then %>
                            <%= trendIcon %>
                            <% End If %>
                        </td>
                        <% Next %>
                        
                        <td class="total-cell">
                            R$ <%= totalFormatado %>
                        </td>
                        
                        <td class="variacao-cell">
                            <span class="<%= changeClass %>">
                                <i class="<%= changeIcon %> me-1"></i>
                                <%= variacaoFormatada %>%
                            </span>
                        </td>
                        
                        <td class="participacao-cell">
                            <div class="d-flex align-items-center">
                                <div class="participacao-bar">
                                    <div class="participacao-fill" style="width: <%= participacaoFormatada %>%"></div>
                                </div>
                                <div>
                                    <%= participacaoFormatada %>%
                                </div>
                            </div>
                        </td>
                    </tr>
                    <% 
                        Next 
                    %>
                    
                    <!-- LINHA DE TOTAIS -->
                    <tr class="table-total-row">
                        <td colspan="3" class="corretor-cell">
                            <strong>TOTAL GERAL</strong>
                        </td>
                        <%
                        ' Calcula totais mensais
                        Dim totalMensal(12)
                        For m = 1 To 12
                            totalMensal(m) = 0
                        Next
                        
                        If count > 0 Then
                            For i = 0 To count - 1
                                totalMensal(1) = totalMensal(1) + GetNum(corretorMeses1(i))
                                totalMensal(2) = totalMensal(2) + GetNum(corretorMeses2(i))
                                totalMensal(3) = totalMensal(3) + GetNum(corretorMeses3(i))
                                totalMensal(4) = totalMensal(4) + GetNum(corretorMeses4(i))
                                totalMensal(5) = totalMensal(5) + GetNum(corretorMeses5(i))
                                totalMensal(6) = totalMensal(6) + GetNum(corretorMeses6(i))
                                totalMensal(7) = totalMensal(7) + GetNum(corretorMeses7(i))
                                totalMensal(8) = totalMensal(8) + GetNum(corretorMeses8(i))
                                totalMensal(9) = totalMensal(9) + GetNum(corretorMeses9(i))
                                totalMensal(10) = totalMensal(10) + GetNum(corretorMeses10(i))
                                totalMensal(11) = totalMensal(11) + GetNum(corretorMeses11(i))
                                totalMensal(12) = totalMensal(12) + GetNum(corretorMeses12(i))
                            Next
                        End If
                        
                        For m = 1 To 12
                            Dim totalMesFormatado, totalMesClass, totalMesTrendIcon
                            Dim totalMesAnterior
                            
                            ' Determina tendência do total mensal
                            totalMesAnterior = 0
                            If m > 1 Then
                                totalMesAnterior = totalMensal(m-1)
                            End If
                            
                            If totalMensal(m) > 0 Then
                                If totalMensal(m) >= 1000000 Then
                                    totalMesFormatado = FormatNumber(totalMensal(m) / 1000000, 1) & "M"
                                ElseIf totalMensal(m) >= 1000 Then
                                    totalMesFormatado = FormatNumber(totalMensal(m) / 1000, 0) & "k"
                                Else
                                    totalMesFormatado = FormatNumber(totalMensal(m), 0)
                                End If
                                
                                If m = 1 Then
                                    totalMesClass = "valor-up"
                                    totalMesTrendIcon = ""
                                Else
                                    If totalMesAnterior = 0 Then
                                        totalMesClass = "valor-up"
                                        totalMesTrendIcon = "<i class='fas fa-arrow-up trend-icon'></i>"
                                    ElseIf totalMensal(m) > totalMesAnterior Then
                                        totalMesClass = "valor-up"
                                        totalMesTrendIcon = "<i class='fas fa-arrow-up trend-icon'></i>"
                                    ElseIf totalMensal(m) < totalMesAnterior Then
                                        totalMesClass = "valor-down"
                                        totalMesTrendIcon = "<i class='fas fa-arrow-down trend-icon'></i>"
                                    Else
                                        totalMesClass = "valor-neutral"
                                        totalMesTrendIcon = "<i class='fas fa-minus trend-icon'></i>"
                                    End If
                                End If
                            Else
                                totalMesFormatado = "-"
                                totalMesClass = "empty-cell"
                                totalMesTrendIcon = ""
                            End If
                        %>
                        <td class="month-cell valor-cell <%= totalMesClass %>">
                            <strong><%= totalMesFormatado %></strong>
                            <% If totalMensal(m) > 0 And totalMesTrendIcon <> "" Then %>
                            <%= totalMesTrendIcon %>
                            <% End If %>
                        </td>
                        <% Next %>
                        
                        <td class="total-cell">
                            <strong>R$ <%= FormatNumber(totalGeral, 2) %></strong>
                        </td>
                        
                        <td class="variacao-cell">
                            <span class="text-muted">-</span>
                        </td>
                        
                        <td class="participacao-cell">
                            <div class="d-flex align-items-center">
                                <div class="participacao-bar">
                                    <div class="participacao-fill" style="width: 100%"></div>
                                </div>
                                <div>
                                    100%
                                </div>
                            </div>
                        </td>
                    </tr>
                    <% Else %>
                    <tr>
                        <td colspan="17" class="text-center py-5">
                            <i class="fas fa-user-tie fa-3x text-muted mb-3"></i>
                            <h3 class="text-muted">Nenhum corretor encontrado</h3>
                            <p class="text-muted">Tente ajustar os filtros ou verificar os dados do ano <%= anoFiltro %></p>
                        </td>
                    </tr>
                    <% End If %>
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- RODAPÉ -->
<div class="container mt-4">
    <div class="text-center text-muted">
        <p>
            <i class="fas fa-info-circle me-1"></i>
            Painel Bolsa de Corretores - Dados referentes ao ano <%= anoFiltro %> 
            | Atualizado em <%= Now() %> 
            | Total de corretores: <%= count %>
        </p>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>

</body>
</html>

<%
' Fecha conexões
If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If
If Not connSales Is Nothing Then
    If connSales.State = 1 Then connSales.Close
    Set connSales = Nothing
End If
%>