<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: BBKVHDODGV          -->
<!-- OBS: Painel Bolsa de Valores - VGVs Gerências -->
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
Dim anoFiltro, gerenciaFiltro
anoFiltro = Request.QueryString("ano")
gerenciaFiltro = Request.QueryString("gerencia")

If anoFiltro = "" Or Not IsNumeric(anoFiltro) Then 
    anoFiltro = Year(Date())
Else
    anoFiltro = CInt(anoFiltro)
End If

' ===============================================
' BUSCA DADOS BÁSICOS
' ===============================================
Dim sql, rs
sql = "SELECT Gerencia, MesVenda, ValorUnidade FROM Vendas "
sql = sql & "WHERE Excluido = 0 AND AnoVenda = " & anoFiltro
sql = sql & " AND Gerencia IS NOT NULL AND Gerencia <> ''"
sql = sql & " AND ValorUnidade IS NOT NULL"

If gerenciaFiltro <> "" Then
    sql = sql & " AND Gerencia = '" & Replace(gerenciaFiltro, "'", "''") & "'"
End If

sql = sql & " ORDER BY Gerencia, MesVenda"

Set rs = connSales.Execute(sql)

' ===============================================
' ARRAYS SIMPLES
' ===============================================
Dim gerenciaNomes(), gerenciaTotais(), gerenciaVariacoes()
Dim gerenciaMeses1(), gerenciaMeses2(), gerenciaMeses3()
Dim gerenciaMeses4(), gerenciaMeses5(), gerenciaMeses6()
Dim gerenciaMeses7(), gerenciaMeses8(), gerenciaMeses9()
Dim gerenciaMeses10(), gerenciaMeses11(), gerenciaMeses12()

Dim count, totalGeral, altas, baixas
count = 0
totalGeral = 0
altas = 0
baixas = 0

Dim currentGerencia, lastGerencia
Dim meses(12), totalGerencia
Dim lastVal1, lastVal2, lastMes1, lastMes2

If Not rs.EOF Then
    lastGerencia = ""
    
    Do While Not rs.EOF
        Dim gNome, gMes, gValor
        gNome = Trim(rs("Gerencia"))
        gMes = GetNum(rs("MesVenda"))
        gValor = GetNum(rs("ValorUnidade"))
        
        ' Nova gerência?
        If lastGerencia <> gNome Then
            ' Processa gerência anterior se existir
            If lastGerencia <> "" Then
                ' Calcula total
                totalGerencia = 0
                For i = 1 To 12
                    totalGerencia = totalGerencia + meses(i)
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
                ReDim Preserve gerenciaNomes(count)
                ReDim Preserve gerenciaTotais(count)
                ReDim Preserve gerenciaVariacoes(count)
                
                ReDim Preserve gerenciaMeses1(count)
                ReDim Preserve gerenciaMeses2(count)
                ReDim Preserve gerenciaMeses3(count)
                ReDim Preserve gerenciaMeses4(count)
                ReDim Preserve gerenciaMeses5(count)
                ReDim Preserve gerenciaMeses6(count)
                ReDim Preserve gerenciaMeses7(count)
                ReDim Preserve gerenciaMeses8(count)
                ReDim Preserve gerenciaMeses9(count)
                ReDim Preserve gerenciaMeses10(count)
                ReDim Preserve gerenciaMeses11(count)
                ReDim Preserve gerenciaMeses12(count)
                
                gerenciaNomes(count) = lastGerencia
                gerenciaTotais(count) = totalGerencia
                gerenciaVariacoes(count) = variacao
                
                gerenciaMeses1(count) = meses(1)
                gerenciaMeses2(count) = meses(2)
                gerenciaMeses3(count) = meses(3)
                gerenciaMeses4(count) = meses(4)
                gerenciaMeses5(count) = meses(5)
                gerenciaMeses6(count) = meses(6)
                gerenciaMeses7(count) = meses(7)
                gerenciaMeses8(count) = meses(8)
                gerenciaMeses9(count) = meses(9)
                gerenciaMeses10(count) = meses(10)
                gerenciaMeses11(count) = meses(11)
                gerenciaMeses12(count) = meses(12)
                
                totalGeral = totalGeral + totalGerencia
                
                If variacao > 0 Then altas = altas + 1
                If variacao < 0 Then baixas = baixas + 1
                
                count = count + 1
            End If
            
            ' Prepara nova gerência
            lastGerencia = gNome
            For i = 1 To 12
                meses(i) = 0
            Next
        End If
        
        ' Acumula valor
        If gMes >= 1 And gMes <= 12 Then
            meses(gMes) = meses(gMes) + gValor
        End If
        
        rs.MoveNext
    Loop
    
    ' Processa última gerência
    If lastGerencia <> "" Then
        totalGerencia = 0
        For i = 1 To 12
            totalGerencia = totalGerencia + meses(i)
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
        
        ReDim Preserve gerenciaNomes(count)
        ReDim Preserve gerenciaTotais(count)
        ReDim Preserve gerenciaVariacoes(count)
        
        ReDim Preserve gerenciaMeses1(count)
        ReDim Preserve gerenciaMeses2(count)
        ReDim Preserve gerenciaMeses3(count)
        ReDim Preserve gerenciaMeses4(count)
        ReDim Preserve gerenciaMeses5(count)
        ReDim Preserve gerenciaMeses6(count)
        ReDim Preserve gerenciaMeses7(count)
        ReDim Preserve gerenciaMeses8(count)
        ReDim Preserve gerenciaMeses9(count)
        ReDim Preserve gerenciaMeses10(count)
        ReDim Preserve gerenciaMeses11(count)
        ReDim Preserve gerenciaMeses12(count)
        
        gerenciaNomes(count) = lastGerencia
        gerenciaTotais(count) = totalGerencia
        gerenciaVariacoes(count) = variacao
                
        gerenciaMeses1(count) = meses(1)
        gerenciaMeses2(count) = meses(2)
        gerenciaMeses3(count) = meses(3)
        gerenciaMeses4(count) = meses(4)
        gerenciaMeses5(count) = meses(5)
        gerenciaMeses6(count) = meses(6)
        gerenciaMeses7(count) = meses(7)
        gerenciaMeses8(count) = meses(8)
        gerenciaMeses9(count) = meses(9)
        gerenciaMeses10(count) = meses(10)
        gerenciaMeses11(count) = meses(11)
        gerenciaMeses12(count) = meses(12)
        
        totalGeral = totalGeral + totalGerencia
        
        If variacao > 0 Then altas = altas + 1
        If variacao < 0 Then baixas = baixas + 1
        
        count = count + 1
    End If
End If

rs.Close
Set rs = Nothing
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>Painel Bolsa - VGVs Gerências</title>
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
    
    .stock-card {
        background: #1a1f2e;
        border-radius: 10px;
        border: 1px solid #2a3142;
        padding: 20px;
        margin-bottom: 20px;
        transition: transform 0.3s, box-shadow 0.3s;
        height: 100%;
    }
    
    .stock-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 10px 25px rgba(0, 0, 0, 0.3);
        border-color: #3a7bd5;
    }
    
    .stock-symbol {
        font-size: 1.8rem;
        font-weight: 700;
        color: #ffffff;
        margin-bottom: 5px;
        letter-spacing: 1px;
    }
    
    .stock-name {
        font-size: 0.9rem;
        color: #8a94a6;
        margin-bottom: 15px;
    }
    
    .stock-price {
        font-size: 2.2rem;
        font-weight: 700;
        margin: 15px 0;
    }
    
    .stock-change {
        font-size: 1.2rem;
        font-weight: 600;
        padding: 5px 12px;
        border-radius: 20px;
        display: inline-block;
    }
    
    .change-up {
        background: rgba(0, 200, 83, 0.15);
        color: #00c853;
    }
    
    .change-down {
        background: rgba(255, 61, 0, 0.15);
        color: #ff3d00;
    }
    
    .change-neutral {
        background: rgba(158, 158, 158, 0.15);
        color: #9e9e9e;
    }
    
    .month-grid {
        display: grid;
        grid-template-columns: repeat(6, 1fr);
        gap: 8px;
        margin-top: 20px;
    }
    
    .month-cell {
        background: #121826;
        border-radius: 6px;
        padding: 8px 5px;
        text-align: center;
        border: 1px solid #2a3142;
        transition: all 0.2s;
    }
    
    .month-cell:hover {
        background: #2a3142;
        border-color: #3a7bd5;
    }
    
    .month-label {
        font-size: 0.7rem;
        color: #8a94a6;
        margin-bottom: 5px;
    }
    
    .month-value {
        font-size: 0.9rem;
        font-weight: 600;
        color: #ffffff;
    }
    
    .value-positive {
        color: #00c853;
    }
    
    .value-zero {
        color: #9e9e9e;
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
    
    .volume-bar {
        width: 100%;
        height: 8px;
        background: #121826;
        border-radius: 4px;
        margin: 15px 0 5px 0;
        overflow: hidden;
    }
    
    .volume-fill {
        height: 100%;
        background: linear-gradient(90deg, #3a7bd5 0%, #00d2ff 100%);
        border-radius: 4px;
        transition: width 0.5s;
    }
    
    /* Estilos para as setas mensais */
    .month-trend {
        display: inline-block;
        margin-left: 3px;
        font-size: 0.7rem;
    }
    
    .trend-up {
        color: #00c853;
    }
    
    .trend-down {
        color: #ff3d00;
    }
    
    .trend-neutral {
        color: #9e9e9e;
    }
    
    .month-value-container {
        display: flex;
        align-items: center;
        justify-content: center;
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
                    <i class="fas fa-chart-line text-primary me-2"></i>
                    PAINEL BOLSA <span class="text-info">VGVs</span>
                </h1>
                <p class="text-muted mb-0">Monitoramento em tempo real das gerências</p>
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
            <div class="col-md-3">
                <label class="form-label">Ano</label>
                <select name="ano" class="form-select" style="background:#121826;border:1px solid #2a3142;color:#e0e0e0;" onchange="this.form.submit()">
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
            
            <div class="col-md-5">
                <label class="form-label">Gerência</label>
                <select name="gerencia" class="form-select" style="background:#121826;border:1px solid #2a3142;color:#e0e0e0;" onchange="this.form.submit()">
                    <option value="">Todas as Gerências</option>
                    <%
                    If count > 0 Then
                        For i = 0 To count - 1
                            Response.Write "<option value=""" & gerenciaNomes(i) & """"
                            If CStr(gerenciaFiltro) = CStr(gerenciaNomes(i)) Then Response.Write " selected"
                            Response.Write ">" & gerenciaNomes(i) & "</option>"
                        Next
                    End If
                    %>
                </select>
            </div>
            
            <div class="col-md-4 text-end">
                <button type="submit" class="btn btn-stock me-2">
                    <i class="fas fa-sync-alt me-2"></i>ATUALIZAR
                </button>
                <a href="?ano=<%= anoFiltro %>" class="btn btn-outline-light">LIMPAR FILTRO</a>
            </div>
        </form>
    </div>
</div>

<!-- RESUMO DO MERCADO -->
<div class="container mb-4">
    <div class="row">
        <div class="col-md-3">
            <div class="stock-card text-center">
                <div class="text-muted">Total Mercado</div>
                <div class="total-market">
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
        
        <div class="col-md-3">
            <div class="stock-card text-center">
                <div class="text-muted">Gerências Ativas</div>
                <div class="total-market">
                    <%= count %>
                </div>
                <div class="text-muted">No pregão</div>
            </div>
        </div>
        
        <div class="col-md-3">
            <div class="stock-card text-center">
                <div class="text-muted">Alta</div>
                <div class="total-market">
                    <%= altas %>
                </div>
                <div class="text-success">
                    <i class="fas fa-arrow-up me-1"></i>Em valorização
                </div>
            </div>
        </div>
        
        <div class="col-md-3">
            <div class="stock-card text-center">
                <div class="text-muted">Baixa</div>
                <div class="total-market">
                    <%= baixas %>
                </div>
                <div class="text-danger">
                    <i class="fas fa-arrow-down me-1"></i>Em desvalorização
                </div>
            </div>
        </div>
    </div>
</div>

<!-- PAINEL PRINCIPAL -->
<div class="container">
    <div class="row">
        <%
        If count > 0 Then
            ' Declarar todas as variáveis temporárias usadas na ordenação
            Dim tempNome, tempTotal, tempVar
            Dim tempM1, tempM2, tempM3, tempM4, tempM5, tempM6
            Dim tempM7, tempM8, tempM9, tempM10, tempM11, tempM12
            Dim i, j, m
            
            ' Ordena por total usando GetNum para garantir valores numéricos
            For i = 0 To count - 2
                For j = i + 1 To count - 1
                    If GetNum(gerenciaTotais(i)) < GetNum(gerenciaTotais(j)) Then
                        ' Troca nomes
                        tempNome = gerenciaNomes(i)
                        gerenciaNomes(i) = gerenciaNomes(j)
                        gerenciaNomes(j) = tempNome
                        
                        ' Troca totais
                        tempTotal = gerenciaTotais(i)
                        gerenciaTotais(i) = gerenciaTotais(j)
                        gerenciaTotais(j) = tempTotal
                        
                        ' Troca variações
                        tempVar = gerenciaVariacoes(i)
                        gerenciaVariacoes(i) = gerenciaVariacoes(j)
                        gerenciaVariacoes(j) = tempVar
                        
                        ' Troca meses
                        tempM1 = gerenciaMeses1(i)
                        gerenciaMeses1(i) = gerenciaMeses1(j)
                        gerenciaMeses1(j) = tempM1
                        
                        tempM2 = gerenciaMeses2(i)
                        gerenciaMeses2(i) = gerenciaMeses2(j)
                        gerenciaMeses2(j) = tempM2
                        
                        tempM3 = gerenciaMeses3(i)
                        gerenciaMeses3(i) = gerenciaMeses3(j)
                        gerenciaMeses3(j) = tempM3
                        
                        tempM4 = gerenciaMeses4(i)
                        gerenciaMeses4(i) = gerenciaMeses4(j)
                        gerenciaMeses4(j) = tempM4
                        
                        tempM5 = gerenciaMeses5(i)
                        gerenciaMeses5(i) = gerenciaMeses5(j)
                        gerenciaMeses5(j) = tempM5
                        
                        tempM6 = gerenciaMeses6(i)
                        gerenciaMeses6(i) = gerenciaMeses6(j)
                        gerenciaMeses6(j) = tempM6
                        
                        tempM7 = gerenciaMeses7(i)
                        gerenciaMeses7(i) = gerenciaMeses7(j)
                        gerenciaMeses7(j) = tempM7
                        
                        tempM8 = gerenciaMeses8(i)
                        gerenciaMeses8(i) = gerenciaMeses8(j)
                        gerenciaMeses8(j) = tempM8
                        
                        tempM9 = gerenciaMeses9(i)
                        gerenciaMeses9(i) = gerenciaMeses9(j)
                        gerenciaMeses9(j) = tempM9
                        
                        tempM10 = gerenciaMeses10(i)
                        gerenciaMeses10(i) = gerenciaMeses10(j)
                        gerenciaMeses10(j) = tempM10
                        
                        tempM11 = gerenciaMeses11(i)
                        gerenciaMeses11(i) = gerenciaMeses11(j)
                        gerenciaMeses11(j) = tempM11
                        
                        tempM12 = gerenciaMeses12(i)
                        gerenciaMeses12(i) = gerenciaMeses12(j)
                        gerenciaMeses12(j) = tempM12
                    End If
                Next
            Next
            
            ' Exibe cada gerência
            For i = 0 To count - 1
                Dim gNome, gTotal, gVariacao
                Dim changeClass, changeIcon, simbolo, palavras, palavra
                Dim participacao, totalFormatado, variacaoFormatada, participacaoFormatada
                
                gNome = gerenciaNomes(i)
                gTotal = gerenciaTotais(i)
                gVariacao = gerenciaVariacoes(i)
                
                ' Determina estilo
                If GetNum(gVariacao) > 0 Then
                    changeClass = "change-up"
                    changeIcon = "fas fa-arrow-up"
                ElseIf GetNum(gVariacao) < 0 Then
                    changeClass = "change-down"
                    changeIcon = "fas fa-arrow-down"
                Else
                    changeClass = "change-neutral"
                    changeIcon = "fas fa-minus"
                End If
                
                ' Símbolo
                simbolo = ""
                palavras = Split(gNome, " ")
                For Each palavra In palavras
                    If Len(palavra) > 0 Then
                        simbolo = simbolo & UCase(Left(palavra, 1))
                    End If
                Next
                If Len(simbolo) > 4 Then simbolo = Left(simbolo, 4)
                
                ' Participação
                If GetNum(totalGeral) > 0 Then
                    participacao = (GetNum(gTotal) / GetNum(totalGeral)) * 100
                Else
                    participacao = 0
                End If
                
                ' Formata números
                totalFormatado = FormatNumber(GetNum(gTotal), 2)
                variacaoFormatada = FormatNumber(Abs(GetNum(gVariacao)), 2)
                participacaoFormatada = FormatNumber(GetNum(participacao), 1)
        %>
        <div class="col-xl-4 col-lg-6 col-md-6 mb-4">
            <div class="stock-card">
                <div class="d-flex justify-content-between align-items-start">
                    <div>
                        <div class="stock-symbol"><%= simbolo %></div>
                        <div class="stock-name"><%= Left(gNome, 30) %></div>
                    </div>
                </div>
                
                <div class="stock-price">
                    R$ <%= totalFormatado %>
                </div>
                
                <div class="d-flex justify-content-between align-items-center">
                    <div>
                        <span class="<%= changeClass %>">
                            <i class="<%= changeIcon %> me-1"></i>
                            <%= variacaoFormatada %>%
                        </span>
                    </div>
                    <div class="text-muted">
                        Variação
                    </div>
                </div>
                
                <!-- Barra de participação -->
                <div class="volume-bar">
                    <div class="volume-fill" style="width: <%= participacaoFormatada %>%"></div>
                </div>
                <div class="text-end text-muted" style="font-size: 0.8rem;">
                    <%= participacaoFormatada %>% do mercado
                </div>
                
                <!-- Grade de meses -->
                <div class="month-grid">
                    <%
                    ' Corrigir a declaração do array mesesValores
                    Dim mesesValores
                    ReDim mesesValores(12)
                    
                    Dim mesNomes
                    ReDim mesNomes(12)
                    
                    mesesValores(1) = GetNum(gerenciaMeses1(i))
                    mesesValores(2) = GetNum(gerenciaMeses2(i))
                    mesesValores(3) = GetNum(gerenciaMeses3(i))
                    mesesValores(4) = GetNum(gerenciaMeses4(i))
                    mesesValores(5) = GetNum(gerenciaMeses5(i))
                    mesesValores(6) = GetNum(gerenciaMeses6(i))
                    mesesValores(7) = GetNum(gerenciaMeses7(i))
                    mesesValores(8) = GetNum(gerenciaMeses8(i))
                    mesesValores(9) = GetNum(gerenciaMeses9(i))
                    mesesValores(10) = GetNum(gerenciaMeses10(i))
                    mesesValores(11) = GetNum(gerenciaMeses11(i))
                    mesesValores(12) = GetNum(gerenciaMeses12(i))
                    
                    mesNomes(1) = "JAN"
                    mesNomes(2) = "FEV"
                    mesNomes(3) = "MAR"
                    mesNomes(4) = "ABR"
                    mesNomes(5) = "MAI"
                    mesNomes(6) = "JUN"
                    mesNomes(7) = "JUL"
                    mesNomes(8) = "AGO"
                    mesNomes(9) = "SET"
                    mesNomes(10) = "OUT"
                    mesNomes(11) = "NOV"
                    mesNomes(12) = "DEZ"
                    
                    Dim valorMes, valorAnterior, valorClass, valorExibir
                    Dim trendIcon, trendClass, variacaoMes
                    
                    For m = 1 To 12
                        valorMes = mesesValores(m)
                        
                        ' Calcula variação em relação ao mês anterior
                        valorAnterior = 0
                        If m > 1 Then
                            valorAnterior = mesesValores(m-1)
                        End If
                        
                        ' Determina tendência (seta)
                        variacaoMes = 0
                        If valorAnterior > 0 And valorMes > 0 Then
                            variacaoMes = ((valorMes - valorAnterior) / valorAnterior) * 100
                        ElseIf valorMes > 0 And valorAnterior = 0 Then
                            variacaoMes = 100 ' Cresceu de zero para algum valor
                        ElseIf valorMes = 0 And valorAnterior > 0 Then
                            variacaoMes = -100 ' Caiu para zero
                        End If
                        
                        ' Define ícone e classe da seta
                        If variacaoMes > 0 Then
                            trendIcon = "fas fa-arrow-up"
                            trendClass = "trend-up"
                        ElseIf variacaoMes < 0 Then
                            trendIcon = "fas fa-arrow-down"
                            trendClass = "trend-down"
                        Else
                            trendIcon = "fas fa-minus"
                            trendClass = "trend-neutral"
                        End If
                        
                        If valorMes > 0 Then
                            valorClass = "value-positive"
                        Else
                            valorClass = "value-zero"
                        End If
                        
                        If valorMes > 0 Then
                            valorExibir = FormatNumber(valorMes / 1000, 0) & "k"
                        Else
                            valorExibir = "-"
                        End If
                    %>
                    <div class="month-cell">
                        <div class="month-label"><%= mesNomes(m) %></div>
                        <div class="month-value-container">
                            <div class="month-value <%= valorClass %>">
                                <%= valorExibir %>
                            </div>
                            <% If valorMes > 0 And m > 1 Then ' Só mostra seta se houver valor e não for o primeiro mês %>
                            <div class="month-trend <%= trendClass %>" title="Variação: <%= FormatNumber(Abs(variacaoMes), 1) %>%">
                                <i class="<%= trendIcon %>"></i>
                            </div>
                            <% End If %>
                        </div>
                    </div>
                    <% Next %>
                </div>
            </div>
        </div>
        <% 
            Next 
        Else
        %>
        <div class="col-12">
            <div class="stock-card text-center py-5">
                <i class="fas fa-chart-bar fa-3x text-muted mb-3"></i>
                <h3 class="text-muted">Nenhuma gerência encontrada</h3>
                <p class="text-muted">Tente ajustar os filtros ou verificar os dados do ano <%= anoFiltro %></p>
            </div>
        </div>
        <% End If %>
    </div>
</div>

<!-- RODAPÉ -->
<div class="container mt-4">
    <div class="text-center text-muted">
        <p>
            <i class="fas fa-info-circle me-1"></i>
            Painel Bolsa de VGVs - Dados referentes ao ano <%= anoFiltro %> 
            | Atualizado em <%= Now() %> 
            | Total de ativos: <%= count %>
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