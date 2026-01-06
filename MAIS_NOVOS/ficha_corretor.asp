<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 18/12/2025               -->
<!-- CODIGO_ARQUIVO: KGWVTLEPKD          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->
<!--#include file="usr_acoes_v4GVendas.inc"-->


<%
' ===============================================
' CONFIGURAÇÃO UTF-8
' ===============================================
Response.CodePage = 65001 ' UTF-8
Response.CharSet = "UTF-8"
Response.ContentType = "text/html; charset=UTF-8"
%>

<%
' ===============================================
' CONFIGURAÇÃO DE BANCO DE DADOS
' ===============================================

Set connSales = Server.CreateObject("ADODB.Connection")
On Error Resume Next
connSales.Open StrConnSales

If Err.Number <> 0 Then
    Response.Write "<div class='alert alert-danger'>Erro ao conectar ao banco de dados: " & Err.Description & "</div>"
    Response.End
End If
On Error GoTo 0

' ===============================================
' OBTER PARÂMETROS DE FILTRO
' ===============================================

Dim filtroAno, filtroCorretor, modoRelatorio
filtroAno = Request.QueryString("ano")
filtroCorretor = Request.QueryString("corretor")
modoRelatorio = Request.QueryString("modo") ' "completo" ou "resumido"

If filtroAno = "" Then 
    filtroAno = Year(Date())
End If

If filtroCorretor = "" Then 
    filtroCorretor = "Todos"
End If

If modoRelatorio = "" Then 
    modoRelatorio = "completo"
End If

' ===============================================
' FUNÇÕES UTILITÁRIAS
' ===============================================

Function GetUniqueValues(tableName, columnName, whereClause)
    Dim dict, rs, sqlQuery
    Set dict = Server.CreateObject("Scripting.Dictionary")
    
    sqlQuery = "SELECT DISTINCT " & columnName & " FROM " & tableName & " "
    If whereClause <> "" Then
        sqlQuery = sqlQuery & " " & whereClause
    End If
    sqlQuery = sqlQuery & " ORDER BY " & columnName
    
    On Error Resume Next
    Set rs = connSales.Execute(sqlQuery)
    If Err.Number <> 0 Then
        GetUniqueValues = Array()
        Response.Write "<div class='alert alert-warning'>Erro na consulta: " & Err.Description & "</div>"
        Exit Function
    End If
    On Error GoTo 0
    
    If Not rs.EOF Then
        Do While Not rs.EOF
            If Not IsNull(rs(0)) Then
                dict.Add CStr(rs(0)), 1
            End If
            rs.MoveNext
        Loop
    End If
    
    If Not rs Is Nothing Then 
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    
    If dict.Count > 0 Then
        GetUniqueValues = dict.Keys
    Else
        GetUniqueValues = Array()
    End If
End Function

' Array com nomes dos meses
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

' ===============================================
' OBTER LISTA DE CORRETORES
' ===============================================

Dim uniqueCorretores, uniqueAnos
uniqueCorretores = GetUniqueValues("Vendas", "Corretor", "WHERE Corretor IS NOT NULL AND Corretor <> '' AND Corretor <> ' '")
uniqueAnos = GetUniqueValues("Vendas", "AnoVenda", "WHERE AnoVenda IS NOT NULL")

' ===============================================
' DADOS PRINCIPAIS - APENAS SE ANO ESTIVER PREENCHIDO
' ===============================================

Dim dadosCorretor, totalGeralVendas, totalGeralVGV, totalGeralComissao
Dim empreendimentosDict, localidadesDict, mesesComVendas, mesesSemVendas
Set dadosCorretor = Server.CreateObject("Scripting.Dictionary")

If filtroAno <> "" Then
    
    'Response.Write "<div class='alert alert-info'>Consultando ano: " & filtroAno & "</div>"
    
    ' Construir WHERE clause baseado no filtro
    Dim whereClause, sqlSafeCorretor
    whereClause = "WHERE Excluido = 0 AND AnoVenda = " & filtroAno
    
    If filtroCorretor <> "Todos" Then
        sqlSafeCorretor = Replace(filtroCorretor, "'", "''")
        whereClause = whereClause & " AND Corretor = '" & sqlSafeCorretor & "'"
    End If
    
   '' Response.Write "<div class='alert alert-info'>WHERE clause: " & whereClause & "</div>"
    
    ' ===============================================
    ' 1. DADOS MENSAIS DETALHADOS
    ' ===============================================
    
    Dim sqlDadosMensais, rsDadosMensais
    sqlDadosMensais = "SELECT " & _
                     "Corretor, " & _
                     "MesVenda, " & _
                     "COUNT(*) as QtdVendas, " & _
                     "SUM(ValorUnidade) as TotalVGV, " & _
                     "SUM(ValorCorretor) as TotalComissao " & _
                     "FROM Vendas " & _
                     whereClause & _
                     " GROUP BY Corretor, MesVenda " & _
                     "ORDER BY Corretor, MesVenda"

    'Response.Write "<div class='alert alert-info'>SQL: " & sqlDadosMensais & "</div>"
    
    Set rsDadosMensais = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rsDadosMensais.Open sqlDadosMensais, connSales
    
    If Err.Number <> 0 Then
        Response.Write "<div class='alert alert-danger'>Erro na consulta: " & Err.Description & "</div>"
    Else
        ' Inicializar dicionários
        Set empreendimentosDict = Server.CreateObject("Scripting.Dictionary")
        Set localidadesDict = Server.CreateObject("Scripting.Dictionary")
        Set mesesComVendas = Server.CreateObject("Scripting.Dictionary")
        Set mesesSemVendas = Server.CreateObject("Scripting.Dictionary")
        
        totalGeralVendas = 0
        totalGeralVGV = 0
        totalGeralComissao = 0
        
        If Not rsDadosMensais.EOF Then
            Response.Write "<div class='alert alert-success'>Registros encontrados: Sim</div>"
            
            Do While Not rsDadosMensais.EOF
                Dim corretorNome, mes, qtdVendas, totalVGV, totalComissao
                corretorNome = CStr(rsDadosMensais("Corretor"))
                mes = CStr(rsDadosMensais("MesVenda"))
                qtdVendas = CLng(rsDadosMensais("QtdVendas"))
                totalVGV = 0
                totalComissao = 0
                
                If Not IsNull(rsDadosMensais("TotalVGV")) Then
                    totalVGV = CDbl(rsDadosMensais("TotalVGV"))
                End If
                
                If Not IsNull(rsDadosMensais("TotalComissao")) Then
                    totalComissao = CDbl(rsDadosMensais("TotalComissao"))
                End If
                
                ' Adicionar corretor ao dicionário principal se não existir
                If Not dadosCorretor.Exists(corretorNome) Then
                    Dim infoCorretor
                    Set infoCorretor = Server.CreateObject("Scripting.Dictionary")
                    infoCorretor.Add "Meses", Server.CreateObject("Scripting.Dictionary")
                    infoCorretor.Add "TotalVendas", 0
                    infoCorretor.Add "TotalVGV", 0
                    infoCorretor.Add "TotalComissao", 0
                    infoCorretor.Add "Empreendimentos", Server.CreateObject("Scripting.Dictionary")
                    infoCorretor.Add "Localidades", Server.CreateObject("Scripting.Dictionary")
                    dadosCorretor.Add corretorNome, infoCorretor
                End If
                
                Set infoCorretor = dadosCorretor(corretorNome)
                
                ' Adicionar dados do mês
                infoCorretor("Meses").Add mes, Array(qtdVendas, totalVGV, totalComissao)
                
                ' Atualizar totais do corretor
                infoCorretor("TotalVendas") = infoCorretor("TotalVendas") + qtdVendas
                infoCorretor("TotalVGV") = infoCorretor("TotalVGV") + totalVGV
                infoCorretor("TotalComissao") = infoCorretor("TotalComissao") + totalComissao
                
                ' Marcar mês como com vendas
                mesesComVendas.Add mes, 1
                
                ' Atualizar totais gerais
                totalGeralVendas = totalGeralVendas + qtdVendas
                totalGeralVGV = totalGeralVGV + totalVGV
                totalGeralComissao = totalGeralComissao + totalComissao
                
                rsDadosMensais.MoveNext
            Loop
            
            ' Identificar meses sem vendas
            For i = 1 To 12
                If Not mesesComVendas.Exists(CStr(i)) Then
                    mesesSemVendas.Add CStr(i), arrMesesNome(i)
                End If
            Next
            
        Else
            Response.Write "<div class='alert alert-warning'>Nenhum registro encontrado para os filtros selecionados.</div>"
        End If
        
        rsDadosMensais.Close
    End If
    Set rsDadosMensais = Nothing
    
    ' ===============================================
    ' 2. OBTER EMPREENDIMENTOS E LOCALIDADES
    ' ===============================================
    
    If dadosCorretor.Count > 0 Then
        Dim sqlEmpreendLocal, rsEmpreendLocal
        sqlEmpreendLocal = "SELECT " & _
                          "Corretor, " & _
                          "Empreendimento, " & _
                          "Cidade " & _
                          "FROM Vendas " & _
                          whereClause & _
                          " AND Empreendimento IS NOT NULL " & _
                          " AND Cidade IS NOT NULL " & _
                          "GROUP BY Corretor, Empreendimento, Cidade"
        
        Set rsEmpreendLocal = Server.CreateObject("ADODB.Recordset")
        On Error Resume Next
        rsEmpreendLocal.Open sqlEmpreendLocal, connSales
        
        If Err.Number = 0 Then
            If Not rsEmpreendLocal.EOF Then
                Do While Not rsEmpreendLocal.EOF
                    Dim corretorEmp, empreendimento, localidade
                    corretorEmp = CStr(rsEmpreendLocal("Corretor"))
                    empreendimento = CStr(rsEmpreendLocal("Empreendimento"))
                    localidade = CStr(rsEmpreendLocal("Cidade"))
                    
                    If dadosCorretor.Exists(corretorEmp) Then
                        Set infoCorretor = dadosCorretor(corretorEmp)
                        
                        ' Adicionar empreendimento
                        If empreendimento <> "" And Not infoCorretor("Empreendimentos").Exists(empreendimento) Then
                            infoCorretor("Empreendimentos").Add empreendimento, 1
                            If Not empreendimentosDict.Exists(empreendimento) Then
                                empreendimentosDict.Add empreendimento, 1
                            End If
                        End If
                        
                        ' Adicionar localidade
                        If localidade <> "" And Not infoCorretor("Localidades").Exists(localidade) Then
                            infoCorretor("Localidades").Add localidade, 1
                            If Not localidadesDict.Exists(localidade) Then
                                localidadesDict.Add localidade, 1
                            End If
                        End If
                    End If
                    
                    rsEmpreendLocal.MoveNext
                Loop
            End If
            rsEmpreendLocal.Close
        End If
        Set rsEmpreendLocal = Nothing
    End If
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SGVendas - Ficha do Corretor</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <style>
        body {
            background-color: #f8f9fa;
            padding: 20px;
            color: #333;
        }
        .container-fluid {
            max-width: 1800px;
            margin: 0 auto;
        }
        .filter-container {
            background-color: #FFF;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .card-dashboard {
            background-color: #FFF;
            padding: 20px;
            margin-bottom: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .kpi-card {
            text-align: center;
            color: #fff;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 15px;
            min-height: 100px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }
        .kpi-card h5 {
            font-size: 0.9rem;
            margin-bottom: 5px;
            font-weight: bold;
        }
        .kpi-card p {
            margin: 0;
            font-size: 1.2rem;
            font-weight: bold;
        }
        .bg-primary-kpi { background-color: #007bff; }
        .bg-success-kpi { background-color: #28a745; }
        .bg-info-kpi { background-color: #17a2b8; }
        .bg-warning-kpi { background-color: #ffc107; color: #000; }
        .bg-danger-kpi { background-color: #dc3545; }
        .bg-purple-kpi { background-color: #6f42c1; }
        .bg-pink-kpi { background-color: #e83e8c; }
        .bg-teal-kpi { background-color: #20c997; }
        
        .table th {
            background-color: #800000;
            color: white;
        }
        .mes-com-venda { background-color: #d4edda !important; }
        .mes-sem-venda { background-color: #f8d7da !important; }
        
        .tab-content {
            background-color: #FFF;
            padding: 20px;
            border-radius: 0 0 8px 8px;
            border: 1px solid #dee2e6;
            border-top: none;
        }
        
        .nav-tabs .nav-link.active {
            background-color: #800000;
            color: white;
            border-color: #800000;
        }
        
        .nav-tabs .nav-link {
            color: #800000;
            font-weight: bold;
        }
        
        .debug-info {
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            padding: 10px;
            border-radius: 5px;
            margin-bottom: 10px;
            font-family: monospace;
            font-size: 12px;
        }
    </style>
</head>
<body>
    <div class="container-fluid">
        <h2 class="mt-4 mb-4 text-center" style="color: #800000;">
            <i class="fas fa-user-tie"></i> SGVendas - Ficha do Corretor
        </h2>
        
        <!-- Filtros -->
        <div class="filter-container">
            <form id="filterForm" method="get">
                <div class="row">
                    <div class="col-md-3">
                        <div class="mb-3">
                            <label class="form-label">Ano</label>
                            <select class="form-select" name="ano" id="anoFilter" required>
                                <option value="">Selecione o ano</option>
                                <%
                                If IsArray(uniqueAnos) Then
                                    For Each ano In uniqueAnos
                                        Response.Write "<option value=""" & ano & """"
                                        If CStr(filtroAno) = CStr(ano) Then Response.Write " selected"
                                        Response.Write ">" & ano & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                    </div>
                    
                    <div class="col-md-4">
                        <div class="mb-3">
                            <label class="form-label">Corretor</label>
                            <select class="form-select" name="corretor" id="corretorFilter">
                                <option value="Todos">Todos os Corretores</option>
                                <%
                                If IsArray(uniqueCorretores) Then
                                    For Each corretor In uniqueCorretores
                                        Response.Write "<option value=""" & Server.HTMLEncode(corretor) & """"
                                        If CStr(filtroCorretor) = CStr(corretor) Then Response.Write " selected"
                                        Response.Write ">" & corretor & "</option>"
                                    Next
                                End If
                                %>
                            </select>
                        </div>
                    </div>
                    
                    <div class="col-md-3">
                        <div class="mb-3">
                            <label class="form-label">Modo de Visualização</label>
                            <select class="form-select" name="modo">
                                <option value="completo" <% If modoRelatorio = "completo" Then Response.Write "selected" %>>Relatório Completo</option>
                                <option value="resumido" <% If modoRelatorio = "resumido" Then Response.Write "selected" %>>Relatório Resumido</option>
                            </select>
                        </div>
                    </div>
                    
                    <div class="col-md-2 d-flex align-items-end">
                        <button type="submit" class="btn btn-primary w-100">
                            <i class="fas fa-chart-bar"></i> Gerar Relatório
                        </button>
                    </div>
                </div>
            </form>
        </div>
        
        <!-- Informações de Debug -->
        <div class="debug-info">
            <strong>Informações do Sistema:</strong><br>
            Ano filtrado: <%= filtroAno %><br>
            Corretor filtrado: <%= filtroCorretor %><br>
            Total de corretores na lista: 
            <%
            If IsArray(uniqueCorretores) Then
                Response.Write UBound(uniqueCorretores) + 1
            Else
                Response.Write "0"
            End If
            %><br>
            Dados encontrados: <%= dadosCorretor.Count %>
        </div>
        
        <% If filtroAno = "" Then %>
            <div class="alert alert-warning text-center">
                <i class="fas fa-info-circle"></i> Por favor, selecione um ano para visualizar a ficha do corretor.
            </div>
        <% ElseIf dadosCorretor.Count = 0 Then %>
            <div class="alert alert-info text-center">
                <i class="fas fa-info-circle"></i> Nenhum dado encontrado para os filtros selecionados.
            </div>
        <% Else %>
        
        <!-- KPIs Gerais -->
        <div class="row mt-4">
            <div class="col-md-3">
                <div class="kpi-card bg-primary-kpi">
                    <h5>Total de Corretores</h5>
                    <p><%= dadosCorretor.Count %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-success-kpi">
                    <h5>Total de Vendas</h5>
                    <p><%= totalGeralVendas %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-info-kpi">
                    <h5>VGV Total (R$)</h5>
                    <p><%= FormatNumber(totalGeralVGV, 2) %></p>
                </div>
            </div>
            <div class="col-md-3">
                <div class="kpi-card bg-warning-kpi">
                    <h5>Comissão Total (R$)</h5>
                    <p><%= FormatNumber(totalGeralComissao, 2) %></p>
                </div>
            </div>
        </div>
        
        <!-- Tabs de Navegação -->
        <ul class="nav nav-tabs mt-4" id="myTab" role="tablist">
            <li class="nav-item" role="presentation">
                <button class="nav-link active" id="resumo-tab" data-bs-toggle="tab" data-bs-target="#resumo" type="button" role="tab">Resumo Geral</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="mensal-tab" data-bs-toggle="tab" data-bs-target="#mensal" type="button" role="tab">Dados Mensais</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="empreendimentos-tab" data-bs-toggle="tab" data-bs-target="#empreendimentos" type="button" role="tab">Empreendimentos</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="localidades-tab" data-bs-toggle="tab" data-bs-target="#localidades" type="button" role="tab">Localidades</button>
            </li>
        </ul>
        
        <div class="tab-content" id="myTabContent">
            
            <!-- Tab 1: Resumo Geral -->
            <div class="tab-pane fade show active" id="resumo" role="tabpanel">
                <% 
                Dim arrCorretoresResumo
                arrCorretoresResumo = dadosCorretor.Keys
                
                ' Ordenar corretores por total de vendas (decrescente)
                If IsArray(arrCorretoresResumo) Then
                    For i = 0 To UBound(arrCorretoresResumo)
                        For j = i + 1 To UBound(arrCorretoresResumo)
                            If dadosCorretor(arrCorretoresResumo(j))("TotalVendas") > dadosCorretor(arrCorretoresResumo(i))("TotalVendas") Then
                                Dim tempCorretor
                                tempCorretor = arrCorretoresResumo(i)
                                arrCorretoresResumo(i) = arrCorretoresResumo(j)
                                arrCorretoresResumo(j) = tempCorretor
                            End If
                        Next
                    Next
                End If
                %>
                
                <div class="row">
                    <div class="col-md-8">
                        <h4>Resumo por Corretor - Ano <%= filtroAno %></h4>
                        <div class="table-responsive">
                            <table class="table table-striped table-hover">
                                <thead>
                                    <tr>
                                        <th>Corretor</th>
                                        <th class="text-center">Vendas</th>
                                        <th class="text-end">VGV (R$)</th>
                                        <th class="text-end">Comissão (R$)</th>
                                        <th class="text-center">Média/Venda</th>
                                        <th class="text-center">Empreend.</th>
                                        <th class="text-center">Localidades</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    If IsArray(arrCorretoresResumo) Then
                                        For Each corretorKey In arrCorretoresResumo
                                            Set infoCorretor = dadosCorretor(corretorKey)
                                            Dim mediaVenda
                                            If infoCorretor("TotalVendas") > 0 Then
                                                mediaVenda = infoCorretor("TotalVGV") / infoCorretor("TotalVendas")
                                            Else
                                                mediaVenda = 0
                                            End If
                                    %>
                                    <tr>
                                        <td><strong><%= corretorKey %></strong></td>
                                        <td class="text-center"><%= infoCorretor("TotalVendas") %></td>
                                        <td class="text-end"><%= FormatNumber(infoCorretor("TotalVGV"), 2) %></td>
                                        <td class="text-end"><%= FormatNumber(infoCorretor("TotalComissao"), 2) %></td>
                                        <td class="text-end"><%= FormatNumber(mediaVenda, 2) %></td>
                                        <td class="text-center">
                                            <%
                                            If Not infoCorretor("Empreendimentos") Is Nothing Then
                                                Response.Write infoCorretor("Empreendimentos").Count
                                            Else
                                                Response.Write "0"
                                            End If
                                            %>
                                        </td>
                                        <td class="text-center">
                                            <%
                                            If Not infoCorretor("Localidades") Is Nothing Then
                                                Response.Write infoCorretor("Localidades").Count
                                            Else
                                                Response.Write "0"
                                            End If
                                            %>
                                        </td>
                                    </tr>
                                    <%
                                        Next
                                    End If
                                    %>
                                </tbody>
                                <tfoot>
                                    <tr class="table-dark">
                                        <td><strong>TOTAIS</strong></td>
                                        <td class="text-center"><strong><%= totalGeralVendas %></strong></td>
                                        <td class="text-end"><strong><%= FormatNumber(totalGeralVGV, 2) %></strong></td>
                                        <td class="text-end"><strong><%= FormatNumber(totalGeralComissao, 2) %></strong></td>
                                        <td class="text-end">
                                            <strong>
                                            <%
                                            If totalGeralVendas > 0 Then
                                                Response.Write FormatNumber(totalGeralVGV / totalGeralVendas, 2)
                                            Else
                                                Response.Write "0,00"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                        <td class="text-center">
                                            <strong>
                                            <%
                                            If Not empreendimentosDict Is Nothing Then
                                                Response.Write empreendimentosDict.Count
                                            Else
                                                Response.Write "0"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                        <td class="text-center">
                                            <strong>
                                            <%
                                            If Not localidadesDict Is Nothing Then
                                                Response.Write localidadesDict.Count
                                            Else
                                                Response.Write "0"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    </div>
                    
                    <div class="col-md-4">
                        <h4>Estatísticas do Período</h4>
                        
                        <div class="row mb-4">
                            <div class="col-6">
                                <div class="kpi-card bg-teal-kpi">
                                    <h5>Meses com Vendas</h5>
                                    <p>
                                        <%
                                        If Not mesesComVendas Is Nothing Then
                                            Response.Write mesesComVendas.Count
                                        Else
                                            Response.Write "0"
                                        End If
                                        %>
                                    </p>
                                </div>
                            </div>
                            <div class="col-6">
                                <div class="kpi-card bg-pink-kpi">
                                    <h5>Meses sem Vendas</h5>
                                    <p>
                                        <%
                                        If Not mesesSemVendas Is Nothing Then
                                            Response.Write mesesSemVendas.Count
                                        Else
                                            Response.Write "12"
                                        End If
                                        %>
                                    </p>
                                </div>
                            </div>
                        </div>
                        
                        <% If Not mesesSemVendas Is Nothing And mesesSemVendas.Count > 0 Then %>
                        <div class="card mb-3">
                            <div class="card-header">
                                <h5 class="mb-0">Meses sem Vendas</h5>
                            </div>
                            <div class="card-body">
                                <%
                                Dim arrMesesSemVenda
                                arrMesesSemVenda = mesesSemVendas.Keys
                                
                                For Each mesKey In arrMesesSemVenda
                                    Response.Write "<span class='badge bg-danger me-1 mb-1'>" & mesesSemVendas(mesKey) & "</span>"
                                Next
                                %>
                            </div>
                        </div>
                        <% End If %>
                        
                        <div class="card">
                            <div class="card-header">
                                <h5 class="mb-0">Top 3 Corretores</h5>
                            </div>
                            <div class="card-body">
                                <div class="list-group">
                                    <%
                                    If IsArray(arrCorretoresResumo) Then
                                        Dim contadorTop
                                        contadorTop = 0
                                        For Each corretorKey In arrCorretoresResumo
                                            If contadorTop < 3 Then
                                                Set infoCorretor = dadosCorretor(corretorKey)
                                    %>
                                    <div class="list-group-item list-group-item-success">
                                        <div class="d-flex w-100 justify-content-between">
                                            <h6 class="mb-1"><%= contadorTop + 1 %>. <%= corretorKey %></h6>
                                            <small><%= infoCorretor("TotalVendas") %> vendas</small>
                                        </div>
                                        <p class="mb-1">VGV: R$ <%= FormatNumber(infoCorretor("TotalVGV"), 2) %></p>
                                        <small>Comissão: R$ <%= FormatNumber(infoCorretor("TotalComissao"), 2) %></small>
                                    </div>
                                    <%
                                            contadorTop = contadorTop + 1
                                            End If
                                        Next
                                    End If
                                    %>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Tab 2: Dados Mensais -->
            <div class="tab-pane fade" id="mensal" role="tabpanel">
                <h4>Desempenho Mensal - Ano <%= filtroAno %></h4>
                
                <%
                ' Calcular totais mensais para todos os corretores
                Dim totaisMensais(12, 3) ' Índice 0: Vendas, 1: VGV, 2: Comissão
                
                For i = 1 To 12
                    totaisMensais(i, 0) = 0
                    totaisMensais(i, 1) = 0
                    totaisMensais(i, 2) = 0
                Next
                %>
                
                <div class="table-responsive">
                    <table class="table table-striped table-bordered">
                        <thead>
                            <tr>
                                <th rowspan="2" class="align-middle">Corretor</th>
                                <%
                                ' Cabeçalhos dos meses
                                For i = 1 To 12
                                    Dim temVendaMes
                                    temVendaMes = False
                                    If Not mesesComVendas Is Nothing Then
                                        temVendaMes = mesesComVendas.Exists(CStr(i))
                                    End If
                                    Response.Write "<th colspan='3' class='text-center " & IIf(temVendaMes, "mes-com-venda", "mes-sem-venda") & "'>" & _
                                                  Left(arrMesesNome(i), 3) & "</th>"
                                Next
                                %>
                                <th colspan="3" class="text-center table-dark">TOTAIS</th>
                            </tr>
                            <tr>
                                <%
                                ' Subcabeçalhos para cada mês
                                For i = 1 To 12
                                    Response.Write "<th class='text-center small'>Vendas</th>"
                                    Response.Write "<th class='text-center small'>VGV</th>"
                                    Response.Write "<th class='text-center small'>Comissão</th>"
                                Next
                                %>
                                <th class="text-center small table-dark">Vendas</th>
                                <th class="text-center small table-dark">VGV</th>
                                <th class="text-center small table-dark">Comissão</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            If IsArray(arrCorretoresResumo) Then
                                For Each corretorKey In arrCorretoresResumo
                                    Set infoCorretor = dadosCorretor(corretorKey)
                                    Set mesesCorretor = infoCorretor("Meses")
                            %>
                            <tr>
                                <td><strong><%= corretorKey %></strong></td>
                                <%
                                For i = 1 To 12
                                    If mesesCorretor.Exists(CStr(i)) Then
                                        Dim dadosMes
                                        dadosMes = mesesCorretor(CStr(i))
                                        totaisMensais(i, 0) = totaisMensais(i, 0) + dadosMes(0)
                                        totaisMensais(i, 1) = totaisMensais(i, 1) + dadosMes(1)
                                        totaisMensais(i, 2) = totaisMensais(i, 2) + dadosMes(2)
                                        
                                        Response.Write "<td class='text-center mes-com-venda'>" & dadosMes(0) & "</td>"
                                        Response.Write "<td class='text-end mes-com-venda'>" & FormatNumber(dadosMes(1), 0) & "</td>"
                                        Response.Write "<td class='text-end mes-com-venda'>" & FormatNumber(dadosMes(2), 0) & "</td>"
                                    Else
                                        Response.Write "<td colspan='3' class='text-center mes-sem-venda'>-</td>"
                                    End If
                                Next
                                %>
                                <td class="text-center table-dark"><strong><%= infoCorretor("TotalVendas") %></strong></td>
                                <td class="text-end table-dark"><strong><%= FormatNumber(infoCorretor("TotalVGV"), 0) %></strong></td>
                                <td class="text-end table-dark"><strong><%= FormatNumber(infoCorretor("TotalComissao"), 0) %></strong></td>
                            </tr>
                            <%
                                Next
                            End If
                            %>
                        </tbody>
                        <tfoot>
                            <tr class="table-dark">
                                <td><strong>TOTAIS MENSAL</strong></td>
                                <%
                                Dim totalMensalVendas, totalMensalVGV, totalMensalComissao
                                totalMensalVendas = 0
                                totalMensalVGV = 0
                                totalMensalComissao = 0
                                
                                For i = 1 To 12
                                    totalMensalVendas = totalMensalVendas + totaisMensais(i, 0)
                                    totalMensalVGV = totalMensalVGV + totaisMensais(i, 1)
                                    totalMensalComissao = totalMensalComissao + totaisMensais(i, 2)
                                    
                                    Response.Write "<td class='text-center'>" & totaisMensais(i, 0) & "</td>"
                                    Response.Write "<td class='text-end'>" & FormatNumber(totaisMensais(i, 1), 0) & "</td>"
                                    Response.Write "<td class='text-end'>" & FormatNumber(totaisMensais(i, 2), 0) & "</td>"
                                Next
                                %>
                                <td class="text-center"><strong><%= totalMensalVendas %></strong></td>
                                <td class="text-end"><strong><%= FormatNumber(totalMensalVGV, 0) %></strong></td>
                                <td class="text-end"><strong><%= FormatNumber(totalMensalComissao, 0) %></strong></td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
            
            <!-- Tab 3: Empreendimentos -->
            <div class="tab-pane fade" id="empreendimentos" role="tabpanel">
                <div class="row">
                    <div class="col-md-8">
                        <h4>Empreendimentos Vendidos - Ano <%= filtroAno %></h4>
                        
                        <%
                        If Not empreendimentosDict Is Nothing And empreendimentosDict.Count > 0 Then
                            Dim arrEmpreendimentosTotal
                            arrEmpreendimentosTotal = empreendimentosDict.Keys
                        %>
                        
                        <div class="table-responsive">
                            <table class="table table-striped">
                                <thead>
                                    <tr>
                                        <th>Empreendimento</th>
                                        <th class="text-center">Corretores</th>
                                        <th class="text-center">Qtd Vendas</th>
                                        <th class="text-center">Localidades</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    For Each empreend In arrEmpreendimentosTotal
                                        ' Contar corretores que venderam este empreendimento
                                        Dim corretoresNoEmpreend, localidadesNoEmpreend, vendasNoEmpreend
                                        Set corretoresNoEmpreend = Server.CreateObject("Scripting.Dictionary")
                                        Set localidadesNoEmpreend = Server.CreateObject("Scripting.Dictionary")
                                        vendasNoEmpreend = 0
                                        
                                        For Each corretorKey In dadosCorretor.Keys
                                            Set infoCorretor = dadosCorretor(corretorKey)
                                            If Not infoCorretor("Empreendimentos") Is Nothing Then
                                                If infoCorretor("Empreendimentos").Exists(empreend) Then
                                                    corretoresNoEmpreend.Add corretorKey, 1
                                                    
                                                    ' Contar localidades deste empreendimento para este corretor
                                                    If Not infoCorretor("Localidades") Is Nothing Then
                                                        Dim arrLocalCorretor
                                                        arrLocalCorretor = infoCorretor("Localidades").Keys
                                                        For Each localCorretor In arrLocalCorretor
                                                            If Not localidadesNoEmpreend.Exists(localCorretor) Then
                                                                localidadesNoEmpreend.Add localCorretor, 1
                                                            End If
                                                        Next
                                                    End If
                                                End If
                                            End If
                                        Next
                                    %>
                                    <tr>
                                        <td><%= empreend %></td>
                                        <td class="text-center">
                                            <span class="badge bg-primary"><%= corretoresNoEmpreend.Count %></span>
                                        </td>
                                        <td class="text-center">
                                            <%
                                            ' Tentar obter quantidade de vendas para este empreendimento
                                            On Error Resume Next
                                            Dim sqlVendasEmp, rsVendasEmp
                                            sqlVendasEmp = "SELECT COUNT(*) as Total FROM Vendas WHERE Excluido = 0 AND AnoVenda = " & filtroAno & _
                                                           " AND Empreendimento = '" & Replace(empreend, "'", "''") & "'"
                                            If filtroCorretor <> "Todos" Then
                                                sqlVendasEmp = sqlVendasEmp & " AND Corretor = '" & Replace(filtroCorretor, "'", "''") & "'"
                                            End If
                                            
                                            Set rsVendasEmp = connSales.Execute(sqlVendasEmp)
                                            If Err.Number = 0 Then
                                                If Not rsVendasEmp.EOF Then
                                                    Response.Write rsVendasEmp("Total")
                                                End If
                                                rsVendasEmp.Close
                                            End If
                                            On Error GoTo 0
                                            %>
                                        </td>
                                        <td class="text-center">
                                            <span class="badge bg-info"><%= localidadesNoEmpreend.Count %></span>
                                        </td>
                                    </tr>
                                    <%
                                    Next
                                    %>
                                </tbody>
                                <tfoot>
                                    <tr class="table-dark">
                                        <td><strong>TOTAIS</strong></td>
                                        <td class="text-center"><strong><%= empreendimentosDict.Count %></strong></td>
                                        <td class="text-center"><strong><%= totalGeralVendas %></strong></td>
                                        <td class="text-center">
                                            <strong>
                                            <%
                                            If Not localidadesDict Is Nothing Then
                                                Response.Write localidadesDict.Count
                                            Else
                                                Response.Write "0"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                        <%
                        Else
                        %>
                        <div class="alert alert-info">
                            Nenhum dado de empreendimento disponível para os filtros selecionados.
                        </div>
                        <%
                        End If
                        %>
                    </div>
                    
                    <div class="col-md-4">
                        <h4>Distribuição por Corretor</h4>
                        <div class="list-group">
                            <%
                            If IsArray(arrCorretoresResumo) Then
                                For Each corretorKey In arrCorretoresResumo
                                    Set infoCorretor = dadosCorretor(corretorKey)
                                    If Not infoCorretor("Empreendimentos") Is Nothing Then
                                        If infoCorretor("Empreendimentos").Count > 0 Then
                            %>
                            <div class="list-group-item">
                                <div class="d-flex w-100 justify-content-between">
                                    <h6 class="mb-1"><%= corretorKey %></h6>
                                    <small class="badge bg-primary rounded-pill"><%= infoCorretor("Empreendimentos").Count %></small>
                                </div>
                                <p class="mb-1 small">
                                    <%
                                    Dim arrEmpCorretor
                                    arrEmpCorretor = infoCorretor("Empreendimentos").Keys
                                    Dim empCount
                                    empCount = 0
                                    For Each emp In arrEmpCorretor
                                        If empCount < 3 Then ' Mostrar apenas os 3 primeiros
                                            Response.Write "<span class='badge bg-info me-1 mb-1'>" & emp & "</span>"
                                            empCount = empCount + 1
                                        End If
                                    Next
                                    If infoCorretor("Empreendimentos").Count > 3 Then
                                        Response.Write "<span class='badge bg-secondary'>+" & (infoCorretor("Empreendimentos").Count - 3) & " mais</span>"
                                    End If
                                    %>
                                </p>
                            </div>
                            <%
                                        End If
                                    End If
                                Next
                            End If
                            %>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Tab 4: Localidades -->
            <div class="tab-pane fade" id="localidades" role="tabpanel">
                <div class="row">
                    <div class="col-md-8">
                        <h4>Localidades - Ano <%= filtroAno %></h4>
                        
                        <%
                        If Not localidadesDict Is Nothing And localidadesDict.Count > 0 Then
                            Dim arrLocalidadesTotal
                            arrLocalidadesTotal = localidadesDict.Keys
                        %>
                        
                        <div class="table-responsive">
                            <table class="table table-striped">
                                <thead>
                                    <tr>
                                        <th>Localidade</th>
                                        <th class="text-center">Corretores</th>
                                        <th class="text-center">Qtd Vendas</th>
                                        <th class="text-center">Empreendimentos</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    For Each localidade In arrLocalidadesTotal
                                        ' Contar corretores nesta localidade
                                        Dim corretoresNaLocalidade, empreendimentosNaLocalidade, vendasNaLocalidade
                                        Set corretoresNaLocalidade = Server.CreateObject("Scripting.Dictionary")
                                        Set empreendimentosNaLocalidade = Server.CreateObject("Scripting.Dictionary")
                                        vendasNaLocalidade = 0
                                        
                                        For Each corretorKey In dadosCorretor.Keys
                                            Set infoCorretor = dadosCorretor(corretorKey)
                                            If Not infoCorretor("Localidades") Is Nothing Then
                                                If infoCorretor("Localidades").Exists(localidade) Then
                                                    corretoresNaLocalidade.Add corretorKey, 1
                                                    
                                                    ' Contar empreendimentos desta localidade para este corretor
                                                    If Not infoCorretor("Empreendimentos") Is Nothing Then
                                                        Dim arrEmpCorretorLocal
                                                        arrEmpCorretorLocal = infoCorretor("Empreendimentos").Keys
                                                        For Each empCorretor In arrEmpCorretorLocal
                                                            If Not empreendimentosNaLocalidade.Exists(empCorretor) Then
                                                                empreendimentosNaLocalidade.Add empCorretor, 1
                                                            End If
                                                        Next
                                                    End If
                                                End If
                                            End If
                                        Next
                                    %>
                                    <tr>
                                        <td><%= localidade %></td>
                                        <td class="text-center">
                                            <span class="badge bg-primary"><%= corretoresNaLocalidade.Count %></span>
                                        </td>
                                        <td class="text-center">
                                            <%
                                            ' Tentar obter quantidade de vendas para esta localidade
                                            On Error Resume Next
                                            Dim sqlVendasLoc, rsVendasLoc
                                            sqlVendasLoc = "SELECT COUNT(*) as Total FROM Vendas WHERE Excluido = 0 AND AnoVenda = " & filtroAno & _
                                                           " AND Cidade = '" & Replace(localidade, "'", "''") & "'"
                                            If filtroCorretor <> "Todos" Then
                                                sqlVendasLoc = sqlVendasLoc & " AND Corretor = '" & Replace(filtroCorretor, "'", "''") & "'"
                                            End If
                                            
                                            Set rsVendasLoc = connSales.Execute(sqlVendasLoc)
                                            If Err.Number = 0 Then
                                                If Not rsVendasLoc.EOF Then
                                                    Response.Write rsVendasLoc("Total")
                                                End If
                                                rsVendasLoc.Close
                                            End If
                                            On Error GoTo 0
                                            %>
                                        </td>
                                        <td class="text-center">
                                            <span class="badge bg-info"><%= empreendimentosNaLocalidade.Count %></span>
                                        </td>
                                    </tr>
                                    <%
                                    Next
                                    %>
                                </tbody>
                                <tfoot>
                                    <tr class="table-dark">
                                        <td><strong>TOTAIS</strong></td>
                                        <td class="text-center">
                                            <strong>
                                            <%
                                            If Not localidadesDict Is Nothing Then
                                                Response.Write localidadesDict.Count
                                            Else
                                                Response.Write "0"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                        <td class="text-center"><strong><%= totalGeralVendas %></strong></td>
                                        <td class="text-center">
                                            <strong>
                                            <%
                                            If Not empreendimentosDict Is Nothing Then
                                                Response.Write empreendimentosDict.Count
                                            Else
                                                Response.Write "0"
                                            End If
                                            %>
                                            </strong>
                                        </td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                        <%
                        Else
                        %>
                        <div class="alert alert-info">
                            Nenhum dado de localidade disponível para os filtros selecionados.
                        </div>
                        <%
                        End If
                        %>
                    </div>
                    
                    <div class="col-md-4">
                        <h4>Mapa de Atuação</h4>
                        <div class="card">
                            <div class="card-body">
                                <h6 class="card-title">Resumo Geográfico</h6>
                                
                                <p><i class="fas fa-map-marker-alt text-danger"></i> <strong>Total de Localidades:</strong> 
                                <%
                                If Not localidadesDict Is Nothing Then
                                    Response.Write localidadesDict.Count
                                Else
                                    Response.Write "0"
                                End If
                                %>
                                </p>
                                
                                <p><i class="fas fa-building text-primary"></i> <strong>Empreendimentos Únicos:</strong> 
                                <%
                                If Not empreendimentosDict Is Nothing Then
                                    Response.Write empreendimentosDict.Count
                                Else
                                    Response.Write "0"
                                End If
                                %>
                                </p>
                                
                                <p><i class="fas fa-chart-pie text-success"></i> <strong>Diversificação:</strong> 
                                <%
                                If Not localidadesDict Is Nothing And Not empreendimentosDict Is Nothing Then
                                    Dim scoreDiversificacao
                                    scoreDiversificacao = (localidadesDict.Count * 2) + (empreendimentosDict.Count * 3)
                                    Response.Write scoreDiversificacao & " pontos"
                                Else
                                    Response.Write "N/A"
                                End If
                                %>
                                </p>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Botões de Ação -->
        <div class="row mt-4">
            <div class="col-12">
                <div class="d-flex justify-content-between">
                    <button class="btn btn-secondary" onclick="window.print()">
                        <i class="fas fa-print"></i> Imprimir Relatório
                    </button>
                    <button class="btn btn-success" onclick="exportToExcel()">
                        <i class="fas fa-file-excel"></i> Exportar para Excel
                    </button>
                </div>
            </div>
        </div>
        
        <% End If %>
    </div>

    <!-- Scripts -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    
    <script>
    $(document).ready(function() {
        // Inicializar tabs do Bootstrap
        var triggerTabList = [].slice.call(document.querySelectorAll('#myTab button'))
        triggerTabList.forEach(function (triggerEl) {
            var tabTrigger = new bootstrap.Tab(triggerEl)
            triggerEl.addEventListener('click', function (event) {
                event.preventDefault()
                tabTrigger.show()
            })
        });
    });
    
    function exportToExcel() {
        // Esta função pode ser expandida para exportar dados para Excel
        alert('Funcionalidade de exportação para Excel será implementada em breve!');
    }
    </script>
</body>
</html>

<%
' Fechar conexão
If connSales.State = 1 Then connSales.Close
Set connSales = Nothing
%>