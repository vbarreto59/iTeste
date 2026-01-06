<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 18/12/2025               -->
<!-- CODIGO_ARQUIVO: JWXHKDGTKZ          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="conexao.asp"-->        <!-- StrConn (Diretorias / Gerencias) -->
<!--#include file="conSunSales.asp"-->    <!-- StrConnSales (Metas: MetaEmpresa, MetaDiretoria, MetaGerencia) -->

<%
'-----------------------------
' CONFIGURAÇÃO E ABERTURA DE CONEXÕES
'-----------------------------
Response.CodePage = 65001
Response.Charset = "utf-8"

Dim connDG, connSales
Set connDG = Server.CreateObject("ADODB.Connection")
connDG.Open StrConn    ' conexão para Diretoria / Gerencia (cadastros)

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales  ' conexão para todas as tabelas Meta*

'-----------------------------
' FUNÇÕES AUXILIARES CORRIGIDAS
'-----------------------------
Function ValorSafe(valor)
    On Error Resume Next
    If IsNull(valor) Then
        ValorSafe = 0
    Else
        If IsNumeric(valor) Then
            ValorSafe = CDbl(valor)
        Else
            ValorSafe = 0
        End If
    End If
    On Error GoTo 0
End Function

Function FxFormat(v)
    If IsNumeric(v) Then
        FxFormat = FormatNumber(v, 2)
    Else
        FxFormat = FormatNumber(0, 2)
    End If
End Function

Function NomeMes(numeroMes)
    Dim meses
    meses = Array("", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", _
                  "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro")
    
    If numeroMes >= 1 And numeroMes <= 12 Then
        NomeMes = meses(numeroMes)
    Else
        NomeMes = "Mês " & numeroMes
    End If
End Function

'-----------------------------
' PARÂMETROS DO RELATÓRIO
'-----------------------------
Dim anoFiltro, mesFiltro
anoFiltro = Request.QueryString("ano")
mesFiltro = Request.QueryString("mes")

' Valores padrão se não informados
If anoFiltro = "" Then
    anoFiltro = Year(Now())
End If

If mesFiltro = "" Then
    mesFiltro = Month(Now())
End If

'-----------------------------
' BUSCAR DADOS DA EMPRESA (META)
'-----------------------------
Dim sqlEmpresa, rsEmpresa, metaEmpresa
sqlEmpresa = "SELECT Meta FROM MetaEmpresa WHERE Ano = " & anoFiltro & " AND Mes = " & mesFiltro
Set rsEmpresa = connSales.Execute(sqlEmpresa)

metaEmpresa = 0
If Not rsEmpresa.EOF Then
    metaEmpresa = ValorSafe(rsEmpresa("Meta"))
End If

If Not rsEmpresa Is Nothing Then
    rsEmpresa.Close
    Set rsEmpresa = Nothing
End If

'-----------------------------
' BUSCAR DIRETORIAS COM SUAS METAS
'-----------------------------
Dim sqlDiretorias, rsDiretorias
sqlDiretorias = "SELECT " & _
                "d.DiretoriaID, " & _
                "d.NomeDiretoria, " & _
                "md.TotalMetas, " & _
                "md.Usuario, " & _
                "md.DataHora " & _
                "FROM Diretorias d " & _
                "LEFT JOIN MetaDiretoria md ON d.DiretoriaID = md.DiretoriaID " & _
                "AND md.Ano = " & anoFiltro & " AND md.Mes = " & mesFiltro & " " & _
                "ORDER BY d.NomeDiretoria"

Set rsDiretorias = connSales.Execute(sqlDiretorias)

'-----------------------------
' CALCULAR TOTAIS
'-----------------------------
Dim totalDiretorias, totalGerencias
totalDiretorias = 0
totalGerencias = 0
%>

<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<title>Relatório Detalhado de Metas</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css">
<style>
    body { background: #f8f9fa; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .header-bg { background: linear-gradient(135deg, #2c3e50, #3498db); color: white; }
    .card-empresa { border: 2px solid #3498db; }
    .card-diretoria { border-left: 4px solid #2ecc71; }
    .card-gerencia { border-left: 2px solid #f39c12; background-color: #fefefe; }
    .valor-meta { font-weight: 600; color: #2c3e50; }
    .valor-zero { color: #95a5a6; font-style: italic; }
    .badge-diretoria { background-color: #2ecc71; }
    .badge-gerencia { background-color: #f39c12; }
    .accordion-button:not(.collapsed) { background-color: #e3f2fd; color: #1565c0; }
    .total-box { background-color: #2c3e50; color: white; border-radius: 5px; padding: 10px; }
    .mes-selector { background-color: #e3f2fd; border-radius: 5px; padding: 15px; margin-bottom: 20px; }
    .hierarchy-icon { color: #3498db; margin-right: 8px; }
</style>
</head>
<body>

<div class="container-fluid py-3">
    
    <!-- Cabeçalho -->
    <div class="card shadow mb-4 header-bg">
        <div class="card-body">
            <div class="d-flex justify-content-between align-items-center">
                <div>
                    <h1 class="h3 mb-1"><i class="bi bi-bar-chart-line"></i> Relatório Detalhado de Metas</h1>
                    <p class="mb-0">Hierarquia completa: Empresa → Diretorias → Gerencias</p>
                </div>
                <div class="text-end">
                    <div class="d-flex gap-2">
                        <button class="btn btn-light btn-sm" onclick="window.print()">
                            <i class="bi bi-printer"></i> Imprimir
                        </button>
                        <a href="meta_gerenciamento.asp" class="btn btn-outline-light btn-sm">
                            <i class="bi bi-pencil-square"></i> Editar Metas
                        </a>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Filtro de Período -->
    <div class="card shadow mb-4 mes-selector">
        <div class="card-body">
            <h5 class="card-title mb-3"><i class="bi bi-filter"></i> Filtro de Período</h5>
            <form method="GET" class="row g-3">
                <div class="col-md-3">
                    <label class="form-label">Ano</label>
                    <select name="ano" class="form-select" id="anoSelect">
                        <%
                        ' Gerar opções de ano (2026 em diante)
                        Dim sqlAnos, rsAnos, anoAtual
                        sqlAnos = "SELECT DISTINCT Ano FROM MetaEmpresa WHERE Ano >= 2026 UNION " & _
                                  "SELECT DISTINCT Ano FROM MetaDiretoria WHERE Ano >= 2026 UNION " & _
                                  "SELECT DISTINCT Ano FROM MetaGerencia WHERE Ano >= 2026 " & _
                                  "ORDER BY Ano DESC"
                        Set rsAnos = connSales.Execute(sqlAnos)
                        
                        Do While Not rsAnos.EOF
                            anoAtual = rsAnos("Ano")
                            Response.Write "<option value='" & anoAtual & "'"
                            If CStr(anoAtual) = CStr(anoFiltro) Then Response.Write " selected"
                            Response.Write ">" & anoAtual & "</option>"
                            rsAnos.MoveNext
                        Loop
                        
                        If Not rsAnos Is Nothing Then
                            rsAnos.Close
                            Set rsAnos = Nothing
                        End If
                        %>
                    </select>
                </div>
                <div class="col-md-3">
                    <label class="form-label">Mês</label>
                    <select name="mes" class="form-select" id="mesSelect">
                        <option value="">Todos os meses</option>
                        <%
                        For i = 1 To 12
                            Response.Write "<option value='" & i & "'"
                            If CStr(i) = CStr(mesFiltro) Then Response.Write " selected"
                            Response.Write ">" & NomeMes(i) & "</option>"
                        Next
                        %>
                    </select>
                </div>
                <div class="col-md-6 d-flex align-items-end">
                    <div>
                        <button type="submit" class="btn btn-primary me-2">
                            <i class="bi bi-search"></i> Filtrar
                        </button>
                        <a href="meta_relatorio.asp" class="btn btn-outline-secondary">
                            <i class="bi bi-x-circle"></i> Limpar
                        </a>
                    </div>
                </div>
            </form>
        </div>
    </div>

    <!-- Meta da Empresa -->
    <div class="card shadow mb-4 card-empresa">
        <div class="card-header bg-primary text-white">
            <h4 class="mb-0">
                <i class="bi bi-building"></i> Meta da Empresa
                <% If mesFiltro <> "" Then %>
                <span class="badge bg-light text-dark float-end"><%= NomeMes(mesFiltro) & "/" & anoFiltro %></span>
                <% Else %>
                <span class="badge bg-light text-dark float-end">Ano: <%= anoFiltro %></span>
                <% End If %>
            </h4>
        </div>
        <div class="card-body">
            <div class="row">
                <div class="col-md-8">
                    <h5 class="text-muted">Meta Total da Empresa</h5>
                    <p class="text-muted mb-0">Soma consolidada de todas as diretorias</p>
                </div>
                <div class="col-md-4 text-end">
                    <h2 class="valor-meta">R$ <%= FxFormat(metaEmpresa) %></h2>
                    <% If metaEmpresa > 0 Then %>
                    <span class="badge bg-success"><i class="bi bi-check-circle"></i> Meta definida</span>
                    <% Else %>
                    <span class="badge bg-warning"><i class="bi bi-exclamation-triangle"></i> Sem meta</span>
                    <% End If %>
                </div>
            </div>
        </div>
    </div>

    <!-- Diretorias e Gerencias -->
    <%
    If Not rsDiretorias.EOF Then
    %>
    <div class="accordion mb-4" id="accordionDiretorias">
        <%
        Dim contadorDiretoria
        contadorDiretoria = 0
        
        Do While Not rsDiretorias.EOF
            Dim diretoriaID, nomeDiretoria, metaDiretoria, usuarioDiretoria, dataHoraDiretoria
            diretoriaID = rsDiretorias("DiretoriaID")
            nomeDiretoria = rsDiretorias("NomeDiretoria")
            metaDiretoria = ValorSafe(rsDiretorias("TotalMetas"))
            usuarioDiretoria = rsDiretorias("Usuario")
            If IsNull(usuarioDiretoria) Then usuarioDiretoria = ""
            dataHoraDiretoria = rsDiretorias("DataHora")
            
            ' Somar ao total das diretorias
            totalDiretorias = totalDiretorias + metaDiretoria
            
            contadorDiretoria = contadorDiretoria + 1
        %>
        <div class="card shadow mb-3 card-diretoria">
            <div class="card-header" id="heading<%= contadorDiretoria %>">
                <div class="d-flex justify-content-between align-items-center">
                    <div>
                        <button class="btn btn-link text-decoration-none" type="button" 
                                data-bs-toggle="collapse" 
                                data-bs-target="#collapse<%= contadorDiretoria %>" 
                                aria-expanded="true" 
                                aria-controls="collapse<%= contadorDiretoria %>">
                            <h5 class="mb-0">
                                <i class="bi bi-chevron-down hierarchy-icon"></i>
                                <span class="badge badge-diretoria me-2">Diretoria</span>
                                <%= nomeDiretoria %>
                            </h5>
                        </button>
                    </div>
                    <div class="text-end">
                        <h4 class="mb-0 valor-meta">R$ <%= FxFormat(metaDiretoria) %></h4>
                        <small class="text-muted">
                            <% If usuarioDiretoria <> "" Then %>
                            <i class="bi bi-person"></i> <%= usuarioDiretoria %>
                            <% If Not IsNull(dataHoraDiretoria) Then %>
                            | <i class="bi bi-clock"></i> <%= FormatDateTime(dataHoraDiretoria, 2) %>
                            <% End If %>
                            <% End If %>
                        </small>
                    </div>
                </div>
            </div>

            <div id="collapse<%= contadorDiretoria %>" class="collapse show" 
                 aria-labelledby="heading<%= contadorDiretoria %>" 
                 data-bs-parent="#accordionDiretorias">
                <div class="card-body">
                    
                    <!-- Buscar gerencias desta diretoria -->
                    <%
                    Dim sqlGerencias, rsGerencias
                    sqlGerencias = "SELECT " & _
                                   "g.GerenciaID, " & _
                                   "g.NomeGerencia, " & _
                                   "mg.ValorMeta, " & _
                                   "mg.Usuario, " & _
                                   "mg.DataHora " & _
                                   "FROM Gerencias g " & _
                                   "LEFT JOIN MetaGerencia mg ON g.GerenciaID = mg.GerenciaID " & _
                                   "AND mg.Ano = " & anoFiltro & " AND mg.Mes = " & mesFiltro & " " & _
                                   "WHERE g.DiretoriaID = " & diretoriaID & " " & _
                                   "ORDER BY g.NomeGerencia"
                    
                    Set rsGerencias = connSales.Execute(sqlGerencias)
                    
                    Dim temGerencia, totalGerenciaD
                    temGerencia = False
                    totalGerenciaD = 0
                    %>
                    
                    <div class="table-responsive">
                        <table class="table table-sm table-hover">
                            <thead class="table-light">
                                <tr>
                                    <th width="60%">Gerência</th>
                                    <th width="25%" class="text-end">Meta</th>
                                    <th width="15%" class="text-center">Status</th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                Do While Not rsGerencias.EOF
                                    Dim gerenciaID, nomeGerencia, metaGerencia, usuarioGerencia, dataHoraGerencia
                                    gerenciaID = rsGerencias("GerenciaID")
                                    nomeGerencia = rsGerencias("NomeGerencia")
                                    metaGerencia = ValorSafe(rsGerencias("ValorMeta"))
                                    usuarioGerencia = rsGerencias("Usuario")
                                    If IsNull(usuarioGerencia) Then usuarioGerencia = ""
                                    dataHoraGerencia = rsGerencias("DataHora")
                                    
                                    ' Somar ao total das gerencias e da diretoria
                                    totalGerencias = totalGerencias + metaGerencia
                                    totalGerenciaD = totalGerenciaD + metaGerencia
                                    
                                    temGerencia = True
                                %>
                                <tr>
                                    <td>
                                        <i class="bi bi-diagram-2 text-muted me-2"></i>
                                        <%= nomeGerencia %>
                                        <% If usuarioGerencia <> "" Then %>
                                        <br><small class="text-muted">
                                            <i class="bi bi-person"></i> <%= usuarioGerencia %>
                                            <% If Not IsNull(dataHoraGerencia) Then %>
                                            | <i class="bi bi-clock"></i> <%= FormatDateTime(dataHoraGerencia, 2) %>
                                            <% End If %>
                                        </small>
                                        <% End If %>
                                    </td>
                                    <td class="text-end">
                                        <% If metaGerencia > 0 Then %>
                                        <span class="valor-meta">R$ <%= FxFormat(metaGerencia) %></span>
                                        <% Else %>
                                        <span class="valor-zero">R$ 0,00</span>
                                        <% End If %>
                                    </td>
                                    <td class="text-center">
                                        <% If metaGerencia > 0 Then %>
                                        <span class="badge bg-success">Definida</span>
                                        <% Else %>
                                        <span class="badge bg-secondary">Não definida</span>
                                        <% End If %>
                                    </td>
                                </tr>
                                <%
                                    rsGerencias.MoveNext
                                Loop
                                
                                If Not rsGerencias Is Nothing Then
                                    rsGerencias.Close
                                    Set rsGerencias = Nothing
                                End If
                                
                                ' Se não tem gerencias, mostrar mensagem
                                If Not temGerencia Then
                                %>
                                <tr>
                                    <td colspan="3" class="text-center text-muted">
                                        <i class="bi bi-info-circle"></i> Nenhuma gerência cadastrada para esta diretoria.
                                    </td>
                                </tr>
                                <%
                                End If
                                %>
                            </tbody>
                            <tfoot class="table-light">
                                <tr>
                                    <th class="text-end">Total da Diretoria (Soma Gerencias):</th>
                                    <th class="text-end">R$ <%= FxFormat(totalGerenciaD) %></th>
                                    <th class="text-center">
                                        <% 
                                        If totalGerenciaD = metaDiretoria Then 
                                            Response.Write "<span class='badge bg-success'><i class='bi bi-check-circle'></i> Consistente</span>"
                                        ElseIf metaDiretoria > 0 Then
                                            Response.Write "<span class='badge bg-warning'><i class='bi bi-exclamation-triangle'></i> Diferença</span>"
                                        End If
                                        %>
                                    </th>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                    
                </div>
            </div>
        </div>
        <%
            rsDiretorias.MoveNext
        Loop
        %>
    </div>
    <%
    Else
    ' Se não tem diretorias
    %>
    <div class="card shadow">
        <div class="card-body text-center py-5">
            <i class="bi bi-building text-muted" style="font-size: 3rem;"></i>
            <h4 class="mt-3">Nenhuma diretoria encontrada</h4>
            <p class="text-muted">Não há diretorias cadastradas ou não existem metas para o período selecionado.</p>
        </div>
    </div>
    <%
    End If
    
    If Not rsDiretorias Is Nothing Then
        rsDiretorias.Close
        Set rsDiretorias = Nothing
    End If
    %>

    <!-- Resumo Total (apenas se houver dados) -->
    <% If totalDiretorias > 0 Or totalGerencias > 0 Or metaEmpresa > 0 Then %>
    <div class="row mb-4">
        <div class="col-md-4">
            <div class="card shadow total-box">
                <div class="card-body text-center">
                    <h6 class="card-subtitle mb-2">Total Gerencias</h6>
                    <h3 class="card-title">R$ <%= FxFormat(totalGerencias) %></h3>
                    <p class="card-text small">Soma de todas as gerencias</p>
                </div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="card shadow total-box" style="background-color: #27ae60;">
                <div class="card-body text-center">
                    <h6 class="card-subtitle mb-2">Total Diretorias</h6>
                    <h3 class="card-title">R$ <%= FxFormat(totalDiretorias) %></h3>
                    <p class="card-text small">Soma de todas as diretorias</p>
                </div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="card shadow total-box" style="background-color: #2980b9;">
                <div class="card-body text-center">
                    <h6 class="card-subtitle mb-2">Meta Empresa</h6>
                    <h3 class="card-title">R$ <%= FxFormat(metaEmpresa) %></h3>
                    <p class="card-text small">Meta consolidada da empresa</p>
                </div>
            </div>
        </div>
    </div>

    <!-- Resumo de Consistência -->
    <div class="card shadow">
        <div class="card-header bg-light">
            <h5 class="mb-0"><i class="bi bi-clipboard-check"></i> Resumo de Consistência</h5>
        </div>
        <div class="card-body">
            <div class="row">
                <div class="col-md-6">
                    <h6>Verificação de Totais:</h6>
                    <ul class="list-group list-group-flush">
                        <li class="list-group-item d-flex justify-content-between align-items-center">
                            Soma Gerencias vs Soma Diretorias
                            <span class="badge <% If totalGerencias = totalDiretorias Then %>bg-success<% Else %>bg-warning<% End If %> rounded-pill">
                                <% If totalGerencias = totalDiretorias Then %>
                                <i class="bi bi-check-circle"></i> OK
                                <% Else %>
                                <i class="bi bi-exclamation-triangle"></i> Diferença
                                <% End If %>
                            </span>
                        </li>
                        <li class="list-group-item d-flex justify-content-between align-items-center">
                            Soma Diretorias vs Meta Empresa
                            <span class="badge <% If totalDiretorias = metaEmpresa Then %>bg-success<% Else %>bg-warning<% End If %> rounded-pill">
                                <% If totalDiretorias = metaEmpresa Then %>
                                <i class="bi bi-check-circle"></i> OK
                                <% Else %>
                                <i class="bi bi-exclamation-triangle"></i> Diferença
                                <% End If %>
                            </span>
                        </li>
                    </ul>
                </div>
                <div class="col-md-6">
                    <h6>Diferenças Detalhadas:</h6>
                    <table class="table table-sm">
                        <tr>
                            <td>Soma Gerencias:</td>
                            <td class="text-end">R$ <%= FxFormat(totalGerencias) %></td>
                        </tr>
                        <tr>
                            <td>Soma Diretorias:</td>
                            <td class="text-end">R$ <%= FxFormat(totalDiretorias) %></td>
                        </tr>
                        <tr>
                            <td>Meta Empresa:</td>
                            <td class="text-end">R$ <%= FxFormat(metaEmpresa) %></td>
                        </tr>
                        <tr class="table-warning">
                            <td><strong>Diferença Gerencias vs Diretorias:</strong></td>
                            <td class="text-end"><strong>R$ <%= FxFormat(totalDiretorias - totalGerencias) %></strong></td>
                        </tr>
                        <tr class="table-warning">
                            <td><strong>Diferença Diretorias vs Empresa:</strong></td>
                            <td class="text-end"><strong>R$ <%= FxFormat(metaEmpresa - totalDiretorias) %></strong></td>
                        </tr>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <% End If %>

</div> <!-- container -->

<!-- Bootstrap JS -->
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>

<script>
// Auto-expand first accordion item
document.addEventListener('DOMContentLoaded', function() {
    // Expandir primeiro item do accordion
    var firstAccordion = document.querySelector('.accordion .collapse');
    if (firstAccordion) {
        var bsCollapse = new bootstrap.Collapse(firstAccordion, {
            toggle: true
        });
    }
});
</script>

</body>
</html>

<%
' Fechar conexões
If Not connSales Is Nothing Then
    connSales.Close
    Set connSales = Nothing
End If

If Not connDG Is Nothing Then
    connDG.Close
    Set connDG = Nothing
End If
%>