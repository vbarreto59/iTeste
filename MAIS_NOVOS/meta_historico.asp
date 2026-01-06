<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 18/12/2025               -->
<!-- CODIGO_ARQUIVO: ZCCCFCCZHB          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="conexao.asp"-->           
<!--#include file="conSunSales.asp"-->       
<!--#include file="usr_acoes_v4GVendas.inc"-->

<%
Response.CodePage = 65001
Response.Charset = "utf-8"

' -----------------------------
'  ABRIR CONEXÕES
' -----------------------------
Dim connOrg, connSales
Set connOrg = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")

connOrg.Open strConn
connSales.Open strConnSales

' -----------------------------
'  FUNÇÕES AUXILIARES
' -----------------------------
Function ArrayMin(arr)
    If Not IsArray(arr) Or UBound(arr) < 0 Then
        ArrayMin = 0
        Exit Function
    End If
    
    Dim i, result
    result = arr(0)
    
    For i = 1 To UBound(arr)
        If arr(i) < result Then
            result = arr(i)
        End If
    Next
    
    ArrayMin = result
End Function

Function ArrayMax(arr)
    If Not IsArray(arr) Or UBound(arr) < 0 Then
        ArrayMax = 0
        Exit Function
    End If
    
    Dim i, result
    result = arr(0)
    
    For i = 1 To UBound(arr)
        If arr(i) > result Then
            result = arr(i)
        End If
    Next
    
    ArrayMax = result
End Function

' -----------------------------
'  RECEBER ANO (OPCIONAL)
' -----------------------------
Dim anoSelecionado, anoMinimo, anoMaximo
anoSelecionado = Request("ano")

If anoSelecionado = "" Then
    anoSelecionado = Year(Now)
End If

' -----------------------------
'  OBTER ANOS DISPONÍVEIS
' -----------------------------
Dim sql, rsAnos, anosArray
sql = "SELECT DISTINCT Ano FROM MetaGerencia WHERE Ano >= 2026 UNION SELECT DISTINCT Ano FROM MetaDiretoria WHERE Ano >= 2026 ORDER BY Ano DESC"
Set rsAnos = connSales.Execute(sql)

ReDim anosArray(0)
Do While Not rsAnos.EOF
    anosArray(UBound(anosArray)) = CInt(rsAnos("Ano"))
    ReDim Preserve anosArray(UBound(anosArray) + 1)
    rsAnos.MoveNext
Loop
rsAnos.Close
Set rsAnos = Nothing

If UBound(anosArray) > 0 Then
    ReDim Preserve anosArray(UBound(anosArray) - 1)
    anoMinimo = ArrayMin(anosArray)
    anoMaximo = ArrayMax(anosArray)
Else
    anoMinimo = Year(Now)
    anoMaximo = Year(Now)
End If

' -----------------------------
'  OBTER METAS DIRETORIAS (MAIS SIMPLES)
' -----------------------------
Dim historicoDiretorias
Set historicoDiretorias = Server.CreateObject("Scripting.Dictionary")

Dim rsDiretorias, sqlDiretorias
sqlDiretorias = "SELECT " & _
      "md.DiretoriaID, " & _
      "md.Ano, " & _
      "md.Mes, " & _
      "md.TotalMetas, " & _
      "md.Usuario, " & _
      "md.DataHora " & _
      "FROM MetaDiretoria md " & _
      "WHERE md.Ano >= 2026 " & _
      "ORDER BY md.DiretoriaID, Ano, Mes"

Set rsDiretorias = connSales.Execute(sqlDiretorias)

Do While Not rsDiretorias.EOF
    Dim chave, info, rsDirNome, sqlDirNome
    chave = CStr(rsDiretorias("DiretoriaID")) & "|" & CStr(rsDiretorias("Ano")) & "|" & CStr(rsDiretorias("Mes"))
    
    ' Buscar nome da diretoria no banco connOrg
    sqlDirNome = "SELECT NomeDiretoria FROM Diretorias WHERE DiretoriaID = " & rsDiretorias("DiretoriaID")
    Set rsDirNome = connOrg.Execute(sqlDirNome)
    
    If Not rsDirNome.EOF Then
        Dim dirNome
        dirNome = "" & rsDirNome("NomeDiretoria")
        
        info = dirNome & "|" & rsDiretorias("Ano") & "|" & rsDiretorias("Mes") & "|" & _
               rsDiretorias("TotalMetas") & "|" & rsDiretorias("Usuario") & "|" & rsDiretorias("DataHora")
        
        If Not historicoDiretorias.Exists(chave) Then
            historicoDiretorias.Add chave, info
        End If
    End If
    
    rsDirNome.Close
    Set rsDirNome = Nothing
    
    rsDiretorias.MoveNext
Loop
rsDiretorias.Close
Set rsDiretorias = Nothing

' -----------------------------
'  OBTER METAS DA EMPRESA
' -----------------------------
Dim historicoEmpresa
Set historicoEmpresa = Server.CreateObject("Scripting.Dictionary")

Dim rsEmpresa, sqlEmpresa
sqlEmpresa = "SELECT " & _
      "Ano, " & _
      "Mes, " & _
      "Meta, " & _
      "DataHora " & _
      "FROM MetaEmpresa " & _
      "WHERE Ano >= 2026 " & _
      "ORDER BY DataHora DESC"

Set rsEmpresa = connSales.Execute(sqlEmpresa)

Do While Not rsEmpresa.EOF
    chave = CStr(rsEmpresa("Ano")) & "|" & CStr(rsEmpresa("Mes"))
    
    info = rsEmpresa("Ano") & "|" & rsEmpresa("Mes") & "|" & rsEmpresa("Meta") & "|" & rsEmpresa("DataHora")
    
    If Not historicoEmpresa.Exists(chave) Then
        historicoEmpresa.Add chave, info
    End If
    
    rsEmpresa.MoveNext
Loop
rsEmpresa.Close
Set rsEmpresa = Nothing

' Definir meses nomes globalmente
Dim mesesNomes
mesesNomes = Array("", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez")
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Histórico de Metas</title>
    <!-- Bootstrap 5 -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css">
    <style>
        :root {
            --primary-color: #4361ee;
            --secondary-color: #3a0ca3;
            --success-color: #2ec4b6;
        }
        .card-header-custom {
            background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
            color: white;
        }
        .table-hover tbody tr:hover {
            background-color: rgba(67, 97, 238, 0.05);
        }
        .badge-ano {
            background-color: #6c757d;
            color: white;
        }
        .badge-mes {
            background-color: #0dcaf0;
            color: white;
        }
        .valor-positivo {
            color: #198754;
            font-weight: bold;
        }
        .valor-zero {
            color: #6c757d;
            font-style: italic;
        }
        .usuario-info {
            font-size: 0.85rem;
            color: #6c757d;
        }
        .data-info {
            font-size: 0.8rem;
            color: #adb5bd;
        }
    </style>
</head>
<body>
    <div class="container-fluid py-4">
        <!-- Cabeçalho -->
        <div class="row mb-4">
            <div class="col-12">
                <div class="card shadow">
                    <div class="card-header card-header-custom">
                        <div class="d-flex justify-content-between align-items-center">
                            <div>
                                <h4 class="mb-0"><i class="bi bi-clock-history me-2"></i>Histórico de Metas</h4>
                                <small>Registro de alterações nas metas organizacionais</small>
                            </div>
                            <div>
                                <span class="badge bg-light text-dark">
                                    <i class="bi bi-calendar-range me-1"></i>
                                    <%=anoMinimo%> a <%=anoMaximo%>
                                </span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Filtro de Ano -->
        <div class="row mb-4">
            <div class="col-12">
                <div class="card shadow">
                    <div class="card-body">
                        <h5 class="card-title mb-3">
                            <i class="bi bi-funnel me-2"></i>Filtrar por Ano
                        </h5>
                        <form method="GET" class="row g-3">
                            <div class="col-md-8">
                                <label class="form-label">Selecionar Ano</label>
                                <select name="ano" class="form-select" onchange="this.form.submit()">
                                    <option value="">Todos os anos</option>
                                    <%
                                    For ano = anoMaximo To anoMinimo Step -1
                                        Response.Write "<option value=""" & ano & """"
                                        If CStr(ano) = CStr(anoSelecionado) Then
                                            Response.Write " selected"
                                        End If
                                        Response.Write ">" & ano & "</option>"
                                    Next
                                    
                                    ' Adicionar anos padrão caso não haja dados
                                    If UBound(anosArray) < 0 Then
                                        For ano = Year(Now) To Year(Now) - 2 Step -1
                                            Response.Write "<option value=""" & ano & """"
                                            If CStr(ano) = CStr(anoSelecionado) Then
                                                Response.Write " selected"
                                            End If
                                            Response.Write ">" & ano & "</option>"
                                        Next
                                    End If
                                    %>
                                </select>
                            </div>
                            <div class="col-md-4 d-flex align-items-end">
                                <div class="d-grid gap-2 d-md-flex w-100">
                                    <a href="meta_gerenciamento.asp" class="btn btn-outline-primary">
                                        <i class="bi bi-pencil-square me-1"></i>Gerenciar
                                    </a>
                                    <a href="meta_resumo_anual.asp" class="btn btn-outline-success">
                                        <i class="bi bi-calendar-check me-1"></i>Resumo
                                    </a>
                                </div>
                            </div>
                        </form>
                    </div>
                </div>
            </div>
        </div>

        <!-- Histórico de Metas por Diretoria -->
        <div class="row mb-4">
            <div class="col-12">
                <div class="card shadow">
                    <div class="card-header bg-light">
                        <h5 class="mb-0">
                            <i class="bi bi-diagram-3 me-2"></i>
                            Histórico por Diretoria
                            <span class="badge bg-secondary ms-2"><%=historicoDiretorias.Count%></span>
                        </h5>
                    </div>
                    <div class="card-body">
                        <%
                        If historicoDiretorias.Count > 0 Then
                        %>
                        <div class="table-responsive">
                            <table class="table table-hover">
                                <thead class="table-light">
                                    <tr>
                                        <th>Diretoria</th>
                                        <th>Período</th>
                                        <th class="text-end">Total de Metas</th>
                                        <th>Alterado por</th>
                                        <th>Data/Hora</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    Dim dirKeys, dirKey, dirInfo, dirParts
                                    dirKeys = historicoDiretorias.Keys
                                    
                                    For Each dirKey In dirKeys
                                        If historicoDiretorias.Exists(dirKey) Then
                                            dirInfo = historicoDiretorias.Item(dirKey)
                                            
                                            ' Separar os dados pelo separador |
                                            dirParts = Split(dirInfo, "|")
                                            
                                            If UBound(dirParts) >= 5 Then
                                                'Dim dirNome, dirAno, dirMes, dirValor, dirUsuario, dirData
                                                dirNome = dirParts(0)
                                                dirAno = dirParts(1)
                                                dirMes = dirParts(2)
                                                dirValor = CDbl(dirParts(3))
                                                dirUsuario = dirParts(4)
                                                dirData = dirParts(5)
                                                
                                                ' Aplicar filtro de ano
                                                If anoSelecionado = "" Or CStr(dirAno) = CStr(anoSelecionado) Then
                                                    %>
                                                    <tr>
                                                        <td>
                                                            <i class="bi bi-building text-primary me-2"></i>
                                                            <%=dirNome%>
                                                        </td>
                                                        <td>
                                                            <span class="badge badge-ano me-1"><%=dirAno%></span>
                                                            <span class="badge badge-mes"><%=mesesNomes(CInt(dirMes))%></span>
                                                        </td>
                                                        <td class="text-end">
                                                            <%
                                                            If dirValor > 0 Then
                                                                Response.Write "<span class='valor-positivo'>R$ " & FormatNumber(dirValor, 2, -1, -1, -1) & "</span>"
                                                            Else
                                                                Response.Write "<span class='valor-zero'>R$ " & FormatNumber(dirValor, 2, -1, -1, -1) & "</span>"
                                                            End If
                                                            %>
                                                        </td>
                                                        <td>
                                                            <span class="usuario-info">
                                                                <i class="bi bi-person-circle me-1"></i>
                                                                <%=dirUsuario%>
                                                            </span>
                                                        </td>
                                                        <td>
                                                            <span class="data-info">
                                                                <%'=FormatDateTime(dirData, 2)%>
                                                                <br>
                                                                <small><%'=FormatDateTime(dirData, 3)%></small>
                                                            </span>
                                                        </td>
                                                    </tr>
                                                    <%
                                                End If
                                            End If
                                        End If
                                    Next
                                    %>
                                </tbody>
                            </table>
                        </div>
                        <%
                        Else
                        %>
                        <div class="alert alert-warning">
                            <i class="bi bi-exclamation-triangle me-2"></i>
                            Nenhum histórico de metas de diretorias encontrado.
                        </div>
                        <%
                        End If
                        %>
                    </div>
                </div>
            </div>
        </div>

        <!-- Histórico de Metas da Empresa -->
        <div class="row">
            <div class="col-12">
                <div class="card shadow">
                    <div class="card-header bg-light">
                        <h5 class="mb-0">
                            <i class="bi bi-building me-2"></i>
                            Histórico da Empresa
                            <span class="badge bg-secondary ms-2"><%=historicoEmpresa.Count%></span>
                        </h5>
                    </div>
                    <div class="card-body">
                        <%
                        If historicoEmpresa.Count > 0 Then
                        %>
                        <div class="table-responsive">
                            <table class="table table-hover">
                                <thead class="table-light">
                                    <tr>
                                        <th>Período</th>
                                        <th class="text-end">Meta da Empresa</th>
                                        <th>Data/Hora</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    Dim empKeys, empKey, empInfo, empParts
                                    empKeys = historicoEmpresa.Keys
                                    
                                    For Each empKey In empKeys
                                        If historicoEmpresa.Exists(empKey) Then
                                            empInfo = historicoEmpresa.Item(empKey)
                                            
                                            ' Separar os dados pelo separador |
                                            empParts = Split(empInfo, "|")
                                            
                                            If UBound(empParts) >= 3 Then
                                                Dim empAno, empMes, empValor, empData
                                                empAno = empParts(0)
                                                empMes = empParts(1)
                                                empValor = CDbl(empParts(2))
                                                empData = empParts(3)
                                                
                                                ' Aplicar filtro de ano
                                                If anoSelecionado = "" Or CStr(empAno) = CStr(anoSelecionado) Then
                                                    %>
                                                    <tr>
                                                        <td>
                                                            <span class="badge badge-ano me-1"><%=empAno%></span>
                                                            <span class="badge badge-mes"><%=mesesNomes(CInt(empMes))%></span>
                                                        </td>
                                                        <td class="text-end">
                                                            <%
                                                            If empValor > 0 Then
                                                                Response.Write "<span class='valor-positivo'>R$ " & FormatNumber(empValor, 2, -1, -1, -1) & "</span>"
                                                            Else
                                                                Response.Write "<span class='valor-zero'>R$ " & FormatNumber(empValor, 2, -1, -1, -1) & "</span>"
                                                            End If
                                                            %>
                                                        </td>
                                                        <td>
                                                            <span class="data-info">
                                                                <%=FormatDateTime(empData, 2)%>
                                                                <br>
                                                                <small><%=FormatDateTime(empData, 3)%></small>
                                                            </span>
                                                        </td>
                                                    </tr>
                                                    <%
                                                End If
                                            End If
                                        End If
                                    Next
                                    %>
                                </tbody>
                            </table>
                        </div>
                        <%
                        Else
                        %>
                        <div class="alert alert-warning">
                            <i class="bi bi-exclamation-triangle me-2"></i>
                            Nenhum histórico de metas da empresa encontrado.
                        </div>
                        <%
                        End If
                        %>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    
    <script>
        // Inicializar tooltips
        document.addEventListener('DOMContentLoaded', function() {
            var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'))
            var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
                return new bootstrap.Tooltip(tooltipTriggerEl)
            });
        });
    </script>
</body>
</html>
<%
' Fechar conexões
If IsObject(connOrg) Then
    connOrg.Close
    Set connOrg = Nothing
End If

If IsObject(connSales) Then
    connSales.Close
    Set connSales = Nothing
End If
%>