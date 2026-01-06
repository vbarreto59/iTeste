<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 18/12/2025               -->
<!-- CODIGO_ARQUIVO: UFENJRTTQM          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<!--#include file="usr_acoes_v4GVendas.inc"-->

<%
' metas_pivot.asp - versão completa com DataTables e totais dinâmicos no rodapé
Response.CodePage = 65001
Response.Charset = "utf-8"

' -----------------------------
' ABRIR CONEXÕES
' -----------------------------
Dim connOrg, connSales
Set connOrg = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")
connOrg.Open StrConn
connSales.Open StrConnSales

' -----------------------------
' FUNÇÕES AUXILIARES
' -----------------------------
Function SafeInt(v, defaultVal)
    On Error Resume Next
    If IsNumeric(v) Then
        SafeInt = CInt(v)
    Else
        SafeInt = defaultVal
    End If
    Err.Clear
    On Error GoTo 0
End Function

Function SafeString(v)
    If IsNull(v) Then
        SafeString = ""
    Else
        SafeString = Trim(CStr(v))
    End If
End Function

Function FormatMoneyBR(v)
    On Error Resume Next
    If IsNumeric(v) Then
        FormatMoneyBR = "R$ " & FormatNumber(CDbl(v), 2, -1, -1, -1)
    Else
        FormatMoneyBR = "R$ 0,00"
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Nova função para extrair valor numérico
Function GetNumericValue(v)
    On Error Resume Next
    If IsNumeric(v) Then
        GetNumericValue = CDbl(v)
    Else
        ' Tenta extrair número de string como "R$ 1.000,00"
        Dim strVal
        strVal = CStr(v)
        strVal = Replace(strVal, "R$", "")
        strVal = Replace(strVal, ".", "")
        strVal = Replace(strVal, ",", ".")
        strVal = Trim(strVal)
        If IsNumeric(strVal) Then
            GetNumericValue = CDbl(strVal)
        Else
            GetNumericValue = 0
        End If
    End If
    If Err.Number <> 0 Then GetNumericValue = 0
    Err.Clear
    On Error GoTo 0
End Function

' -----------------------------
' RECEBER FILTROS
' -----------------------------
Dim anoSelecionado, mesSelecionado
anoSelecionado = Request("ano")
mesSelecionado = Request("mes") ' 1..12 ou "" para todos

If anoSelecionado = "" Then anoSelecionado = ""
If mesSelecionado = "" Then mesSelecionado = ""

If anoSelecionado <> "" And Not IsNumeric(anoSelecionado) Then anoSelecionado = ""
If mesSelecionado <> "" And Not IsNumeric(mesSelecionado) Then mesSelecionado = ""

If anoSelecionado <> "" Then anoSelecionado = CInt(anoSelecionado)
If mesSelecionado <> "" Then mesSelecionado = CInt(mesSelecionado)

' -----------------------------
' MESES
' -----------------------------
Dim mesesNomes
mesesNomes = Array("", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez")

' -----------------------------
' OBTER ANOS DISPONÍVEIS (MetaDiretoria + MetaGerencia)
' -----------------------------
Dim rsAnos, sqlAnos, anosDict, anoKey
Set anosDict = Server.CreateObject("Scripting.Dictionary")
sqlAnos = "SELECT DISTINCT Ano FROM (SELECT Ano FROM MetaDiretoria UNION SELECT Ano FROM MetaGerencia) AS t ORDER BY Ano DESC"
Set rsAnos = connSales.Execute(sqlAnos)
Do While Not rsAnos.EOF
    anoKey = SafeInt(rsAnos("Ano"), 0)
    If anoKey > 0 Then
        If Not anosDict.Exists(anoKey) Then anosDict.Add anoKey, anoKey
    End If
    rsAnos.MoveNext
Loop
rsAnos.Close
Set rsAnos = Nothing

Dim anosArray()
If anosDict.Count > 0 Then
    ReDim anosArray(anosDict.Count - 1)
    Dim i
    i = 0
    For Each anoKey In anosDict.Keys
        anosArray(i) = anoKey
        i = i + 1
    Next
Else
    ReDim anosArray(2)
    anosArray(0) = Year(Date())
    anosArray(1) = Year(Date()) - 1
    anosArray(2) = Year(Date()) - 2
End If

' -----------------------------
' CARREGAR DADOS: DIRETORIAS, GERENCIAS, EMPRESA
' Armazenamos em arrays (cada item: Nome|Ano|Mes|Valor|Usuario|DataHora)
' -----------------------------
Dim dictDiretorias, dictGerencias, dictEmpresa
Set dictDiretorias = Server.CreateObject("Scripting.Dictionary")
Set dictGerencias = Server.CreateObject("Scripting.Dictionary")
Set dictEmpresa = Server.CreateObject("Scripting.Dictionary")

' -- MetaDiretoria
Dim rsDir, sqlDir
sqlDir = "SELECT DiretoriaID, Ano, Mes, TotalMetas, Usuario, DataHora FROM MetaDiretoria WHERE 1=1"
If anoSelecionado <> "" Then sqlDir = sqlDir & " AND Ano = " & CStr(anoSelecionado)
If mesSelecionado <> "" Then sqlDir = sqlDir & " AND Mes = " & CStr(mesSelecionado)
sqlDir = sqlDir & " ORDER BY Ano DESC, Mes DESC"
Set rsDir = connSales.Execute(sqlDir)

Do While Not rsDir.EOF
    Dim dID, dAno, dMes, dValor, dUsuario, dData, dNome, chaveD
    dID = SafeInt(rsDir("DiretoriaID"), 0)
    dAno = SafeInt(rsDir("Ano"), 0)
    dMes = SafeInt(rsDir("Mes"), 0)
    dValor = rsDir("TotalMetas")
    dUsuario = SafeString(rsDir("Usuario"))
    dData = SafeString(rsDir("DataHora"))

    ' buscar nome da diretoria
    Dim rsDN, sqlDN
    sqlDN = "SELECT NomeDiretoria FROM Diretorias WHERE DiretoriaID = " & CStr(dID)
    Set rsDN = connOrg.Execute(sqlDN)
    If Not rsDN.EOF Then
        dNome = SafeString(rsDN("NomeDiretoria"))
    Else
        dNome = "Diretoria " & CStr(dID)
    End If
    rsDN.Close
    Set rsDN = Nothing

    chaveD = "D|" & CStr(dID) & "|" & CStr(dAno) & "|" & CStr(dMes)
    If Not dictDiretorias.Exists(chaveD) Then
        dictDiretorias.Add chaveD, dNome & "|" & dAno & "|" & dMes & "|" & dValor & "|" & dUsuario & "|" & dData
    End If

    rsDir.MoveNext
Loop
rsDir.Close
Set rsDir = Nothing

' -- MetaGerencia
Dim rsGer, sqlGer
sqlGer = "SELECT GerenciaID, Ano, Mes, ValorMeta, Usuario, DataHora FROM MetaGerencia WHERE 1=1"
If anoSelecionado <> "" Then sqlGer = sqlGer & " AND Ano = " & CStr(anoSelecionado)
If mesSelecionado <> "" Then sqlGer = sqlGer & " AND Mes = " & CStr(mesSelecionado)
sqlGer = sqlGer & " ORDER BY Ano DESC, Mes DESC"
Set rsGer = connSales.Execute(sqlGer)

Do While Not rsGer.EOF
    Dim gID, gAno, gMes, gValor, gUsuario, gData, gNome, chaveG
    gID = SafeInt(rsGer("GerenciaID"), 0)
    gAno = SafeInt(rsGer("Ano"), 0)
    gMes = SafeInt(rsGer("Mes"), 0)
    gValor = rsGer("ValorMeta")
    gUsuario = SafeString(rsGer("Usuario"))
    gData = SafeString(rsGer("DataHora"))

    ' buscar nome da gerencia
    Dim rsGN, sqlGN
    sqlGN = "SELECT NomeGerencia FROM Gerencias WHERE GerenciaID = " & CStr(gID)
    Set rsGN = connOrg.Execute(sqlGN)
    If Not rsGN.EOF Then
        gNome = SafeString(rsGN("NomeGerencia"))
    Else
        gNome = "Gerência " & CStr(gID)
    End If
    rsGN.Close
    Set rsGN = Nothing

    chaveG = "G|" & CStr(gID) & "|" & CStr(gAno) & "|" & CStr(gMes)
    If Not dictGerencias.Exists(chaveG) Then
        dictGerencias.Add chaveG, gNome & "|" & gAno & "|" & gMes & "|" & gValor & "|" & gUsuario & "|" & gData
    End If

    rsGer.MoveNext
Loop
rsGer.Close
Set rsGer = Nothing

' -- MetaEmpresa
Dim rsEmp, sqlEmp
sqlEmp = "SELECT Ano, Mes, Meta AS TotalMetas, DataHora FROM MetaEmpresa WHERE 1=1"
If anoSelecionado <> "" Then sqlEmp = sqlEmp & " AND Ano = " & CStr(anoSelecionado)
If mesSelecionado <> "" Then sqlEmp = sqlEmp & " AND Mes = " & CStr(mesSelecionado)
sqlEmp = sqlEmp & " ORDER BY Ano DESC, Mes DESC"
Set rsEmp = connSales.Execute(sqlEmp)

Do While Not rsEmp.EOF
    Dim eAno, eMes, eValor, eData, chaveE
    eAno = SafeInt(rsEmp("Ano"), 0)
    eMes = SafeInt(rsEmp("Mes"), 0)
    eValor = rsEmp("TotalMetas")
    eData = SafeString(rsEmp("DataHora"))

    chaveE = "E|0|" & CStr(eAno) & "|" & CStr(eMes)
    If Not dictEmpresa.Exists(chaveE) Then
        dictEmpresa.Add chaveE, "Empresa|" & eAno & "|" & eMes & "|" & eValor & "||" & eData
    End If

    rsEmp.MoveNext
Loop
rsEmp.Close
Set rsEmp = Nothing

' -----------------------------
' Transformar dicionários em arrays
' -----------------------------
Function DictToArray(dict)
    Dim arr(), k, idx
    If dict.Count = 0 Then
        ReDim arr(-1)
        DictToArray = arr
        Exit Function
    End If
    ReDim arr(dict.Count - 1)
    idx = 0
    For Each k In dict.Keys
        arr(idx) = dict.Item(k)
        idx = idx + 1
    Next
    DictToArray = arr
End Function

Dim arrDiret, arrGer, arrEmp
arrDiret = DictToArray(dictDiretorias)
arrGer = DictToArray(dictGerencias)
arrEmp = DictToArray(dictEmpresa)

' -----------------------------
' Ordenar arrays por Ano DESC, Mes DESC (bubble simples)
' -----------------------------
Sub SortByAnoMesDesc(ByRef a)
    Dim swapped, i, tmp, p1, p2, a1, a2, m1, m2
    If UBound(a) < 0 Then Exit Sub
    Do
        swapped = False
        For i = 0 To UBound(a) - 1
            p1 = Split(a(i), "|")
            p2 = Split(a(i+1), "|")
            a1 = SafeInt(p1(1), 0) : m1 = SafeInt(p1(2), 0)
            a2 = SafeInt(p2(1), 0) : m2 = SafeInt(p2(2), 0)
            If a1 < a2 Or (a1 = a2 And m1 < m2) Then
                tmp = a(i)
                a(i) = a(i+1)
                a(i+1) = tmp
                swapped = True
            End If
        Next
    Loop While swapped
End Sub

If UBound(arrDiret) >= 0 Then Call SortByAnoMesDesc(arrDiret)
If UBound(arrGer) >= 0 Then Call SortByAnoMesDesc(arrGer)
If UBound(arrEmp) >= 0 Then Call SortByAnoMesDesc(arrEmp)

' -----------------------------
' HTML RENDER
' -----------------------------
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>Metas - Diretoria / Gerência</title>
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css">
    <style>
        .card-header-grad { background: linear-gradient(90deg,#4361ee,#3a0ca3); color:#fff; }
        .small-muted{ font-size:0.85rem;color:#6c757d; }
        tfoot td { font-weight: 700; background:#f8f9fa !important; }
        .total-row { background-color: #e9ecef !important; }
        .table th { background-color: #f8f9fa; }
        .valor-num { font-weight: 500; }
    </style>
</head>
<body>
<div class="container-fluid py-4">
    <div class="row mb-3">
        <div class="col-12">
            <div class="card shadow-sm">
                <div class="card-body d-flex justify-content-between align-items-center">
                    <div>
                        <h4 class="mb-0"><i class="bi bi-flag-fill me-2"></i>Metas por Diretoria / Gerência</h4>
                        <div class="small-muted">Filtros por Ano e Mês — totais dinâmicos conforme busca / filtro</div>
                    </div>
                    <div>
                        <a href="meta_gerenciamento.asp" class="btn btn-outline-primary me-2"><i class="bi bi-pencil-square"></i> Gerenciar</a>
                        <a href="meta_resumo_anual.asp" class="btn btn-outline-success"><i class="bi bi-calendar-check"></i> Resumo Anual</a>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- filtros -->
    <div class="row mb-3">
        <div class="col-lg-6">
            <form id="filtrosForm" class="row g-2" method="GET">
                <div class="col-6">
                    <label class="form-label">Ano</label>
<select name="ano" class="form-select" onchange="this.form.submit();">
                        <option value="">Todos</option>
                        <% 
                        Dim idxOpt, optYear

                        ' ** CORREÇÃO: Verifica se o array existe e se há elementos (UBound >= 0) **
                        If IsArray(anosArray) And UBound(anosArray) >= 0 Then
                            For idxOpt = 0 To UBound(anosArray)
                                optYear = CStr(anosArray(idxOpt))
                                Response.Write "<option value=""" & optYear & """"
                                
                                ' Adiciona verificação IsNumeric para evitar Type Mismatch no CInt
                                If anoSelecionado <> "" And IsNumeric(optYear) And IsNumeric(anoSelecionado) Then
                                    If CInt(optYear) = CInt(anoSelecionado) Then Response.Write " selected"
                                End If
                                
                                Response.Write ">" & optYear & "</option>"
                            Next
                        End If
                        %>
                    </select>
                </div>
                <div class="col-6">
                    <label class="form-label">Mês</label>
                    <select name="mes" class="form-select" onchange="this.form.submit();">
                        <option value="">Todos</option>
                        <% 
                        For idxOpt = 1 To 12
                            Response.Write "<option value=""" & idxOpt & """"
                            If IsNumeric(mesSelecionado) Then
                                If CInt(mesSelecionado) = idxOpt Then Response.Write " selected"
                            End If
                            Response.Write ">" & mesesNomes(idxOpt) & "</option>"
                        Next
                        %>
                    </select>
                </div>
            </form>
        </div>
    </div>

    <!-- Metas por Diretoria -->
    <div class="row mb-4">
        <div class="col-12">
            <div class="card shadow-sm">
                <div class="card-header card-header-grad">
                    <h5 class="mb-0"><i class="bi bi-building me-2"></i>Metas por Diretoria <span class="badge bg-light text-dark ms-2"><%=dictDiretorias.Count%></span></h5>
                </div>
                <div class="card-body">
                    <% If dictDiretorias.Count > 0 Then %>
                        <div class="table-responsive">
                            <table id="tabelaDiretorias" class="table table-striped table-bordered table-sm">
                                <thead class="table-light">
                                    <tr>
                                        <th>Diretoria</th>
                                        <th>Ano</th>
                                        <th>Mês</th>
                                        <th class="text-end">Meta</th>
                                        <th>Alterado por</th>
                                        <th>Data/Hora</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <% 
                                    'Dim it, parts, dnome, dano, dmes, dvalor, dusuario, ddata, dvalorNum
                                    For Each it In arrDiret
                                        parts = Split(it, "|")
                                        dnome = Server.HTMLEncode(parts(0))
                                        dano = parts(1)
                                        dmes = parts(2)
                                        dvalor = parts(3)
                                        dusuario = Server.HTMLEncode(parts(4))
                                        ddata = Server.HTMLEncode(parts(5))
                                        
                                        dvalorNum = GetNumericValue(dvalor)
                                    %>
                                    <tr>
                                        <td><i class="bi bi-diagram-3 me-1 text-primary"></i><%= dnome %></td>
                                        <td><span class="badge bg-secondary"><%= dano %></span></td>
                                        <td><span class="badge bg-info"><%= mesesNomes(CInt(dmes)) %></span></td>
                                        <td class="text-end valor-num" data-num="<%= Replace(CStr(dvalorNum), ",", ".") %>"><strong><%= FormatMoneyBR(dvalor) %></strong></td>
                                        <td><%= dusuario %></td>
                                        <td class="small-muted"><%= ddata %></td>
                                    </tr>
                                    <% Next %>
                                </tbody>
                                <tfoot>
                                    <tr class="total-row">
                                        <td colspan="3" class="text-end"><strong>TOTAL GERAL:</strong></td>
                                        <td class="text-end" id="footerDiretorias"><strong>R$ 0,00</strong></td>
                                        <td colspan="2"></td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    <% Else %>
                        <div class="alert alert-warning">Nenhuma meta de diretoria encontrada para os filtros selecionados.</div>
                    <% End If %>
                </div>
            </div>
        </div>
    </div>

    <!-- Metas por Gerência -->
    <div class="row mb-4">
        <div class="col-12">
            <div class="card shadow-sm">
                <div class="card-header bg-light">
                    <h5 class="mb-0"><i class="bi bi-people-fill me-2"></i>Metas por Gerência <span class="badge bg-secondary ms-2"><%=dictGerencias.Count%></span></h5>
                </div>
                <div class="card-body">
                    <% If dictGerencias.Count > 0 Then %>
                        <div class="table-responsive">
                            <table id="tabelaGerencias" class="table table-striped table-bordered table-sm">
                                <thead class="table-light">
                                    <tr>
                                        <th>Gerência</th>
                                        <th>Ano</th>
                                        <th>Mês</th>
                                        <th class="text-end">Meta</th>
                                        <th>Alterado por</th>
                                        <th>Data/Hora</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <% 
                                    'Dim gparts, gnome, gano, gmes, gvalor, gusuario, gdata, gvalorNum
                                    For Each it In arrGer
                                        gparts = Split(it, "|")
                                        gnome = Server.HTMLEncode(gparts(0))
                                        gano = gparts(1)
                                        gmes = gparts(2)
                                        gvalor = gparts(3)
                                        gusuario = Server.HTMLEncode(gparts(4))
                                        gdata = Server.HTMLEncode(gparts(5))
                                        
                                        gvalorNum = GetNumericValue(gvalor)
                                    %>
                                    <tr>
                                        <td><i class="bi bi-person-badge me-1 text-success"></i><%= gnome %></td>
                                        <td><span class="badge bg-secondary"><%= gano %></span></td>
                                        <td><span class="badge bg-info"><%= mesesNomes(CInt(gmes)) %></span></td>
                                        <td class="text-end valor-num" data-num="<%= Replace(CStr(gvalorNum), ",", ".") %>"><strong><%= FormatMoneyBR(gvalor) %></strong></td>
                                        <td><%= gusuario %></td>
                                        <td class="small-muted"><%= gdata %></td>
                                    </tr>
                                    <% Next %>
                                </tbody>
                                <tfoot>
                                    <tr class="total-row">
                                        <td colspan="3" class="text-end"><strong>TOTAL GERAL:</strong></td>
                                        <td class="text-end" id="footerGerencias"><strong>R$ 0,00</strong></td>
                                        <td colspan="2"></td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    <% Else %>
                        <div class="alert alert-warning">Nenhuma meta de gerência encontrada para os filtros selecionados.</div>
                    <% End If %>
                </div>
            </div>
        </div>
    </div>

    <!-- Meta Empresa -->
    <div class="row mb-4">
        <div class="col-12">
            <div class="card shadow-sm">
                <div class="card-header bg-light">
                    <h5 class="mb-0"><i class="bi bi-building me-2"></i>Meta da Empresa <span class="badge bg-secondary ms-2"><%=dictEmpresa.Count%></span></h5>
                </div>
                <div class="card-body">
                    <% If dictEmpresa.Count > 0 Then %>
                        <div class="table-responsive">
                            <table id="tabelaEmpresa" class="table table-striped table-bordered table-sm">
                                <thead class="table-light">
                                    <tr>
                                        <th>Período</th>
                                        <th class="text-end">Meta</th>
                                        <th>Data/Hora</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <% 
                                    'Dim eparts, ename, eano, emes, evalor, edata, evalorNum
                                    For Each it In arrEmp
                                        eparts = Split(it, "|")
                                        ename = Server.HTMLEncode(eparts(0))
                                        eano = eparts(1)
                                        emes = eparts(2)
                                        evalor = eparts(3)
                                        edata = Server.HTMLEncode(eparts(5))
                                        
                                        evalorNum = GetNumericValue(evalor)
                                    %>
                                    <tr>
                                        <td><span class="badge bg-secondary"><%= eano %></span> <span class="badge bg-info"><%= mesesNomes(CInt(emes)) %></span></td>
                                        <td class="text-end valor-num" data-num="<%= Replace(CStr(evalorNum), ",", ".") %>"><strong><%= FormatMoneyBR(evalor) %></strong></td>
                                        <td class="small-muted"><%= edata %></td>
                                    </tr>
                                    <% Next %>
                                </tbody>
                                <tfoot>
                                    <tr class="total-row">
                                        <td class="text-end"><strong>TOTAL GERAL:</strong></td>
                                        <td class="text-end" id="footerEmpresa"><strong>R$ 0,00</strong></td>
                                        <td></td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    <% Else %>
                        <div class="alert alert-warning">Nenhuma meta da empresa encontrada para os filtros selecionados.</div>
                    <% End If %>
                </div>
            </div>
        </div>
    </div>

</div>

<!-- Scripts -->
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>

<script>
$(document).ready(function() {
    // Função para converter texto em número
    function parseCurrency(value) {
        if (typeof value === 'number') return value;
        
        // Se for objeto jQuery, pega o texto
        if (value instanceof jQuery) {
            value = value.text();
        }
        
        // Extrai valor numérico
        var str = String(value).trim();
        
        // Primeiro tenta pegar do data-num
        if (value instanceof jQuery && value.attr('data-num')) {
            var dataNum = parseFloat(value.attr('data-num').replace(',', '.'));
            if (!isNaN(dataNum)) return dataNum;
        }
        
        // Remove R$, pontos e converte vírgula para ponto
        str = str.replace(/[^\d,.-]/g, '');
        str = str.replace(/\./g, '');
        str = str.replace(',', '.');
        
        var num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    }
    
    // Função para formatar número como moeda brasileira
    function formatCurrency(num) {
        return num.toLocaleString('pt-BR', {
            style: 'currency',
            currency: 'BRL',
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        });
    }
    
    // Função para calcular total de uma coluna
    function calculateTotal(table, columnIndex) {
        var total = 0;
        var api = table.DataTable();
        
        // Para todas as linhas (incluindo filtradas)
        api.rows({ search: 'applied' }).every(function() {
            var cell = this.cell(this.index(), columnIndex).node();
            total += parseCurrency($(cell));
        });
        
        return total;
    }
    
    // Configuração da tabela Diretorias
    var tableDiret = $('#tabelaDiretorias').DataTable({
        pageLength: 25,
        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, "Todos"]],
        order: [],
        language: {
            url: "https://cdn.datatables.net/plug-ins/1.13.6/i18n/pt-BR.json"
        },
        footerCallback: function(row, data, start, end, display) {
            var api = this.api();
            var colIndex = 3; // Coluna de valores (base 0)
            
            // Calcular total para todas as linhas visíveis (com filtros aplicados)
            var total = 0;
            api.rows({ search: 'applied' }).every(function() {
                var data = this.data();
                var cellValue = api.cell(this, colIndex).data();
                total += parseCurrency(cellValue);
            });
            
            // Atualizar rodapé
            $('#footerDiretorias').html('<strong>' + formatCurrency(total) + '</strong>');
        },
        drawCallback: function(settings) {
            // Recalcular ao redesenhar a tabela
            var api = this.api();
            var colIndex = 3;
            var total = 0;
            
            api.rows({ search: 'applied' }).every(function() {
                var data = this.data();
                var cellValue = api.cell(this, colIndex).data();
                total += parseCurrency(cellValue);
            });
            
            $('#footerDiretorias').html('<strong>' + formatCurrency(total) + '</strong>');
        }
    });
    
    // Configuração da tabela Gerências
    var tableGer = $('#tabelaGerencias').DataTable({
        pageLength: 25,
        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, "Todos"]],
        order: [],
        language: {
            url: "https://cdn.datatables.net/plug-ins/1.13.6/i18n/pt-BR.json"
        },
        footerCallback: function(row, data, start, end, display) {
            var api = this.api();
            var colIndex = 3;
            
            var total = 0;
            api.rows({ search: 'applied' }).every(function() {
                var data = this.data();
                var cellValue = api.cell(this, colIndex).data();
                total += parseCurrency(cellValue);
            });
            
            $('#footerGerencias').html('<strong>' + formatCurrency(total) + '</strong>');
        },
        drawCallback: function(settings) {
            var api = this.api();
            var colIndex = 3;
            var total = 0;
            
            api.rows({ search: 'applied' }).every(function() {
                var data = this.data();
                var cellValue = api.cell(this, colIndex).data();
                total += parseCurrency(cellValue);
            });
            
            $('#footerGerencias').html('<strong>' + formatCurrency(total) + '</strong>');
        }
    });
    
    // Configuração da tabela Empresa
    var tableEmp = $('#tabelaEmpresa').DataTable({
        pageLength: 25,
        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, "Todos"]],
        order: [],
        language: {
            url: "https://cdn.datatables.net/plug-ins/1.13.6/i18n/pt-BR.json"
        },
        footerCallback: function(row, data, start, end, display) {
            var api = this.api();
            var colIndex = 1;
            
            var total = 0;
            api.rows({ search: 'applied' }).every(function() {
                var data = this.data();
                var cellValue = api.cell(this, colIndex).data();
                total += parseCurrency(cellValue);
            });
            
            $('#footerEmpresa').html('<strong>' + formatCurrency(total) + '</strong>');
        },
        drawCallback: function(settings) {
            var api = this.api();
            var colIndex = 1;
            var total = 0;
            
            api.rows({ search: 'applied' }).every(function() {
                var data = this.data();
                var cellValue = api.cell(this, colIndex).data();
                total += parseCurrency(cellValue);
            });
            
            $('#footerEmpresa').html('<strong>' + formatCurrency(total) + '</strong>');
        }
    });
    
    // Atualizar totais quando o usuário digitar na busca
    $('.dataTables_filter input').on('keyup', function() {
        setTimeout(function() {
            tableDiret.draw();
            tableGer.draw();
            tableEmp.draw();
        }, 500);
    });
    
    // Calcular totais iniciais
    setTimeout(function() {
        tableDiret.draw();
        tableGer.draw();
        tableEmp.draw();
    }, 300);
});
</script>

</body>
</html>

<%
' -----------------------------
' LIMPAR OBJETOS
' -----------------------------
If IsObject(dictDiretorias) Then Set dictDiretorias = Nothing
If IsObject(dictGerencias) Then Set dictGerencias = Nothing
If IsObject(dictEmpresa) Then Set dictEmpresa = Nothing

If IsObject(connOrg) Then
    connOrg.Close
    Set connOrg = Nothing
End If

If IsObject(connSales) Then
    connSales.Close
    Set connSales = Nothing
End If
%>