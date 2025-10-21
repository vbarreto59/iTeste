<%@ Language=VBScript Codepage=65001 %>
<% 
' Configuração UTF-8 para todo o documento
Response.CodePage = 65001
Response.CharSet = "UTF-8"
%>
<%if Trim(StrConn)="" then%>
     <!--#include file="conexao.asp"-->
<%end if%>     
<%if Trim(StrConnSales)="" then%>
     <!--#include file="conSunSales.asp"-->
<%end if%>  


<%
' =========================================================================
' VARIÁVEIS ADO E GLOBAIS
' =========================================================================
Dim conn, rs
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

Dim ano, mes, trimestre, semestre, whereClause, isFiltered
Dim connUsuarios

isFiltered = False


' =========================================================================
' 1. PROCESSAMENTO DE FORMULÁRIO E CONSTRUÇÃO DA WHERE CLAUSE
' =========================================================================

' Leitura dos filtros do formulário
ano = Request.Form("ano")
mes = Request.Form("mes")
trimestre = Request.Form("trimestre")
semestre = Request.Form("semestre")

' Cláusula WHERE inicial (Usando FALSE para booleanos, assumindo que Excluido é um campo Sim/Não)
whereClause = " WHERE [Vendas].[Excluido] = FALSE" 
' Se o campo 'Excluido' for numérico (0/1), use: whereClause = " WHERE [Vendas].[Excluido] = 0"

' =========================================================================
' Construção da Cláusula WHERE de Vendas
' (Usando CLng() nos campos de data/período para evitar o erro 80040e10)
' =========================================================================
If Not IsEmpty(ano) And ano <> "" And IsNumeric(ano) Then
    whereClause = whereClause & " AND CLng([Vendas].[AnoVenda]) = " & CLng(ano)
    isFiltered = True
End If

If Not IsEmpty(mes) And mes <> "" And IsNumeric(mes) Then
    whereClause = whereClause & " AND CLng([Vendas].[MesVenda]) = " & CLng(mes)
    isFiltered = True
End If

If Not IsEmpty(trimestre) And trimestre <> "" And IsNumeric(trimestre) Then
    whereClause = whereClause & " AND CLng([Vendas].[Trimestre]) = " & CLng(trimestre)
    isFiltered = True
End If

If Not IsEmpty(semestre) And semestre <> "" And IsNumeric(semestre) Then
    whereClause = whereClause & " AND CLng([Vendas].[Semestre]) = " & CLng(semestre)
    isFiltered = True
End If

' =========================================================================
' 2. TENTATIVA INICIAL DE ABERTURA DA CONEXÃO PRINCIPAL (VENDAS) - StrConnSales
' =========================================================================
If isFiltered Then
    On Error Resume Next
    conn.Open StrConnSales
    If Err.Number <> 0 Then
        Response.Write "<!DOCTYPE html><html><body><div class='alert alert-danger'>Erro crítico ao conectar ao banco de dados de Vendas (StrConnSales): " & Err.Description & "</div></body></html>"
        conn.Close
        Set conn = Nothing
        Response.End
    End If
    On Error GoTo 0 
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <title>Relatório de Corretores Sem Vendas</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f8f9fa;
        }
        .container {
            background-color: white;
            border-radius: 10px;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
            padding: 30px;
            margin-top: 20px;
            margin-bottom: 20px;
        }
        h1 {
            color: #2c3e50;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
            margin-bottom: 30px;
        }
        .card-header {
            background-color: #3498db;
            color: white;
            font-weight: bold;
        }
        .btn-primary {
            background-color: #3498db;
            border-color: #3498db;
        }
        .btn-primary:hover {
            background-color: #2980b9;
            border-color: #2980b9;
        }
        .table th {
            background-color: #2c3e50;
            color: white;
        }
        .alert-success {
            background-color: #d4edda;
            border-color: #c3e6cb;
            color: #155724;
        }
        .alert-warning {
            background-color: #fff3cd;
            border-color: #ffeaa7;
            color: #856404;
        }
        .gerencia-header {
            background-color: #e9ecef !important;
            font-weight: bold;
            font-size: 1.1em;
            color: #495057;
        }
        .gerencia-total {
            background-color: #d1ecf1;
            font-weight: bold;
        }
    </style>
</head>
<body>

<div class="container mt-5">
    <h1>📊 Relatório: Corretores Sem Vendas</h1>

    <div class="card mb-4">
        <div class="card-header">Filtros do Relatório</div>
        <form method="POST" action="">
            <div class="card-body row">
                
                <div class="form-group col-md-3">
                    <label for="ano">Ano:</label>
                    <input type="number" id="ano" name="ano" class="form-control" value="<%=Server.HTMLEncode(ano)%>" placeholder="Ex: 2024">
                </div>
                
                <div class="form-group col-md-3">
                    <label for="mes">Mês:</label>
                    <select id="mes" name="mes" class="form-control">
                        <option value="">Todos os meses</option>
                        <%
                        Dim i
                        For i = 1 To 12
                            Response.Write "<option value='" & i & "'"
                            If CStr(mes) = CStr(i) Then Response.Write " selected"
                            Response.Write ">" & MonthName(i, False) & "</option>"
                        Next
                        %>
                    </select>
                </div>

                <div class="form-group col-md-3">
                    <label for="trimestre">Trimestre:</label>
                    <select id="trimestre" name="trimestre" class="form-control">
                        <option value="">Todos os trimestres</option>
                        <%
                        For i = 1 To 4
                            Response.Write "<option value='" & i & "'"
                            If CStr(trimestre) = CStr(i) Then Response.Write " selected"
                            Response.Write ">" & i & "º Trimestre</option>"
                        Next
                        %>
                    </select>
                </div>
                
                <div class="form-group col-md-3">
                    <label for="semestre">Semestre:</label>
                    <select id="semestre" name="semestre" class="form-control">
                        <option value="">Todos os semestres</option>
                        <%
                        For i = 1 To 2
                            Response.Write "<option value='" & i & "'"
                            If CStr(semestre) = CStr(i) Then Response.Write " selected"
                            Response.Write ">" & i & "º Semestre</option>"
                        Next
                        %>
                    </select>
                </div>
                
                <div class="col-12 mt-3">
                    <button type="submit" class="btn btn-primary">🔍 Gerar Relatório</button>
                    <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>" class="btn btn-secondary">🔄 Limpar Filtros</a>
                </div>

            </div>
        </form>
    </div>
    
    <% If Not isFiltered Then %>
        <div class="alert alert-info text-center">
            <h5>👋 Bem-vindo ao Relatório de Corretores Sem Vendas</h5>
            <p>Selecione pelo menos um critério de filtro acima e clique em <strong>'Gerar Relatório'</strong> para visualizar os corretores que não realizaram vendas no período selecionado.</p>
        </div>
    <% Else %>

        <%
        Dim isConnUsuariosOk
        isConnUsuariosOk = False 
        Dim isConnSalesOk
        isConnSalesOk = True 

        ' 1. Configurar e Abrir a Conexão para a tabela 'usuarios' (StrConn)
        Set connUsuarios = Server.CreateObject("ADODB.Connection")
        On Error Resume Next
        connUsuarios.Open StrConn
        
        If Err.Number <> 0 Then
            Response.Write "<div class='alert alert-danger'>❌ Erro ao conectar ao banco de dados de Usuários (StrConn): " & Err.Description & "</div>"
            connUsuarios.Close
            Set connUsuarios = Nothing
        Else
            isConnUsuariosOk = True 
        End If
        On Error GoTo 0 

        ' Somente executa o relatório se a conexão de usuários estiver OK
        If isConnUsuariosOk Then

            ' 2. GARANTIR A REABERTURA DA CONEXÃO DE VENDAS (conn / StrConnSales)
            If Not IsObject(conn) Or conn.State <> 1 Then 
                On Error Resume Next
                If IsObject(conn) Then conn.Open StrConnSales
                If Err.Number <> 0 Then
                    Response.Write "<div class='alert alert-danger'>❌ Erro ao reabrir a conexão de Vendas (StrConnSales): " & Err.Description & "</div>"
                    isConnSalesOk = False
                End If
                On Error GoTo 0
            End If
            
            ' Verifica se a conexão de vendas está realmente aberta
            If conn.State <> 1 Then isConnSalesOk = False

            ' 3. EXECUÇÃO DO PASSO 1, SOMENTE SE A CONEXÃO DE VENDAS ESTIVER OK
            If isConnSalesOk Then
                
                ' Passo 1: Obter a lista de IDs de corretores que fizeram vendas (StrConnSales/conn)
                Dim sql_vendedores_com_venda, rsVendedores
                Set rsVendedores = Server.CreateObject("ADODB.Recordset")
                
                ' A query usa o whereClause construído acima.
                sql_vendedores_com_venda = "SELECT DISTINCT [Vendas].[corretorid] FROM [Vendas]" & whereClause
                
                On Error Resume Next
                rsVendedores.Open sql_vendedores_com_venda, conn ' Usa a conexão de Vendas
                
                If Err.Number <> 0 Then
                     Response.Write "<div class='alert alert-danger'>❌ Erro ao executar a query de Vendas (Corretores COM Vendas): " & Err.Description & " - SQL: " & Server.HTMLEncode(sql_vendedores_com_venda) & "</div>"
                     
                     If IsObject(rsVendedores) And rsVendedores.State = 1 Then rsVendedores.Close
                     Set rsVendedores = Nothing
                     isConnSalesOk = False 
                End If
                On Error GoTo 0
            
            End If
            
            ' 4. PROCESSAMENTO DO RELATÓRIO
            If isConnSalesOk Then 
                
                Dim idCorretoresComVendas
                idCorretoresComVendas = ""

                ' Verifica se o recordset foi criado e preenchido
                If IsObject(rsVendedores) And Not rsVendedores.EOF Then
                    Do While Not rsVendedores.EOF
                        If IsNumeric(rsVendedores("corretorid")) And Not IsNull(rsVendedores("corretorid")) Then
                            If idCorretoresComVendas <> "" Then idCorretoresComVendas = idCorretoresComVendas & ","
                            idCorretoresComVendas = idCorretoresComVendas & rsVendedores("corretorid")
                        End If
                        rsVendedores.MoveNext
                    Loop
                End If
                
                If IsObject(rsVendedores) And rsVendedores.State = 1 Then rsVendedores.Close
                Set rsVendedores = Nothing

                ' Passo 2: Construir a consulta final AGRUPADA POR GERÊNCIA
                Dim filtro_not_in
                filtro_not_in = ""
                ' Aqui, UserId deve corresponder ao corretorid na tabela de usuários
                If idCorretoresComVendas <> "" Then
                    filtro_not_in = " AND usuarios.UserId NOT IN (" & idCorretoresComVendas & ")"
                End If
                
                ' CONSULTA MODIFICADA: Selecionando apenas nome e gerência
                Dim sql_corretores_sem_vendas
                sql_corretores_sem_vendas = "SELECT nome, Gerencia FROM usuarios " & _
                                            "WHERE usuarios.permissao = 5 AND usuarios.idEmp = 2" & filtro_not_in & " " & _
                                            "ORDER BY Gerencia, nome ASC"

                ' Abrir o recordset usando a conexão 'connUsuarios'
                Set rs.ActiveConnection = connUsuarios 
                rs.Open sql_corretores_sem_vendas

                ' Variáveis para controle do agrupamento
                Dim currentGerencia, firstGerencia, totalCorretores, gerenciaCount
                currentGerencia = ""
                firstGerencia = True
                totalCorretores = 0
                %>
                <div class="card mb-4 mt-4">
                    <div class="card-body">
                        <h2 class="card-title">📈 Resultado da Análise</h2>
                        
                        <% ' Exibir resumo dos filtros ativos %>
                        <div class="alert alert-warning mb-4">
                            <h6>🔍 Filtros Aplicados:</h6>
                            <%
                            Dim filtrosAtivos
                            filtrosAtivos = ""
                            If ano <> "" Then filtrosAtivos = filtrosAtivos & " <span class='badge badge-primary'>Ano: " & ano & "</span>"
                            If mes <> "" Then filtrosAtivos = filtrosAtivos & " <span class='badge badge-primary'>Mês: " & MonthName(mes, False) & "</span>"
                            If trimestre <> "" Then filtrosAtivos = filtrosAtivos & " <span class='badge badge-primary'>" & trimestre & "º Trimestre</span>"
                            If semestre <> "" Then filtrosAtivos = filtrosAtivos & " <span class='badge badge-primary'>" & semestre & "º Semestre</span>"
                            
                            If filtrosAtivos = "" Then
                                Response.Write "<em>Nenhum filtro específico aplicado</em>"
                            Else
                                Response.Write filtrosAtivos
                            End If
                            %>
                        </div>
                        
                        <%
                        If Not rs.EOF Then
                            %>
                            <p class="card-text">📋 Lista de corretores ativos <strong>(Permissão 5, Empresa 2)</strong> que <strong class='text-danger'>NÃO</strong> registraram vendas no período selecionado, agrupados por gerência:</p>
                            <div class="table-responsive">
                                <table class="table table-striped table-bordered table-hover">
                                    <thead>
                                        <tr>
                                            <th>Gerência</th>
                                            <th>Nome do Corretor</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <%
                                        Do While Not rs.EOF
                                            ' Verifica se mudou de gerência
                                            If currentGerencia <> rs("Gerencia") Then
                                                ' Se não é a primeira gerência, fecha a anterior
                                                If Not firstGerencia Then
                                                    ' Exibe total da gerência anterior
                                                    %>
                                                    <tr class="gerencia-total">
                                                        <td class="text-right"><strong>Total da Gerência:</strong></td>
                                                        <td><strong><%=gerenciaCount%> corretor(es)</strong></td>
                                                    </tr>
                                                    <%
                                                End If
                                                
                                                ' Nova gerência
                                                currentGerencia = rs("Gerencia")
                                                gerenciaCount = 0
                                                
                                                ' Exibe header da nova gerência
                                                Dim displayGerencia
                                                displayGerencia = currentGerencia
                                                If IsNull(displayGerencia) Or displayGerencia = "" Then
                                                    displayGerencia = "Sem Gerência Definida"
                                                End If
                                                %>
                                                <tr class="gerencia-header">
                                                    <td colspan="2">
                                                        <strong>🏢 <%=displayGerencia%></strong>
                                                    </td>
                                                </tr>
                                                <%
                                                firstGerencia = False
                                            End If
                                            
                                            ' Incrementa contadores
                                            gerenciaCount = gerenciaCount + 1
                                            totalCorretores = totalCorretores + 1
                                            %>
                                            <tr>
                                                <td></td>
                                                <td><%=rs("nome")%></td>
                                            </tr>
                                            <%
                                            rs.MoveNext
                                        Loop
                                        
                                        ' Exibe total da última gerência
                                        If Not firstGerencia Then
                                            %>
                                            <tr class="gerencia-total">
                                                <td class="text-right"><strong>Total da Gerência:</strong></td>
                                                <td><strong><%=gerenciaCount%> corretor(es)</strong></td>
                                            </tr>
                                            <%
                                        End If
                                        
                                        ' Exibe total geral
                                        %>
                                        <tr class="table-primary">
                                            <td class="text-right"><strong>📊 TOTAL GERAL:</strong></td>
                                            <td><strong><%=totalCorretores%> corretor(es) sem vendas</strong></td>
                                        </tr>
                                        <%
                                        %>
                                    </tbody>
                                </table>
                            </div>
                            <%
                        Else
                            If idCorretoresComVendas <> "" Then
                                %>
                                <div class="alert alert-success text-center">
                                    <h5>🎉 Parabéns!</h5>
                                    <p>Todos os corretores ativos <strong>(Permissão 5, Empresa 2)</strong> registraram vendas no período selecionado!</p>
                                    <p class="mb-0"><small>Isso demonstra um excelente desempenho da equipe comercial.</small></p>
                                </div>
                                <%
                            Else
                                %>
                                <div class="alert alert-info text-center">
                                    <h5>ℹ️ Informação</h5>
                                    <p>Nenhum corretor ativo <strong>(Permissão 5, Empresa 2)</strong> foi encontrado sem vendas no período.</p>
                                </div>
                                <%
                            End If
                        End If
                        
                        rs.Close
                        %>
                    </div>
                </div>
                <%
            End If ' end if isConnSalesOk
        End If ' end if isConnUsuariosOk
    End If ' end if isFiltered
    %>
    
</div>

<script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.5.4/dist/umd/popper.min.js"></script>
<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>

</body>
</html>

<%
' =========================================================================
' FECHAMENTO DE CONEXÕES E LIMPEZA DE OBJETOS (FINAL DA PÁGINA)
' =========================================================================

If IsObject(rs) Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If

If IsObject(conn) Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If

If IsObject(connUsuarios) Then
    If connUsuarios.State = 1 Then connUsuarios.Close 
    Set connUsuarios = Nothing 
End If
%>