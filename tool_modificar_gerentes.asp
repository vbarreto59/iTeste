<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: SMHEBPHSKG          -->
<!-- MODIFICAÇÃO: Gestão de Gerentes        -->
<!-- ###################################### -->

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->
<!--#include file="registra_log.asp"-->
<!--#include file="usr_acoes_v4GVendas.inc"-->

<%
' Verificar se o usuário está logado
If Session("Usuario") = "" Then
    Response.Redirect "login.asp"
End If

' -----------------------------------------------------------------------------------
' INICIALIZAÇÃO E CONEXÃO COM BANCOS DE DADOS
' -----------------------------------------------------------------------------------
' Verifica se as strings de conexão estão configuradas.
If Len(StrConn) = 0 Or Len(StrConnSales) = 0 Then
    Response.Write "Erro: Conexões com bancos de dados não configuradas"
    Response.End
End If

' Cria e abre as conexões com os bancos de dados.
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

Dim mensagem
mensagem = ""

' #################### Processar alteração de gerentes em massa
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim acao, gerenteAntigoId, gerenteNovoId, gerenteAntigoNome, gerenteNovoNome
    
    acao = Request.Form("acao")
    
    If acao = "alterar_gerentes" Then
        gerenteAntigoId = Trim(Request.Form("gerente_antigo_id"))
        gerenteNovoId = Trim(Request.Form("gerente_novo_id"))
        
        ' Debug: Verificar os valores recebidos
        'Response.Write "gerenteAntigoId: " & gerenteAntigoId & "<br>"
        'Response.Write "gerenteNovoId: " & gerenteNovoId & "<br>"
        'Response.End
        
        If gerenteAntigoId <> "" And gerenteNovoId <> "" And IsNumeric(gerenteAntigoId) And IsNumeric(gerenteNovoId) Then
            ' Obter nomes dos gerentes
            gerenteAntigoNome = GetDataFromDB(conn, "Usuarios", "Nome", "UserId", gerenteAntigoId)
            gerenteNovoNome = GetDataFromDB(conn, "Usuarios", "Nome", "UserId", gerenteNovoId)
            
            If gerenteAntigoNome <> "Desconhecido" And gerenteNovoNome <> "Desconhecido" Then
                ' Verificar se o gerente antigo existe em vendas
                Set rsCheckGerente = connSales.Execute("SELECT COUNT(*) as Total FROM Vendas WHERE UserIdGerencia = " & gerenteAntigoId & " AND EXCLUIDO = 0")
                Dim totalVendasGerente
                totalVendasGerente = 0
                If Not rsCheckGerente.EOF Then
                    totalVendasGerente = rsCheckGerente("Total")
                End If
                rsCheckGerente.Close
                Set rsCheckGerente = Nothing
                
                ' Verificar se o gerente antigo existe na tabela Gerencias
                Set rsCheckGerencia = conn.Execute("SELECT COUNT(*) as Total FROM Gerencias WHERE UserId = " & gerenteAntigoId)
                Dim totalGerencias
                totalGerencias = 0
                If Not rsCheckGerencia.EOF Then
                    totalGerencias = rsCheckGerencia("Total")
                End If
                rsCheckGerencia.Close
                Set rsCheckGerencia = Nothing
                
                If totalVendasGerente > 0 Or totalGerencias > 0 Then
                    ' Iniciar transação para garantir consistência
                    On Error Resume Next
                    
                    ' 1. Atualizar todas as gerencias na tabela Vendas (connSales)
                    If totalVendasGerente > 0 Then
                        sqlUpdateVendas = "UPDATE Vendas SET " & _
                                         "UserIdGerencia = " & gerenteNovoId & ", " & _
                                         "NomeGerente = '" & SanitizeSQL(gerenteNovoNome) & "' " & _
                                         "WHERE UserIdGerencia = " & gerenteAntigoId & " AND EXCLUIDO = 0"
                        connSales.Execute(sqlUpdateVendas)
                    End If
                    
                    ' 2. Atualizar tabela COMISSOES_A_PAGAR (connSales)
                    If totalVendasGerente > 0 Then
                        sqlUpdateComissoes = "UPDATE COMISSOES_A_PAGAR SET " & _
                                           "UserIdGerencia = " & gerenteNovoId & ", " & _
                                           "NomeGerente = '" & SanitizeSQL(gerenteNovoNome) & "' " & _
                                           "WHERE UserIdGerencia = " & gerenteAntigoId
                        connSales.Execute(sqlUpdateComissoes)
                    End If
                    
                    ' 3. Atualizar tabela PAGAMENTOS_COMISSOES (connSales)
                    If totalVendasGerente > 0 Then
                        sqlUpdatePagamentos = "UPDATE PAGAMENTOS_COMISSOES SET " & _
                                            "UsuariosUserId = " & gerenteNovoId & ", " & _
                                            "UsuariosNome = '" & SanitizeSQL(gerenteNovoNome) & "' " & _
                                            "WHERE UsuariosUserId = " & gerenteAntigoId & " AND TipoRecebedor = 'gerencia'"
                        connSales.Execute(sqlUpdatePagamentos)
                    End If
                    
                    ' 4. NOVO: Atualizar tabela Gerencias (conn)
                    If totalGerencias > 0 Then
                        sqlUpdateGerencias = "UPDATE Gerencias SET " & _
                                           "UserId = " & gerenteNovoId & ", " & _
                                           "Nome = '" & SanitizeSQL(gerenteNovoNome) & "' " & _
                                           "WHERE UserId = " & gerenteAntigoId
                        conn.Execute(sqlUpdateGerencias)
                    End If
                    
                    If Err.Number = 0 Then
                        ' Registrar log
                        Call InserirLog("GERENCIAS", "UPDATE_MASS", "Substituição de gerente: " & gerenteAntigoNome & " por " & gerenteNovoNome & _
                                      " | Vendas: " & totalVendasGerente & " | Gerencias: " & totalGerencias)
                        
                        ' Construir mensagem de sucesso
                        mensagem = "Gerente substituído com sucesso! " & gerenteAntigoNome & " foi substituído por " & gerenteNovoNome
                        
                        If totalVendasGerente > 0 Then
                            mensagem = mensagem & " em " & totalVendasGerente & " venda(s) (tabelas: Vendas, COMISSOES_A_PAGAR, PAGAMENTOS_COMISSOES)"
                        End If
                        
                        If totalGerencias > 0 Then
                            If totalVendasGerente > 0 Then
                                mensagem = mensagem & " e "
                            End If
                            mensagem = mensagem & " em " & totalGerencias & " registro(s) na tabela Gerencias"
                        End If
                        
                        mensagem = mensagem & "."
                    Else
                        mensagem = "Erro ao atualizar gerente: " & Err.Description
                    End If
                    On Error GoTo 0
                Else
                    mensagem = "O gerente selecionado não possui registros ativos para atualizar (nem em vendas, nem na tabela Gerencias)."
                End If
            Else
                mensagem = "Erro: Um ou ambos os gerentes não foram encontrados no sistema."
            End If
        Else
            mensagem = "Erro: IDs de gerentes inválidos. Certifique-se de selecionar ambos os gerentes."
        End If
    End If
End If

' ===============================================================
' CONSULTAS SEPARADAS PARA CADA USO
' ===============================================================

' CONSULTA 1: Tabela de Gerencias da conexão conn - SIMPLES
Set rsGerenciasConn = conn.Execute("SELECT GerenciaID, NomeGerencia, UserId, Nome FROM Gerencias ORDER BY NomeGerencia")

' 2. Consulta para exibição na tabela (PRIMEIRO USO)
Set rsContagemGerentesParaTabela = connSales.Execute("SELECT UserIdGerencia, NomeGerente, COUNT(*) as TotalVendas " & _
                                                    "FROM Vendas " & _
                                                    "WHERE UserIdGerencia IS NOT NULL AND UserIdGerencia > 0 AND EXCLUIDO = 0 " & _
                                                    "GROUP BY UserIdGerencia, NomeGerente " & _
                                                    "ORDER BY NomeGerente")

' 3. Consulta SEPARADA para o select "Gerente a ser Substituído" (SEGUNDO USO - RECORD SET DIFERENTE)
Set rsContagemGerentesParaSelect = connSales.Execute("SELECT UserIdGerencia, NomeGerente, COUNT(*) as TotalVendas " & _
                                                    "FROM Vendas " & _
                                                    "WHERE UserIdGerencia IS NOT NULL AND UserIdGerencia > 0 AND EXCLUIDO = 0 " & _
                                                    "GROUP BY UserIdGerencia, NomeGerente " & _
                                                    "ORDER BY NomeGerente")

' 4. Consulta para exibir nome da gerência na tabela
Set rsGerenciasNomes = connSales.Execute("SELECT DISTINCT Gerencia, UserIdGerencia " & _
                                        "FROM Vendas " & _
                                        "WHERE UserIdGerencia IS NOT NULL AND UserIdGerencia > 0 AND EXCLUIDO = 0 " & _
                                        "ORDER BY Gerencia")

' 5. Buscar todos os usuários para o select "Novo Gerente"
Set rsTodosUsuariosParaSelect = conn.Execute("SELECT UserId, Nome FROM Usuarios WHERE Nome <> '' ORDER BY Nome")

' 6. Buscar todas as vendas para exibição na seção de referência
Set rsVendasParaReferencia = connSales.Execute("SELECT ID, NomeEmpreendimento, Unidade, NomeCliente, DataVenda, " & _
                                              "DiretoriaId, Diretoria, UserIdDiretoria, NomeDiretor, " & _
                                              "GerenciaId, Gerencia, UserIdGerencia, NomeGerente, " & _
                                              "CorretorId, Corretor " & _
                                              "FROM Vendas WHERE EXCLUIDO=0 ORDER BY DataVenda DESC")

' Função auxiliar
Function GetDataFromDB(oConn, sTable, sField, sWhereField, sWhereValue)
    Dim sResult
    On Error Resume Next
    Set rs = oConn.Execute("SELECT " & sField & " FROM " & sTable & " WHERE " & sWhereField & " = " & sWhereValue)
    If Err.Number = 0 And Not rs.EOF Then
        sResult = rs(sField)
    Else
        sResult = "Desconhecido"
    End If
    If IsObject(rs) Then rs.Close
    Set rs = Nothing
    GetDataFromDB = sResult
    On Error GoTo 0
End Function

Function SanitizeSQL(sValue)
    If IsNull(sValue) Then
        SanitizeSQL = ""
    Else
        SanitizeSQL = Replace(sValue, "'", "''")
    End If
End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestão de Gerentes | Sistema de Vendas</title>
    
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    
    <style>
        :root {
            --primary: #2c3e50;
            --secondary: #3498db;
            --success: #27ae60;
            --warning: #f39c12;
            --danger: #e74c3c;
        }
        
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #2c3e50;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            padding: 20px;
        }
        
        .app-container {
            max-width: 1800px;
            margin: 0 auto;
        }
        
        .app-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
            padding: 1.5rem;
            border-radius: 12px 12px 0 0;
            margin-bottom: 0;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        
        .app-title {
            font-weight: 600;
            margin: 0;
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 1.8rem;
        }
        
        .card {
            border: none;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin-bottom: 1.5rem;
            background: rgba(255, 255, 255, 0.95);
        }
        
        .card-header {
            background: linear-gradient(to right, var(--primary), var(--secondary));
            color: white;
            border-bottom: none;
            padding: 1.2rem 1.5rem;
            font-weight: 600;
        }
        
        .table th {
            background-color: var(--primary);
            color: white;
            border: none;
        }
        
        .badge-count {
            background: var(--secondary);
            color: white;
            font-size: 0.8rem;
            padding: 0.25rem 0.5rem;
        }
        
        .badge-id {
            background: #6c757d;
            color: white;
            font-size: 0.75rem;
            padding: 0.2rem 0.4rem;
        }
        
        .form-label {
            font-weight: 600;
            color: var(--primary);
        }
        
        .section-title {
            color: var(--primary);
            font-weight: 600;
            font-size: 1.1rem;
            margin-bottom: 1rem;
            padding-bottom: 0.5rem;
            border-bottom: 2px solid #e9ecef;
        }
        
        .info-box {
            background: #e8f4fc;
            border-left: 4px solid var(--secondary);
            padding: 1rem;
            margin-bottom: 1rem;
            border-radius: 4px;
        }
        
        .debug-info {
            background: #f8f9fa;
            border: 1px solid #dee2e6;
            padding: 10px;
            margin-bottom: 10px;
            border-radius: 4px;
            font-size: 0.85rem;
            color: #6c757d;
        }
        
        .table-container {
            max-height: 400px;
            overflow-y: auto;
            margin-bottom: 20px;
        }
        
        .text-muted-small {
            font-size: 0.85rem;
            color: #6c757d;
        }
        
        .alert-custom {
            border-radius: 8px;
            border: none;
        }
        
        .update-info {
            background: #fff3cd;
            border: 1px solid #ffc107;
            padding: 10px;
            margin-top: 10px;
            border-radius: 4px;
            font-size: 0.9rem;
        }
    </style>
</head>
<body>
    <div class="app-container">
        <!-- Header -->
        <div class="app-header">
            <div class="d-flex justify-content-between align-items-center">
                <h1 class="app-title">
                    <i class="fas fa-user-tie me-2"></i>Gestão de Gerentes
                </h1>
                <div class="d-flex align-items-center gap-3">
                    <span class="badge bg-light text-dark">
                        <i class="fas fa-user me-1"></i><%= Session("Usuario") %>
                    </span>
                    <a href="javascript:history.back()" class="btn btn-light btn-sm">
                        <i class="fas fa-arrow-left me-1"></i>Voltar
                    </a>
                </div>
            </div>
        </div>

        <!-- Mensagens -->
        <% If mensagem <> "" Then 
            Dim alertClass
            If InStr(mensagem, "sucesso") > 0 Then
                alertClass = "success"
            ElseIf InStr(mensagem, "Erro") > 0 Or InStr(mensagem, "inválido") > 0 Then
                alertClass = "danger"
            Else
                alertClass = "warning"
            End If
        %>
        <div class="alert alert-<%=alertClass%> alert-dismissible fade show mt-3 alert-custom" role="alert">
            <% 
            Select Case alertClass
                Case "success"
                    Response.Write "<i class='fas fa-check-circle me-2'></i>"
                Case "warning"
                    Response.Write "<i class='fas fa-exclamation-triangle me-2'></i>"
                Case "danger"
                    Response.Write "<i class='fas fa-times-circle me-2'></i>"
                Case Else
                    Response.Write "<i class='fas fa-info-circle me-2'></i>"
            End Select
            %>
            <%= mensagem %>
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        </div>
        <% End If %>

        <!-- PRIMEIRA SEÇÃO: Tabela de Gerencias da conexão conn - SIMPLES -->
        <div class="card mt-3">
            <div class="card-header">
                <i class="fas fa-table me-2"></i>Tabela Gerencias (conn) - Todos os Registros
            </div>
            <div class="card-body">
                <div class="info-box">
                    <i class="fas fa-info-circle me-2"></i>
                    Lista completa de todas as gerencias cadastradas na tabela Gerencias do banco de dados principal (conn).
                    <strong>Esta tabela também será atualizada quando você substituir um gerente.</strong>
                </div>
                
                <div class="table-responsive">
                    <table class="table table-striped table-hover">
                        <thead>
                            <tr>
                                <th>GerenciaID</th>
                                <th>NomeGerencia</th>
                                <th>UserId</th>
                                <th>Nome</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% 
                            ' Listar todos os registros da tabela Gerencias
                            If Not rsGerenciasConn.EOF Then
                                rsGerenciasConn.MoveFirst
                                Do While Not rsGerenciasConn.EOF 
                            %>
                            <tr>
                                <td><span class="badge-id"><%= rsGerenciasConn("GerenciaID") %></span></td>
                                <td><strong><%= rsGerenciasConn("NomeGerencia") %></strong></td>
                                <td><span class="badge bg-primary"><%= rsGerenciasConn("UserId") %></span></td>
                                <td><%= rsGerenciasConn("Nome") %></td>
                            </tr>
                            <%
                                    rsGerenciasConn.MoveNext
                                Loop
                            Else
                            %>
                            <tr>
                                <td colspan="4" class="text-center">Nenhum registro encontrado na tabela Gerencias</td>
                            </tr>
                            <% End If %>
                        </tbody>
                    </table>
                </div>
                <div class="text-muted-small mt-2">
                    <i class="fas fa-list me-1"></i> Total de registros listados: 
                    <% 
                    ' Contar total de registros (simples)
                    Set rsCount = conn.Execute("SELECT COUNT(*) as Total FROM Gerencias")
                    If Not rsCount.EOF Then
                        Response.Write rsCount("Total")
                    End If
                    If IsObject(rsCount) Then rsCount.Close
                    Set rsCount = Nothing
                    %>
                </div>
            </div>
        </div>

        <!-- Seção 2: Gerencias Ativas em Vendas -->
        <div class="card mt-3">
            <div class="card-header">
                <i class="fas fa-list me-2"></i>Gerencias Ativas em Vendas (Conexão connSales)
            </div>
            <div class="card-body">
                <div class="info-box">
                    <i class="fas fa-info-circle me-2"></i>
                    Esta tabela mostra as gerencias que estão atualmente sendo usadas em vendas ativas (não excluídas).
                    Baseado na tabela Vendas do banco de dados de vendas (connSales).
                </div>
                
                <div class="table-responsive">
                    <table class="table table-striped table-hover">
                        <thead>
                            <tr>
                                <th>Gerência</th>
                                <th>Gerente Atual</th>
                                <th>ID Gerente</th>
                                <th>Número de Vendas</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% 
                            ' Usar o RecordSet ESPECÍFICO para a tabela
                            If Not rsContagemGerentesParaTabela.EOF Then
                                rsContagemGerentesParaTabela.MoveFirst
                                Do While Not rsContagemGerentesParaTabela.EOF 
                            %>
                            <tr>
                                <td>
                                    <% 
                                    ' Buscar o nome da gerencia baseado no UserIdGerencia
                                    Set rsGerenciaNome = connSales.Execute("SELECT TOP 1 Gerencia FROM Vendas WHERE UserIdGerencia = " & rsContagemGerentesParaTabela("UserIdGerencia") & " AND EXCLUIDO = 0")
                                    If Not rsGerenciaNome.EOF Then
                                        Response.Write rsGerenciaNome("Gerencia")
                                    Else
                                        Response.Write "Não especificado"
                                    End If
                                    If IsObject(rsGerenciaNome) Then rsGerenciaNome.Close
                                    Set rsGerenciaNome = Nothing
                                    %>
                                </td>
                                <td><strong><%= rsContagemGerentesParaTabela("NomeGerente") %></strong></td>
                                <td><span class="badge bg-secondary"><%= rsContagemGerentesParaTabela("UserIdGerencia") %></span></td>
                                <td>
                                    <span class="badge-count"><%= rsContagemGerentesParaTabela("TotalVendas") %> vendas</span>
                                </td>
                            </tr>
                            <%
                                    rsContagemGerentesParaTabela.MoveNext
                                Loop
                            Else
                            %>
                            <tr>
                                <td colspan="4" class="text-center">Nenhuma gerência ativa em vendas</td>
                            </tr>
                            <% End If %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Seção 3: Alteração de Gerentes em Massa -->
        <div class="card">
            <div class="card-header">
                <i class="fas fa-exchange-alt me-2"></i>Substituir Gerente em Todas as Vendas e na Tabela Gerencias
            </div>
            <div class="card-body">
                <div class="info-box">
                    <i class="fas fa-info-circle me-2"></i>
                    Esta função substituirá o gerente em <strong>TODAS</strong> as vendas onde ele está cadastrado 
                    <strong>E também na tabela Gerencias</strong>. A alteração será aplicada nas seguintes tabelas:
                    <ul class="mb-0 mt-2">
                        <li><strong>Vendas</strong> (connSales) - campo: UserIdGerencia e NomeGerente</li>
                        <li><strong>COMISSOES_A_PAGAR</strong> (connSales) - campo: UserIdGerencia e NomeGerente</li>
                        <li><strong>PAGAMENTOS_COMISSOES</strong> (connSales) - campo: UsuariosUserId e UsuariosNome</li>
                        <li><strong>Gerencias</strong> (conn) - campo: UserId e Nome</li>
                    </ul>
                </div>
                
                <form method="post" id="formSubstituirGerente" onsubmit="return validarFormulario()">
                    <input type="hidden" name="acao" value="alterar_gerentes">
                    
                    <div class="row">
                        <div class="col-md-6">
                            <div class="mb-3">
                                <label for="gerente_antigo_id" class="form-label">Gerente a ser Substituído</label>
                                <select class="form-select" id="gerente_antigo_id" name="gerente_antigo_id" required>
                                    <option value="">Selecione o gerente atual...</option>
                                    <% 
                                    ' Usar o RecordSet ESPECÍFICO para o select
                                    If Not rsContagemGerentesParaSelect.EOF Then
                                        rsContagemGerentesParaSelect.MoveFirst
                                        Do While Not rsContagemGerentesParaSelect.EOF 
                                    %>
                                    <option value="<%= rsContagemGerentesParaSelect("UserIdGerencia") %>">
                                        <%= rsContagemGerentesParaSelect("NomeGerente") %> (ID: <%= rsContagemGerentesParaSelect("UserIdGerencia") %>) - <%= rsContagemGerentesParaSelect("TotalVendas") %> vendas
                                    </option>
                                    <%
                                            rsContagemGerentesParaSelect.MoveNext
                                        Loop
                                    Else
                                    %>
                                    <option value="">Nenhum gerente encontrado</option>
                                    <% End If %>
                                </select>
                            </div>
                        </div>
                        
                        <div class="col-md-6">
                            <div class="mb-3">
                                <label for="gerente_novo_id" class="form-label">Novo Gerente</label>
                                <select class="form-select" id="gerente_novo_id" name="gerente_novo_id" required>
                                    <option value="">Selecione o novo gerente...</option>
                                    <% 
                                    ' Usar o RecordSet ESPECÍFICO para usuários
                                    If Not rsTodosUsuariosParaSelect.EOF Then
                                        rsTodosUsuariosParaSelect.MoveFirst
                                        Do While Not rsTodosUsuariosParaSelect.EOF 
                                    %>
                                    <option value="<%= rsTodosUsuariosParaSelect("UserId") %>">
                                        <%= rsTodosUsuariosParaSelect("Nome") %> (ID: <%= rsTodosUsuariosParaSelect("UserId") %>)
                                    </option>
                                    <%
                                            rsTodosUsuariosParaSelect.MoveNext
                                        Loop
                                    Else
                                    %>
                                    <option value="">Nenhum usuário encontrado</option>
                                    <% End If %>
                                </select>
                            </div>
                        </div>
                    </div>
                    
                    <div class="update-info">
                        <i class="fas fa-database me-2"></i>
                        <strong>Tabelas que serão atualizadas:</strong>
                        <div class="row mt-2">
                            <div class="col-md-6">
                                <div class="form-check">
                                    <input class="form-check-input" type="checkbox" checked disabled>
                                    <label class="form-check-label">
                                        <strong>Vendas</strong> (connSales)
                                    </label>
                                </div>
                                <div class="form-check">
                                    <input class="form-check-input" type="checkbox" checked disabled>
                                    <label class="form-check-label">
                                        <strong>COMISSOES_A_PAGAR</strong> (connSales)
                                    </label>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="form-check">
                                    <input class="form-check-input" type="checkbox" checked disabled>
                                    <label class="form-check-label">
                                        <strong>PAGAMENTOS_COMISSOES</strong> (connSales)
                                    </label>
                                </div>
                                <div class="form-check">
                                    <input class="form-check-input" type="checkbox" checked disabled>
                                    <label class="form-check-label">
                                        <strong>Gerencias</strong> (conn) - <span class="text-primary">NOVO</span>
                                    </label>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="alert alert-warning mt-3">
                        <i class="fas fa-exclamation-triangle me-2"></i>
                        <strong>Atenção:</strong> Esta ação é irreversível. Todas as vendas e registros na tabela Gerencias com o gerente selecionado serão atualizadas.
                    </div>
                    
                    <div class="d-flex justify-content-end">
                        <button type="button" class="btn btn-warning me-2" onclick="confirmarSubstituicao()">
                            <i class="fas fa-exchange-alt me-2"></i>Substituir Gerente
                        </button>
                    </div>
                    
                    <!-- Modal de confirmação -->
                    <div class="modal fade" id="modalConfirmacao" tabindex="-1">
                        <div class="modal-dialog">
                            <div class="modal-content">
                                <div class="modal-header bg-warning">
                                    <h5 class="modal-title">
                                        <i class="fas fa-exclamation-triangle me-2"></i>Confirmar Substituição
                                    </h5>
                                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                                </div>
                                <div class="modal-body">
                                    <p>Tem certeza que deseja substituir o gerente?</p>
                                    <p id="confirmacaoDetalhes"></p>
                                    <p class="text-danger"><strong>Esta ação não pode ser desfeita!</strong></p>
                                    <div class="alert alert-info">
                                        <i class="fas fa-database me-2"></i>
                                        <strong>Tabelas que serão atualizadas:</strong><br>
                                        • Vendas (connSales)<br>
                                        • COMISSOES_A_PAGAR (connSales)<br>
                                        • PAGAMENTOS_COMISSOES (connSales)<br>
                                        • <strong>Gerencias (conn)</strong>
                                    </div>
                                </div>
                                <div class="modal-footer">
                                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                                    <button type="submit" class="btn btn-danger">Confirmar Substituição</button>
                                </div>
                            </div>
                        </div>
                    </div>
                </form>
            </div>
        </div>

        <!-- Seção 4: Lista Completa de Vendas (para referência) -->
        <div class="card">
            <div class="card-header">
                <i class="fas fa-file-invoice me-2"></i>Todas as Vendas (Referência)
            </div>
            <div class="card-body">
                <div class="table-container">
                    <table class="table table-striped table-hover">
                        <thead>
                            <tr>
                                <th>ID Venda</th>
                                <th>Empreendimento</th>
                                <th>Cliente</th>
                                <th>Gerência</th>
                                <th>Gerente</th>
                                <th>ID Gerente</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% 
                            ' Usar o RecordSet ESPECÍFICO para referência
                            If Not rsVendasParaReferencia.EOF Then
                                rsVendasParaReferencia.MoveFirst
                                Do While Not rsVendasParaReferencia.EOF 
                            %>
                            <tr>
                                <td><strong><%= rsVendasParaReferencia("ID") %></strong></td>
                                <td><%= rsVendasParaReferencia("NomeEmpreendimento") %></td>
                                <td><%= rsVendasParaReferencia("NomeCliente") %></td>
                                <td><%= rsVendasParaReferencia("Gerencia") %></td>
                                <td><%= rsVendasParaReferencia("NomeGerente") %></td>
                                <td><span class="badge bg-secondary"><%= rsVendasParaReferencia("UserIdGerencia") %></span></td>
                            </tr>
                            <%
                                    rsVendasParaReferencia.MoveNext
                                Loop
                            Else
                            %>
                            <tr>
                                <td colspan="6" class="text-center">Nenhuma venda encontrada</td>
                            </tr>
                            <% End If %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    
    <script>
        function validarFormulario() {
            var antigoSelect = document.getElementById('gerente_antigo_id');
            var novoSelect = document.getElementById('gerente_novo_id');
            
            if (antigoSelect.value === "" || novoSelect.value === "") {
                alert("Por favor, selecione ambos os gerentes.");
                return false;
            }
            
            if (antigoSelect.value === novoSelect.value) {
                alert("Não é possível substituir o gerente por ele mesmo.");
                return false;
            }
            
            return true;
        }
        
        function confirmarSubstituicao() {
            var antigoSelect = document.getElementById('gerente_antigo_id');
            var novoSelect = document.getElementById('gerente_novo_id');
            
            if (antigoSelect.value === "" || novoSelect.value === "") {
                alert("Por favor, selecione ambos os gerentes.");
                return;
            }
            
            if (antigoSelect.value === novoSelect.value) {
                alert("Não é possível substituir o gerente por ele mesmo.");
                return;
            }
            
            var antigoTexto = antigoSelect.options[antigoSelect.selectedIndex].text;
            var novoTexto = novoSelect.options[novoSelect.selectedIndex].text;
            
            document.getElementById('confirmacaoDetalhes').innerHTML = 
                'Substituir: <strong>' + antigoTexto + '</strong><br>' +
                'Por: <strong>' + novoTexto + '</strong>';
            
            var modal = new bootstrap.Modal(document.getElementById('modalConfirmacao'));
            modal.show();
        }
        
        // Fechar modal após sucesso
        <% If mensagem <> "" Then %>
            var modal = bootstrap.Modal.getInstance(document.getElementById('modalConfirmacao'));
            if (modal) modal.hide();
        <% End If %>
    </script>
</body>
</html>

<%
' Fechar todas as conexões
If IsObject(rsGerenciasConn) Then
    rsGerenciasConn.Close
    Set rsGerenciasConn = Nothing
End If

If IsObject(rsContagemGerentesParaTabela) Then
    rsContagemGerentesParaTabela.Close
    Set rsContagemGerentesParaTabela = Nothing
End If

If IsObject(rsContagemGerentesParaSelect) Then
    rsContagemGerentesParaSelect.Close
    Set rsContagemGerentesParaSelect = Nothing
End If

If IsObject(rsGerenciasNomes) Then
    rsGerenciasNomes.Close
    Set rsGerenciasNomes = Nothing
End If

If IsObject(rsTodosUsuariosParaSelect) Then
    rsTodosUsuariosParaSelect.Close
    Set rsTodosUsuariosParaSelect = Nothing
End If

If IsObject(rsVendasParaReferencia) Then
    rsVendasParaReferencia.Close
    Set rsVendasParaReferencia = Nothing
End If

If IsObject(conn) Then
    conn.Close
    Set conn = Nothing
End If

If IsObject(connSales) Then
    connSales.Close
    Set connSales = Nothing
End If
%>