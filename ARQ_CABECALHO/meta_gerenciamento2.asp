<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 18/12/2025               -->
<!-- CODIGO_ARQUIVO: LPMGCFXNZR          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="conexao.asp"-->           
<!--#include file="conSunSales.asp"-->       
<!--#include file="usr_acoes_v4GVendas.inc"-->

<%
' =============================================
'      META – GERENCIAMENTO (BLOCO ANTES DO HTML)
' =============================================
Response.CodePage = 65001
Response.Charset = "utf-8"

' -----------------------------
'  CONSTANTES
' -----------------------------
Const ANO_MINIMO = 2026

' -----------------------------
'  ABRIR CONEXÕES
' -----------------------------
Dim connOrg, connSales
Set connOrg = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")

' strConn (para BD organizacional) deve vir de conexao.asp
' strConnSales (para BD de vendas/metas) deve vir de conSunSales.asp
connOrg.Open strConn
connSales.Open strConnSales

' -----------------------------
'  RECEBER ANO E MÊS (GET/POST)
' -----------------------------
Dim anoSelecionado, mesSelecionado
anoSelecionado = Request("ano")
mesSelecionado = Request("mes")

If anoSelecionado = "" Then anoSelecionado = Year(Now)
If mesSelecionado = "" Then mesSelecionado = Month(Now)

' Ajuste para mínimo, se necessário
If IsNumeric(anoSelecionado) Then
    If CInt(anoSelecionado) < ANO_MINIMO Then
        anoSelecionado = ANO_MINIMO
    End If
Else
    anoSelecionado = ANO_MINIMO
End If

' -----------------------------
'  PREPARAR DICIONÁRIOS
' -----------------------------
Dim rs, sql

' Dicionário de gerencias (vindas do banco organizacional)
Dim dicGerencias
Set dicGerencias = Server.CreateObject("Scripting.Dictionary")

' Dicionário de relação gerencia->diretoria
Dim dicGerenciaDiretoria
Set dicGerenciaDiretoria = Server.CreateObject("Scripting.Dictionary")

sql = "SELECT GerenciaID, NomeGerencia, DiretoriaID FROM Gerencias ORDER BY NomeGerencia"
Set rs = connOrg.Execute(sql)
Do While Not rs.EOF
    dicGerencias.Add CStr(rs("GerenciaID")), rs("NomeGerencia")
    dicGerenciaDiretoria.Add CStr(rs("GerenciaID")), CInt(rs("DiretoriaID"))
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

' Dicionário de metas carregadas (do banco de sales)
Dim dicMetas
Set dicMetas = Server.CreateObject("Scripting.Dictionary")

sql = "SELECT GerenciaID, ValorMeta FROM MetaGerencia " & _
      "WHERE Ano = " & anoSelecionado & " AND Mes = " & mesSelecionado

Set rs = connSales.Execute(sql)
Do While Not rs.EOF
    If Not dicMetas.Exists(CStr(rs("GerenciaID"))) Then
        dicMetas.Add CStr(rs("GerenciaID")), CDbl(rs("ValorMeta"))
    End If
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

' -----------------------------
'  FUNÇÃO: ATUALIZAR META DIRETORIA
' -----------------------------
' -----------------------------
'  FUNÇÃO: ATUALIZAR META DIRETORIA (VERSÃO FINAL)
' -----------------------------
Sub AtualizarMetaDiretoria(diretoriaID, ano, mes, usuario)
    On Error Resume Next

    Call RegistrarLog("=== Atualizando MetaDiretoria para Diretoria " & diretoriaID & " ===")

    Dim sqlGerencias, rsGer, listaGerencias
    listaGerencias = ""

    ' ---------------------------------------------------------
    ' 1. Buscar TODAS as gerências vinculadas à diretoria
    ' ---------------------------------------------------------
    sqlGerencias = "SELECT IdGerencia FROM Gerencia WHERE IdDiretoria=" & diretoriaID

    Set rsGer = connSales.Execute(sqlGerencias)

    Do Until rsGer.EOF
        If listaGerencias <> "" Then listaGerencias = listaGerencias & ","
        listaGerencias = listaGerencias & rsGer("IdGerencia")
        rsGer.MoveNext
    Loop

    rsGer.Close
    Set rsGer = Nothing

    If listaGerencias = "" Then
        Call RegistrarLog("Nenhuma gerência encontrada para a diretoria " & diretoriaID)
        Exit Sub
    End If

    Call RegistrarLog("Gerências da diretoria: " & listaGerencias)

    ' ---------------------------------------------------------
    ' 2. Rodar SUA CONSULTA — AGRUPAMENTO REAL por Gerência/Ano/Mês
    ' ---------------------------------------------------------
    Dim sqlSoma, rsSoma, totalDiretoria
    totalDiretoria = 0

    sqlSoma = "SELECT Sum(ValorMeta) AS SomaTotal " & _
              "FROM MetaGerencia " & _
              "WHERE GerenciaID IN (" & listaGerencias & ") " & _
              "AND Ano=" & ano & " AND Mes=" & mes & ";"

    Call RegistrarLog("SQL Soma Diretoria => " & sqlSoma)

    Set rsSoma = connSales.Execute(sqlSoma)

    If Not rsSoma.EOF And Not IsNull(rsSoma("SomaTotal")) Then
        totalDiretoria = CDbl(rsSoma("SomaTotal"))
    End If

    rsSoma.Close
    Set rsSoma = Nothing

    Call RegistrarLog("Total calculado diretoria=" & diretoriaID & " -> " & totalDiretoria)

    ' ---------------------------------------------------------
    ' 3. Consultar MetaDiretoria existente
    ' ---------------------------------------------------------
    Dim sqlCheck, rsCheck, sqlAcao

    sqlCheck = "SELECT MetaDir_ID FROM MetaDiretoria " & _
               "WHERE DiretoriaID=" & diretoriaID & " AND Ano=" & ano & " AND Mes=" & mes

    Set rsCheck = Server.CreateObject("ADODB.Recordset")
    rsCheck.Open sqlCheck, connSales, 1, 3

    ' ---------------------------------------------------------
    ' 4. INSERT ou UPDATE automaticamente
    ' ---------------------------------------------------------

    If rsCheck.EOF Then
        ' ---------------------------
        ' NÃO EXISTE → INSERT
        ' ---------------------------
        If totalDiretoria > 0 Then
            sqlAcao = "INSERT INTO MetaDiretoria (DiretoriaID, Ano, Mes, TotalMetas, Usuario, DataHora) VALUES (" & _
                      diretoriaID & ", " & ano & ", " & mes & ", " & _
                      Replace(CStr(totalDiretoria), ",", ".") & ", '" & _
                      Replace(usuario, "'", "''") & "', Now())"

            Call RegistrarLog("INSERT MetaDiretoria => " & sqlAcao)
            connSales.Execute sqlAcao

        Else
            Call RegistrarLog("Sem metas válidas — nada a inserir.")
        End If

    Else
        ' ---------------------------
        ' EXISTE → UPDATE
        ' ---------------------------
        sqlAcao = "UPDATE MetaDiretoria SET TotalMetas=" & _
                  Replace(CStr(totalDiretoria), ",", ".") & _
                  ", Usuario='" & Replace(usuario, "'", "''") & "', DataHora=Now() " & _
                  "WHERE MetaDir_ID=" & rsCheck("MetaDir_ID")

        Call RegistrarLog("UPDATE MetaDiretoria => " & sqlAcao)
        connSales.Execute sqlAcao

    End If

    rsCheck.Close
    Set rsCheck = Nothing

    Call RegistrarLog("=== FIM Atualização MetaDiretoria Diretoria " & diretoriaID & " ===")

    On Error GoTo 0
End Sub

' -----------------------------
'  FUNÇÃO: ATUALIZAR META EMPRESA
' -----------------------------
Sub AtualizarMetaEmpresa(ano, mes, usuario)
    On Error Resume Next
    
    Call RegistrarLog("Atualizando meta da empresa...")
    
    ' Calcular soma total de todas as diretorias
    Dim sqlTotal, rsTotal, totalEmpresa
    sqlTotal = "SELECT SUM(TotalMetas) AS Total FROM MetaDiretoria " & _
               "WHERE Ano = " & ano & " AND Mes = " & mes & " AND TotalMetas > 0"
    
    Call RegistrarLog("SQL Soma Empresa: " & sqlTotal)
    
    Set rsTotal = Server.CreateObject("ADODB.Recordset")
    rsTotal.Open sqlTotal, connSales, 1, 3
    
    totalEmpresa = 0
    If Not rsTotal.EOF And Not IsNull(rsTotal("Total")) Then
        totalEmpresa = rsTotal("Total")
    End If
    
    rsTotal.Close
    Set rsTotal = Nothing
    
    Call RegistrarLog("Total empresa calculado: " & totalEmpresa)
    
    ' Verificar se já existe registro
    Dim sqlCheckEmp, rsCheckEmp, sqlAcaoEmp
    sqlCheckEmp = "SELECT MetaEmp_ID FROM MetaEmpresa WHERE Ano = " & ano & " AND Mes = " & mes
    
    Set rsCheckEmp = Server.CreateObject("ADODB.Recordset")
    rsCheckEmp.Open sqlCheckEmp, connSales, 1, 3
    
    If rsCheckEmp.EOF Then
        ' INSERT se houver meta
        If totalEmpresa > 0 Then
            sqlAcaoEmp = "INSERT INTO MetaEmpresa (Ano, Mes, Meta, DataAtualizacao) VALUES (" & _
                        ano & ", " & mes & ", " & totalEmpresa & ", Now())"
            
            Call RegistrarLog("INSERT MetaEmpresa: " & sqlAcaoEmp)
            connSales.Execute sqlAcaoEmp
        End If
    Else
        ' UPDATE ou DELETE
        If totalEmpresa > 0 Then
            sqlAcaoEmp = "UPDATE MetaEmpresa SET Meta = " & totalEmpresa & ", DataAtualizacao = Now() " & _
                        "WHERE MetaEmp_ID = " & rsCheckEmp("MetaEmp_ID")
            
            Call RegistrarLog("UPDATE MetaEmpresa: " & sqlAcaoEmp)
            connSales.Execute sqlAcaoEmp
        Else
            sqlAcaoEmp = "DELETE FROM MetaEmpresa WHERE MetaEmp_ID = " & rsCheckEmp("MetaEmp_ID")
            
            Call RegistrarLog("DELETE MetaEmpresa: " & sqlAcaoEmp)
            connSales.Execute sqlAcaoEmp
        End If
    End If
    
    rsCheckEmp.Close
    Set rsCheckEmp = Nothing
    
    Call RegistrarLog("Meta da empresa atualizada com sucesso")
    
    On Error GoTo 0
End Sub

' -----------------------------
'  PROCESSAR POST (SALVAR METAS)
' -----------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    ' Registrar início no log
    Call RegistrarLog("=== INÍCIO SALVAMENTO METAS ===")
    Call RegistrarLog("ANO: " & anoSelecionado & " | MES: " & mesSelecionado & " | USUARIO: " & Session("Usuario"))

    ' Iniciar transação na conexão de vendas
    On Error Resume Next
    connSales.BeginTrans

    Dim totalInseridos
    totalInseridos = 0
    
    ' Lista para armazenar diretorias afetadas
    Dim diretoriasAfetadas
    Set diretoriasAfetadas = Server.CreateObject("Scripting.Dictionary")

    ' Percorrer campos do form
    Dim key, valorRaw, valorConv, gerID

    For Each key In Request.Form
        ' aceitar nomes tipo "gerencia123" ou "gerencia_123" conforme seu html
        If Left(key,8) = "gerencia" Then
            valorRaw = Trim(Request.Form(key))
            If valorRaw <> "" Then
                ' normalizar: remover pontos e trocar vírgula por ponto
                valorConv = Replace(valorRaw, ".", "")
                valorConv = Replace(valorConv, ",", ".")
                
                ' Verificar se precisa dividir por 100 (se for entrada com centavos)
                If InStr(valorRaw, ",") > 0 Then
                    ' Já foi convertido acima, manter como está
                Else
                    ' Se não tinha vírgula, provavelmente já está em formato correto
                End If
                
                If IsNumeric(valorConv) Then
                    If valorConv > 0 Then
                        ' extrair ID (suporta gerencia123 e gerencia_123)
                        gerID = Mid(key, 9)
                        If Left(gerID,1) = "_" Then gerID = Mid(gerID,2)
                        
                        ' garantir numeric
                        If IsNumeric(gerID) Then
                            gerID = CInt(gerID)
                            
                            ' Obter diretoria desta gerencia
                            Dim dirID
                            If dicGerenciaDiretoria.Exists(CStr(gerID)) Then
                                dirID = dicGerenciaDiretoria(CStr(gerID))
                                
                                ' Adicionar à lista de diretorias afetadas
                                If Not diretoriasAfetadas.Exists(CStr(dirID)) Then
                                    diretoriasAfetadas.Add CStr(dirID), dirID
                                End If
                            End If
                            
                            ' Verificar se já existe registro
                            Dim sqlCheck, rsCheck, idExist
                            sqlCheck = "SELECT MetaGer_ID FROM MetaGerencia WHERE GerenciaID=" & gerID & _
                                       " AND Ano=" & anoSelecionado & " AND Mes=" & mesSelecionado
                            Set rsCheck = Server.CreateObject("ADODB.Recordset")
                            rsCheck.Open sqlCheck, connSales, 1, 3
                            
                            If rsCheck.EOF Then
                                ' INSERT
                                Dim sqlInsert
                                sqlInsert = "INSERT INTO MetaGerencia (GerenciaID, Ano, Mes, ValorMeta, Usuario, DataHora) VALUES (" & _
                                            gerID & ", " & anoSelecionado & ", " & mesSelecionado & ", " & _
                                            valorConv & ", '" & _
                                            Replace(Session("Usuario"), "'", "''") & "', Now())"
                                
                                Call RegistrarLog("INSERT -> GerenciaID:" & gerID & " Valor:" & valorConv & " Diretoria:" & dirID)
                                Call RegistrarLog("SQLINSERT -> " & sqlInsert)
                                
                                connSales.Execute sqlInsert
                            Else
                                ' UPDATE
                                Dim sqlUpdate
                                sqlUpdate = "UPDATE MetaGerencia SET ValorMeta=" & _
                                            Replace(FormatNumber(CDbl(valorConv), 2, -1, -1, -1), ",", ".") & ", " & _
                                            "Usuario='" & Replace(Session("Usuario"), "'", "''") & "', DataHora=Now() " & _
                                            "WHERE MetaGer_ID=" & rsCheck("MetaGer_ID")
                                            
                                Call RegistrarLog("UPDATE -> GerenciaID:" & gerID & " Valor:" & valorConv & " Diretoria:" & dirID)
                                Call RegistrarLog("SQLUPDATE -> " & sqlUpdate)
                                
                                connSales.Execute sqlUpdate
                            End If
                            
                            If Err.Number <> 0 Then
                                Call RegistrarLog("ERRO SQL (" & Err.Number & "): " & Err.Description)
                                Err.Clear
                            Else
                                totalInseridos = totalInseridos + 1
                            End If
                            
                            rsCheck.Close
                            Set rsCheck = Nothing
                        Else
                            Call RegistrarLog("GerenciaID inválido extraído do campo: " & key)
                        End If
                    Else
                        Call RegistrarLog("Valor <= 0 ignorado para campo: " & key & " valorRaw: " & valorRaw)
                    End If
                Else
                    Call RegistrarLog("Valor não numérico ignorado para campo: " & key & " valorRaw: " & valorRaw)
                End If
            Else
                ' campo vazio - verificar se precisa remover meta existente
                gerID = Mid(key, 9)
                If Left(gerID,1) = "_" Then gerID = Mid(gerID,2)
                
                If IsNumeric(gerID) Then
                    gerID = CInt(gerID)
                    
                    ' Verificar se existe registro para deletar
                    Dim sqlCheckDelete, rsCheckDelete
                    sqlCheckDelete = "SELECT MetaGer_ID FROM MetaGerencia WHERE GerenciaID=" & gerID & _
                                     " AND Ano=" & anoSelecionado & " AND Mes=" & mesSelecionado
                    Set rsCheckDelete = Server.CreateObject("ADODB.Recordset")
                    rsCheckDelete.Open sqlCheckDelete, connSales, 1, 3
                    
                    If Not rsCheckDelete.EOF Then
                        ' Obter diretoria para atualizar
                        Dim dirIDDelete
                        If dicGerenciaDiretoria.Exists(CStr(gerID)) Then
                            dirIDDelete = dicGerenciaDiretoria(CStr(gerID))
                            If Not diretoriasAfetadas.Exists(CStr(dirIDDelete)) Then
                                diretoriasAfetadas.Add CStr(dirIDDelete), dirIDDelete
                            End If
                        End If
                        
                        ' Deletar registro
                        Dim sqlDelete
                        sqlDelete = "DELETE FROM MetaGerencia WHERE MetaGer_ID=" & rsCheckDelete("MetaGer_ID")
                        
                        Call RegistrarLog("DELETE -> GerenciaID:" & gerID & " (campo vazio)")
                        Call RegistrarLog("SQLDELETE -> " & sqlDelete)
                        
                        connSales.Execute sqlDelete
                        
                        If Err.Number <> 0 Then
                            Call RegistrarLog("ERRO DELETE (" & Err.Number & "): " & Err.Description)
                            Err.Clear
                        End If
                    End If
                    
                    rsCheckDelete.Close
                    Set rsCheckDelete = Nothing
                End If
            End If
        End If
    Next

    ' Atualizar as diretorias afetadas
    Call RegistrarLog("Diretorias afetadas: " & diretoriasAfetadas.Count)
    
    For Each dirKey In diretoriasAfetadas.Keys
        Dim dirIDUpdate
        dirIDUpdate = diretoriasAfetadas(dirKey)
        Call AtualizarMetaDiretoria(dirIDUpdate, anoSelecionado, mesSelecionado, Session("Usuario"))
    Next
    
    Set diretoriasAfetadas = Nothing

    ' Commit ou rollback conforme erro
    If Err.Number = 0 Then
        connSales.CommitTrans
        Call RegistrarLog("COMMIT - Total processado: " & totalInseridos)
        
        ' Mostrar mensagem de sucesso
        Session("msg_sucesso") = "Metas salvas com sucesso! (" & totalInseridos & " gerencias atualizadas)"
    Else
        connSales.RollbackTrans
        Call RegistrarLog("ROLLBACK por erro. Err: " & Err.Number & " - " & Err.Description)
        
        ' Mostrar mensagem de erro
        Session("msg_erro") = "Erro ao salvar metas: " & Err.Description
        
        Err.Clear
    End If

    Call RegistrarLog("=== FIM SALVAMENTO METAS ===")
    
    ' redirecionar para evitar reenvio
    Response.Redirect "meta_gerenciamento2.asp?ano=" & anoSelecionado & "&mes=" & mesSelecionado & "&saved=1"
End If

' -----------------------------
'  FUNÇÃO DE LOG em metas.log
' -----------------------------
Sub RegistrarLog(msg)
    On Error Resume Next
    Dim fso, arq, caminho
    caminho = Server.MapPath("metas.log")
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(caminho) Then
        Set arq = fso.OpenTextFile(caminho, 8, True) ' append
    Else
        Set arq = fso.CreateTextFile(caminho, True)
    End If
    arq.WriteLine Now() & " - " & msg
    arq.Close
    Set arq = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Sub

' -----------------------------
'  OBS: As conexões serão fechadas no final da página (depois do HTML)
' -----------------------------
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gerenciamento de Metas</title>
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
        .diretoria-card {
            border-left: 5px solid var(--primary-color);
        }
        .meta-input {
            font-weight: 600;
            text-align: right;
        }
        .loading-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.5);
            z-index: 9999;
            justify-content: center;
            align-items: center;
        }
        .spinner {
            width: 50px;
            height: 50px;
            border: 5px solid #f3f3f3;
            border-top: 5px solid var(--primary-color);
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .campo-vazio {
            border-color: #ffc107 !important;
            background-color: #fff3cd !important;
        }
        .campo-valido {
            border-color: #198754 !important;
            background-color: #d1e7dd !important;
        }
    </style>
</head>
<body>
    <!-- Loading Overlay -->
    <div class="loading-overlay" id="loadingOverlay">
        <div class="spinner"></div>
        <div class="text-white mt-3">Processando, aguarde...</div>
    </div>

    <div class="container-fluid py-4">
        <!-- Cabeçalho -->
        <div class="row mb-4">
            <div class="col-12">
                <div class="card shadow">
                    <div class="card-header card-header-custom">
                        <div class="d-flex justify-content-between align-items-center">
                            <div>
                                <h4 class="mb-0"><i class="bi bi-bullseye me-2"></i>Sistema de Metas</h4>
                                <small>Gestão hierárquica de metas organizacionais</small>
                            </div>
                            <div>
                                <span class="badge bg-light text-dark">
                                    <i class="bi bi-calendar me-1"></i>
                                    <%
                                    Dim mesesArray
                                    mesesArray = Array("", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro")
                                    
                                    If mesSelecionado <> "" Then
                                        Response.Write mesesArray(CInt(mesSelecionado)) & "/" & anoSelecionado
                                    Else
                                        Response.Write "Selecione um período"
                                    End If
                                    %>
                                </span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Mensagens -->
        <%
        If Session("msg_sucesso") <> "" Then
            Response.Write "<div class='alert alert-success alert-dismissible fade show mb-4'>" & _
                         "<i class='bi bi-check-circle me-2'></i>" & Session("msg_sucesso") & _
                         "<button type='button' class='btn-close' data-bs-dismiss='alert'></button></div>"
            Session("msg_sucesso") = ""
        End If
        
        If Session("msg_erro") <> "" Then
            Response.Write "<div class='alert alert-danger alert-dismissible fade show mb-4'>" & _
                         "<i class='bi bi-exclamation-triangle me-2'></i>" & Session("msg_erro") & _
                         "<button type='button' class='btn-close' data-bs-dismiss='alert'></button></div>"
            Session("msg_erro") = ""
        End If
        
        ' Mensagem de sucesso do redirect
        If Request.QueryString("saved") = "1" Then
            If Session("msg_sucesso") = "" And Session("msg_erro") = "" Then
                Response.Write "<div class='alert alert-info alert-dismissible fade show mb-4'>" & _
                             "<i class='bi bi-check-circle me-2'></i>Operação concluída com sucesso!" & _
                             "<button type='button' class='btn-close' data-bs-dismiss='alert'></button></div>"
            End If
        End If
        %>

        <!-- Filtros -->
        <div class="row mb-4">
            <div class="col-12">
                <div class="card shadow">
                    <div class="card-body">
                        <h5 class="card-title mb-3">
                            <i class="bi bi-funnel me-2"></i>Filtrar Período
                        </h5>
                        <form method="GET" class="row g-3" id="filtroForm">
                            <div class="col-md-3">
                                <label class="form-label">Ano</label>
<select name="ano" class="form-select" id="anoSelect" required>
        <option value="">Selecione o ano</option>
        <%
        ' Gerar opções de ano a partir de 2026
        Dim anoInicial, anoFinal
        anoInicial = 2026
        anoFinal = Year(Now) + 1 ' Até 3 anos no futuro
        
        For i = anoInicial to anoFinal
            Response.Write "<option value='" & i & "'"
            If CStr(i) = CStr(anoSelecionado) Then
                Response.Write " selected"
            End If
            Response.Write ">" & i & "</option>"
        Next
        %>
    </select>
                            </div>
                            <div class="col-md-3">
                                <label class="form-label">Mês</label>
                                <select name="mes" class="form-select" id="mesSelect">
                                    <option value="">Selecione o mês</option>
                                    <%
                                    For i = 1 to 12
                                        Response.Write "<option value='" & i & "'"
                                        If CStr(i) = CStr(mesSelecionado) Then
                                            Response.Write " selected"
                                        End If
                                        Response.Write ">" & mesesArray(i) & "</option>"
                                    Next
                                    %>
                                </select>
                            </div>
                            <div class="col-md-3">
                                <label class="form-label">Diretoria</label>
                                <select name="diretoria" class="form-select">
                                    <option value="">Todas as Diretorias</option>
                                    <%
                                    sql = "SELECT * FROM Diretorias ORDER BY NomeDiretoria"
                                    Set rs = connOrg.Execute(sql)
                                    
                                    Do While Not rs.EOF
                                        Response.Write "<option value='" & rs("DiretoriaID") & "'"
                                        If CStr(rs("DiretoriaID")) = Request.QueryString("diretoria") Then
                                            Response.Write " selected"
                                        End If
                                        Response.Write ">" & rs("NomeDiretoria") & "</option>"
                                        rs.MoveNext
                                    Loop
                                    rs.Close
                                    %>
                                </select>
                            </div>
                            <div class="col-md-3 d-flex align-items-end">
                                <div class="d-grid gap-2 d-md-flex">
                                    <button type="submit" class="btn btn-primary">
                                        <i class="bi bi-search me-1"></i>Filtrar
                                    </button>
                                    <a href="meta_historico.asp" class="btn btn-outline-secondary">
                                        <i class="bi bi-clock-history me-1"></i>Histórico
                                    </a>
                                </div>
                            </div>
                        </form>
                    </div>
                </div>
            </div>
        </div>

        <!-- Formulário de Metas -->
        <%
        If mesSelecionado <> "" Then
        %>
        <form method="POST" action="meta_gerenciamento2.asp" id="formMetas">
            <!-- Campos hidden com valores atuais -->
            <input type="hidden" name="ano" id="anoHidden" value="<%=anoSelecionado%>">
            <input type="hidden" name="mes" id="mesHidden" value="<%=mesSelecionado%>">
            
            <!-- Instruções -->
            <div class="alert alert-info mb-4">
                <i class="bi bi-info-circle me-2"></i>
                <strong>Instruções:</strong> Preencha apenas os campos onde deseja definir metas. Campos vazios ou com valor zero serão ignorados. As metas das diretorias e da empresa serão calculadas automaticamente.
            </div>
            
            <!-- Metas por Diretoria -->
            <%
            ' Buscar diretorias com suas gerencias
            Dim filtroDiretoriaSQL
            filtroDiretoriaSQL = ""
            If Request.QueryString("diretoria") <> "" Then
                filtroDiretoriaSQL = " WHERE d.DiretoriaID = " & Request.QueryString("diretoria")
            End If
            
            sql = "SELECT d.DiretoriaID, d.NomeDiretoria, " & _
                  "g.GerenciaID, g.NomeGerencia " & _
                  "FROM Diretorias d " & _
                  "INNER JOIN Gerencias g ON d.DiretoriaID = g.DiretoriaID " & _
                  filtroDiretoriaSQL & _
                  " ORDER BY d.NomeDiretoria, g.NomeGerencia"
            
            Set rsDiretorias = Server.CreateObject("ADODB.Recordset")
            rsDiretorias.Open sql, connOrg, 1, 3
            
            Dim currentDiretoriaID, diretoriaTotal
            currentDiretoriaID = 0
            diretoriaTotal = 0
            
            Do While Not rsDiretorias.EOF
                If currentDiretoriaID <> rsDiretorias("DiretoriaID") Then
                    If currentDiretoriaID > 0 Then
                        ' Exibir total da diretoria
                        %>
                        </div>
                        <div class="card-footer bg-light">
                            <div class="row align-items-center">
                                <div class="col-md-8">
                                    <strong>Total da Diretoria:</strong>
                                    <small class="text-muted ms-2">Soma automática das gerencias</small>
                                </div>
                                <div class="col-md-4">
                                    <div class="input-group">
                                        <span class="input-group-text bg-white">
                                            <i class="bi bi-calculator"></i>
                                        </span>
                                        <input type="text" class="form-control bg-white meta-input" value="<%=FormatNumber(diretoriaTotal, 2, -1, -1, -1)%>" readonly>
                                        <span class="input-group-text bg-white">R$</span>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                        <%
                        diretoriaTotal = 0
                    End If
                    
                    currentDiretoriaID = rsDiretorias("DiretoriaID")
                    %>
                    <div class="diretoria-card card shadow mb-4">
                        <div class="card-header card-header-custom">
                            <h5 class="mb-0">
                                <i class="bi bi-building me-2"></i>
                                <%=rsDiretorias("NomeDiretoria")%>
                            </h5>
                        </div>
                        <div class="card-body">
                    <%
                End If
                
                ' Buscar meta atual
                Dim metaAtualGerencia
                metaAtualGerencia = 0
                
                sql = "SELECT ValorMeta FROM MetaGerencia WHERE GerenciaID = " & rsDiretorias("GerenciaID") & _
                      " AND Ano = " & anoSelecionado & " AND Mes = " & mesSelecionado
                
                Set rsMeta = Server.CreateObject("ADODB.Recordset")
                rsMeta.Open sql, connSales, 1, 3
                
                If Not rsMeta.EOF And Not IsNull(rsMeta("ValorMeta")) Then
                    metaAtualGerencia = CDbl(rsMeta("ValorMeta"))
                End If
                rsMeta.Close
                Set rsMeta = Nothing
                
                If metaAtualGerencia > 0 Then
                    diretoriaTotal = diretoriaTotal + metaAtualGerencia
                End If
                
                ' Determinar classe CSS baseada no valor
                Dim classeInput
                If metaAtualGerencia > 0 Then
                    classeInput = "campo-valido"
                ElseIf metaAtualGerencia = 0 Then
                    classeInput = "campo-vazio"
                Else
                    classeInput = ""
                End If
                %>
                <div class="row align-items-center mb-2">
                    <div class="col-md-5">
                        <label class="form-label mb-0">
                            <i class="bi bi-diagram-2 me-2 text-primary"></i>
                            <%=rsDiretorias("NomeGerencia")%>
                            <% If metaAtualGerencia > 0 Then %>
                            <span class="badge bg-success ms-2">✓ Definida</span>
                            <% End If %>
                        </label>
                    </div>
                    <div class="col-md-7">
                        <div class="input-group">
                            <span class="input-group-text">R$</span>
<input type="text" 
       name="gerencia<%=rsDiretorias("GerenciaID")%>"
       class="form-control meta-input <%=classeInput%>"
       value="<% 
                 If metaAtualGerencia > 0 Then 
                     Response.Write FormatNumber(metaAtualGerencia, 2, -1, -1, -1) 
                 Else 
                     Response.Write "" 
                 End If 
              %>"
       placeholder="0,00"
       onkeyup="formatarMoeda(this)"
       onblur="validarCampo(this)"
       data-gerenciaid="<%=rsDiretorias("GerenciaID")%>">
                            <button type="button" class="btn btn-outline-secondary" onclick="limparCampo(this)">
                                <i class="bi bi-x"></i>
                            </button>
                        </div>
                        <div class="form-text">
                            <% If metaAtualGerencia > 0 Then %>
                            <small class="text-success">Meta atual: R$ <%=FormatNumber(metaAtualGerencia, 2, -1, -1, -1)%></small>
                            <% Else %>
                            <small class="text-muted">Deixe vazio para manter sem meta</small>
                            <% End If %>
                        </div>
                    </div>
                </div>
                <%
                rsDiretorias.MoveNext
            Loop
            
            ' Fechar última diretoria
            If currentDiretoriaID > 0 Then
                %>
                        </div>
                        <div class="card-footer bg-light">
                            <div class="row align-items-center">
                                <div class="col-md-8">
                                    <strong>Total da Diretoria:</strong>
                                    <small class="text-muted ms-2">Soma automática das gerencias</small>
                                </div>
                                <div class="col-md-4">
                                    <div class="input-group">
                                        <span class="input-group-text bg-white">
                                            <i class="bi bi-calculator"></i>
                                        </span>
                                        <input type="text" class="form-control bg-white meta-input" value="<%=FormatNumber(diretoriaTotal, 2, -1, -1, -1)%>" readonly>
                                        <span class="input-group-text bg-white">R$</span>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                <%
            End If
            
            rsDiretorias.Close
            Set rsDiretorias = Nothing
            
            If currentDiretoriaID = 0 Then
                Response.Write "<div class='alert alert-warning'>Nenhuma diretoria/gerencia encontrada.</div>"
            End If
            %>
            
            <!-- Botões -->
            <div class="row mt-4">
                <div class="col-12">
                    <div class="card shadow">
                        <div class="card-body">
                            <div class="d-flex justify-content-between align-items-center">
                                <div>
                                    <button type="button" class="btn btn-outline-warning" onclick="limparTodosCampos()">
                                        <i class="bi bi-eraser me-1"></i>Limpar Todos
                                    </button>
                                    <button type="button" class="btn btn-outline-info ms-2" onclick="validarFormulario()">
                                        <i class="bi bi-check2-circle me-1"></i>Validar Antes de Salvar
                                    </button>
                                </div>
                                <div>
                                    <button type="button" class="btn btn-secondary" onclick="window.location.href='meta_gerenciamento2.asp'">
                                        <i class="bi bi-x-circle me-1"></i>Cancelar
                                    </button>
                                    <button type="submit" class="btn btn-success ms-2" id="btnSalvar">
                                        <i class="bi bi-check-circle me-1"></i>Salvar Metas Válidas
                                    </button>
                                </div>
                            </div>
                            <div class="mt-2">
                                <small class="text-muted">
                                    <i class="bi bi-lightbulb me-1"></i>
                                    As diretorias serão atualizadas automaticamente com a soma das respectivas gerencias.
                                </small>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </form>
        <%
        Else
        %>
        <div class="row">
            <div class="col-12">
                <div class="card shadow">
                    <div class="card-body text-center py-5">
                        <i class="bi bi-bullseye text-muted" style="font-size: 4rem;"></i>
                        <h4 class="mt-3">Selecione um período</h4>
                        <p class="text-muted">Escolha um ano e mês para gerenciar as metas.</p>
                    </div>
                </div>
            </div>
        </div>
        <%
        End If
        %>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    
    <script>
// ===============================
// FORMATAR MOEDA
// ===============================
function formatarMoeda(input) {
    let valor = input.value.replace(/\D/g, '');
    if (valor === '') {
        input.value = '';
        validarCampo(input);
        return;
    }
    valor = (parseInt(valor) / 100).toLocaleString('pt-BR', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
    input.value = valor;
    validarCampo(input);
}

// ===============================
// VALIDAR CAMPO INDIVIDUAL
// ===============================
function validarCampo(input) {
    let valor = input.value.replace(/\./g, '').replace(',', '.');

    input.classList.remove('campo-vazio', 'campo-valido', 'is-invalid');

    if (input.value.trim() === '') {
        input.classList.add('campo-vazio');
        return;
    }

    if (isNaN(parseFloat(valor))) {
        input.classList.add('is-invalid');
        return;
    }

    let num = parseFloat(valor);

    if (num > 0) {
        input.classList.add('campo-valido');
    } else {
        input.classList.add('campo-vazio');
    }
}

// ===============================
// LIMPAR CAMPO INDIVIDUAL
// ===============================
function limparCampo(botao) {
    const input = botao.closest('.input-group').querySelector('input');
    if (!input) return;
    input.value = '';
    validarCampo(input);
    input.focus();
}

// ===============================
// LIMPAR TODOS
// ===============================
function limparTodosCampos() {
    if (!confirm("Tem certeza que deseja limpar TODOS os campos?")) return;

    const inputs = document.querySelectorAll('input[name^="gerencia"]');
    inputs.forEach(input => {
        input.value = "";
        validarCampo(input);
    });
}

// ===============================
// VALIDAR FORMULÁRIO COMPLETO
// ===============================
function validarFormulario() {
    const inputs = document.querySelectorAll('input[name^="gerencia"]');

    let validos = 0;
    let vazios = 0;
    let invalidos = 0;

    inputs.forEach(input => {
        let val = input.value.replace(/\./g, '').replace(',', '.');

        if (input.value.trim() === '') {
            vazios++;
        }
        else if (isNaN(parseFloat(val))) {
            invalidos++;
        }
        else if (parseFloat(val) > 0) {
            validos++;
        }
        else {
            vazios++;
        }
    });

    alert(
        "Validação:\n\n" +
        "✓ Válidos (>0): " + validos + "\n" +
        "○ Vazios/Zero: " + vazios + "\n" +
        "✗ Inválidos: " + invalidos + "\n" +
        "\nAtenção: As diretorias serão atualizadas automaticamente com a soma das respectivas gerencias."
    );

    return validos > 0;
}

// ===============================
// MOSTRAR LOADING
// ===============================
function mostrarLoading() {
    const overlay = document.getElementById("loadingOverlay");
    const btn = document.getElementById("btnSalvar");

    if (overlay) overlay.style.display = "flex";
    if (btn) btn.disabled = true;
}

// ===============================
// PROCESSAR SUBMIT
// ===============================
function configurarEnvioFormulario() {
    const form = document.getElementById("formMetas");
    if (!form) return; // evita erro

    form.addEventListener("submit", function (e) {
        e.preventDefault();

        if (!validarFormulario()) {
            alert("Preencha ao menos um campo > 0.");
            return;
        }

        mostrarLoading();

        const inputs = form.querySelectorAll('input[name^="gerencia"]');
        let enviados = 0;

        inputs.forEach(input => {
            let valOriginal = input.value;
            let val = valOriginal.replace(/\./g, '').replace(',', '.');

            if (val === "" || isNaN(parseFloat(val)) || parseFloat(val) <= 0) {
                input.disabled = true; // não enviar
            } else {
                input.value = parseFloat(val).toFixed(2);
                enviados++;
            }
        });

        if (enviados === 0) {
            alert("Nenhum valor válido encontrado!");
            inputs.forEach(i => i.disabled = false);
            document.getElementById("loadingOverlay").style.display = "none";
            document.getElementById("btnSalvar").disabled = false;
            return;
        }

        setTimeout(() => form.submit(), 500);
    });
}

// ===============================
// SINCRONIZAR SELECTS COM HIDDEN
// ===============================
function configurarSelectAnoMes() {
    const anoSel = document.getElementById("anoSelect");
    const mesSel = document.getElementById("mesSelect");
    const anoHid = document.getElementById("anoHidden");
    const mesHid = document.getElementById("mesHidden");

    if (anoSel && anoHid) {
        anoHid.value = anoSel.value;
        anoSel.addEventListener("change", () => {
            anoHid.value = anoSel.value;
        });
    }

    if (mesSel && mesHid) {
        mesHid.value = mesSel.value;
        mesSel.addEventListener("change", () => {
            mesHid.value = mesSel.value;
        });
    }
}

// ===============================
// INICIALIZAR AO CARREGAR A PÁGINA
// ===============================
document.addEventListener("DOMContentLoaded", function () {
    configurarEnvioFormulario();
    configurarSelectAnoMes();

    const inputs = document.querySelectorAll('input[name^="gerencia"]');
    inputs.forEach(input => validarCampo(input));
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