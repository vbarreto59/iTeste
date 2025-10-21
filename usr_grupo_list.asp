<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp" -->

<%
' =========================================================
' FUNÇÕES DE GERENCIAMENTO DE CONEXÃO (DO CONEXAO.ASP)
' Agora reutilizamos a string de conexão StrConn do conexao.asp
' =========================================================

Dim objConn ' Este será nosso objeto de conexão local

Function GetOpenConnection()
    Dim tempConn
    Set tempConn = Server.CreateObject("ADODB.Connection")
    On Error Resume Next ' Habilita tratamento de erro para a abertura da conexão
    ' Usa StrConn que vem do conexao.asp como a string de conexão
    tempConn.Open StrConn
    If Err.Number <> 0 Then
        Response.Write "<div class='alert alert-danger' role='alert'>"
        Response.Write "<strong>Erro de Conexão:</strong> Não foi possível abrir a conexão com o banco de dados.<br>"
        Response.Write "Detalhes: " & Err.Description
        Response.Write "</div>"
        Response.End ' Termina o script se não puder conectar
    End If
    On Error GoTo 0
    Set GetOpenConnection = tempConn
End Function

Sub CloseConnection(ByRef connectionObject)
    If Not connectionObject Is Nothing Then
        If connectionObject.State = 1 Then
            connectionObject.Close
        End If
        Set connectionObject = Nothing
    End If
End Sub

' =========================================================
' FIM DAS FUNÇÕES DE CONEXÃO
' =========================================================
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gerenciamento de Grupos</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.datatables.net/2.0.7/css/dataTables.bootstrap5.min.css" rel="stylesheet">
    <style>
        body { padding-top: 20px; }
        .container { max-width: 960px; }
    </style>
</head>
<body>
    <div class="container">
        <h1 class="mb-4">Gerenciamento de Grupos</h1>
        <a href="usr_listar.asp" class="btn btn-secondary">Voltar para a Lista de Usuários</a>

        <%
        ' --- EXIBIÇÃO DE MENSAGENS DE SUCESSO/ERRO ---
        Dim Mensagem
        Mensagem = Request.QueryString("msg")
        If Mensagem <> "" Then
            Dim alertClass
            Select Case Left(Mensagem, 3)
                Case "suc": alertClass = "alert-success" ' sucsso
                Case "err": alertClass = "alert-danger"  ' erro
                Case Else:  alertClass = "alert-info"
            End Select
            Response.Write "<div class='alert " & alertClass & " alert-dismissible fade show' role='alert'>"
            Response.Write Right(Mensagem, Len(Mensagem) - 3)
            Response.Write "<button type='button' class='btn-close' data-bs-dismiss='alert' aria-label='Close'></button>"
            Response.Write "</div>"
        End If

        ' --- LÓGICA DE EDIÇÃO / ADIÇÃO (FORMULÁRIOS) ---
        Dim Acao
        Acao = Request.QueryString("acao")

        If Acao = "editar" Then
            Dim GrupoID
            GrupoID = Request.QueryString("id")

            ' Validação robusta para o ID do Grupo
            If Not IsNumeric(GrupoID) OR Trim(CStr(GrupoID)) = "" OR CLng(GrupoID) <= 0 Then
                Response.Redirect "usr_grupo_list.asp?msg=errID Inválido ou não fornecido."
                Response.End
            End If

            Set objConn = GetOpenConnection() ' Abre a conexão

            Dim rsGrupoEditar
            Set rsGrupoEditar = objConn.Execute("SELECT ID_Grupo, Nome_Grupo, Descricao_Grupo FROM Grupo WHERE ID_Grupo = " & GrupoID)

            If Not rsGrupoEditar.EOF Then
        %>
                <h2 class="mb-3">Editar Grupo: <%=Server.HTMLEncode(rsGrupoEditar("Nome_Grupo"))%></h2>
                <form action="usr_grupo_list.asp" method="post" class="mb-4">
                    <input type="hidden" name="acao" value="salvar_edicao">
                    <input type="hidden" name="id" value="<%=rsGrupoEditar("ID_Grupo")%>">
                    <div class="mb-3">
                        <label for="nomeGrupo" class="form-label">Nome do Grupo:</label>
                        <input type="text" class="form-control" id="nomeGrupo" name="nomeGrupo" value="<%=Server.HTMLEncode(rsGrupoEditar("Nome_Grupo"))%>" required>
                    </div>
                    <div class="mb-3">
                        <label for="descricaoGrupo" class="form-label">Descrição:</label>
                        <textarea class="form-control" id="descricaoGrupo" name="descricaoGrupo" rows="3"><%=Server.HTMLEncode(rsGrupoEditar("Descricao_Grupo"))%></textarea>
                    </div>
                    <button type="submit" class="btn btn-primary">Salvar Alterações</button>
                    <a href="usr_grupo_list.asp" class="btn btn-secondary">Cancelar</a>
                </form>
        <%
            Else
                Response.Redirect "usr_grupo_list.asp?msg=errGrupo não encontrado para edição."
            End If
            rsGrupoEditar.Close
            Set rsGrupoEditar = Nothing
            Call CloseConnection(objConn) ' Fecha a conexão
        Else ' Se Acao não for "editar", mostra o formulário de adicionar
        %>
            <h2 class="mb-3">Adicionar Novo Grupo</h2>
            <form action="usr_grupo_list.asp" method="post" class="mb-4">
                <input type="hidden" name="acao" value="adicionar">
                <div class="mb-3">
                    <label for="nomeGrupo" class="form-label">Nome do Grupo:</label>
                    <input type="text" class="form-control" id="nomeGrupo" name="nomeGrupo" required>
                </div>
                <div class="mb-3">
                    <label for="descricaoGrupo" class="form-label">Descrição:</label>
                    <textarea class="form-control" id="descricaoGrupo" name="descricaoGrupo" rows="3"></textarea>
                </div>
                <button type="submit" class="btn btn-success">Adicionar Grupo</button>
            </form>
        <% End If %>

        <h2 class="mb-3">Grupos Cadastrados</h2>
        <table id="tabelaGrupos" class="table table-striped table-bordered" style="width:100%">
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Nome do Grupo</th>
                    <th>Descrição</th>
                    <th>Ações</th>
                </tr>
            </thead>
            <tbody>
                <%
                ' --- EXIBIÇÃO DA TABELA DE GRUPOS ---
                Set objConn = GetOpenConnection() ' Abre a conexão

                Dim rsGrupos
                Set rsGrupos = objConn.Execute("SELECT ID_Grupo, Nome_Grupo, Descricao_Grupo FROM Grupo ORDER BY Nome_Grupo")

                If Not rsGrupos.EOF Then
                    Do While Not rsGrupos.EOF
                %>
                        <tr>
                            <td><%=rsGrupos("ID_Grupo")%></td>
                            <td><%=Server.HTMLEncode(rsGrupos("Nome_Grupo"))%></td>
                            <td><%=Server.HTMLEncode(rsGrupos("Descricao_Grupo"))%></td>
                            <td>
                                <a href="usr_grupo_list.asp?acao=editar&id=<%=rsGrupos("ID_Grupo")%>" class="btn btn-sm btn-primary">Editar</a>
                                <a href="usr_gerenciar_grupos.asp?id=<%=rsGrupos("ID_Grupo")%>" class="btn btn-sm btn-info">Gerenciar Usuários</a>
                                <button type="button" class="btn btn-sm btn-danger" onclick="confirmarExclusao(<%=rsGrupos("ID_Grupo")%>, '<%=Server.HTMLEncode(rsGrupos("Nome_Grupo"))%>')">Excluir</button>
                            </td>
                        </tr>
                <%
                        rsGrupos.MoveNext
                    Loop
                Else
                    Response.Write "<tr><td colspan='4'>Nenhum grupo cadastrado.</td></tr>"
                End If

                rsGrupos.Close
                Set rsGrupos = Nothing
                Call CloseConnection(objConn) ' Fecha a conexão
                %>
            </tbody>
        </table>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
    <script src="https://cdn.datatables.net/2.0.7/js/dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/2.0.7/js/dataTables.bootstrap5.min.js"></script>

    <script>
        $(document).ready(function() {
            $('#tabelaGrupos').DataTable({
                "language": {
                    "url": "https://cdn.datatables.net/plug-ins/2.0.7/i18n/pt-BR.json"
                }
            });
        });

        function confirmarExclusao(id, nome) {
            if (confirm('Tem certeza que deseja excluir o grupo "' + nome + '"? Isso removerá todas as associações de usuários com este grupo.')) {
                window.location.href = 'usr_grupo_list.asp?acao=excluir&id=' + id;
            }
        }
    </script>
</body>
</html>

<%
' =========================================================
' PROCESSAMENTO DE FORMULÁRIOS E AÇÕES (POST e GET)
' =========================================================

' --- Processamento POST ---
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim AcaoForm, NomeGrupo, DescricaoGrupo, GrupoIDForm
    AcaoForm = Request.Form("acao")
    NomeGrupo = Request.Form("nomeGrupo")
    DescricaoGrupo = Request.Form("descricaoGrupo")
    GrupoIDForm = Request.Form("id") ' Apenas para "salvar_edicao"

    Dim SQL, MsgRedirect

    Set objConn = GetOpenConnection() ' Abre a conexão para a operação POST

    Select Case AcaoForm
        Case "adicionar"
            If NomeGrupo <> "" Then
                SQL = "INSERT INTO Grupo (Nome_Grupo, Descricao_Grupo) VALUES ('" & Replace(NomeGrupo, "'", "''") & "', '" & Replace(DescricaoGrupo, "'", "''") & "')"
                On Error Resume Next
                objConn.Execute SQL
                If Err.Number = 0 Then
                    MsgRedirect = "sucssoGrupo adicionado com sucesso!"
                Else
                    ' Erro de violação de UNIQUE (se houver índice único no Nome_Grupo)
                    If InStr(Err.Description, "duplicate") > 0 Or InStr(Err.Description, "índice único") > 0 Then
                        MsgRedirect = "errJá existe um grupo com este nome."
                    Else
                        MsgRedirect = "errErro ao adicionar grupo: " & Err.Description
                    End If
                End If
                On Error GoTo 0
            Else
                MsgRedirect = "errO nome do grupo é obrigatório."
            End If

        Case "salvar_edicao"
            If IsNumeric(GrupoIDForm) And Trim(CStr(GrupoIDForm)) <> "" And CLng(GrupoIDForm) > 0 And NomeGrupo <> "" Then
                SQL = "UPDATE Grupo SET Nome_Grupo = '" & Replace(NomeGrupo, "'", "''") & "', Descricao_Grupo = '" & Replace(DescricaoGrupo, "'", "''") & "' WHERE ID_Grupo = " & GrupoIDForm
                On Error Resume Next
                objConn.Execute SQL
                If Err.Number = 0 Then
                    MsgRedirect = "sucssoGrupo atualizado com sucesso!"
                Else
                    ' Erro de violação de UNIQUE (se houver índice único no Nome_Grupo)
                    If InStr(Err.Description, "duplicate") > 0 Or InStr(Err.Description, "índice único") > 0 Then
                        MsgRedirect = "errJá existe outro grupo com este nome."
                    Else
                        MsgRedirect = "errErro ao atualizar grupo: " & Err.Description
                    End If
                End If
                On Error GoTo 0
            Else
                MsgRedirect = "errDados inválidos para edição."
            End If

        Case Else
            MsgRedirect = "errAção POST não reconhecida."
    End Select

    Call CloseConnection(objConn) ' Fecha a conexão após a operação POST
    Response.Redirect "usr_grupo_list.asp?msg=" & Server.URLEncode(MsgRedirect)
    Response.End ' Importante para parar o script após o redirect

' --- Processamento GET (para exclusão) ---
ElseIf Request.ServerVariables("REQUEST_METHOD") = "GET" And Acao = "excluir" Then
    Dim GrupoIDExcluir
    GrupoIDExcluir = Request.QueryString("id")

    If IsNumeric(GrupoIDExcluir) And Trim(CStr(GrupoIDExcluir)) <> "" And CLng(GrupoIDExcluir) > 0 Then
        Set objConn = GetOpenConnection() ' Abre a conexão para a operação GET (exclusão)

        On Error Resume Next ' Inicia tratamento de erro para a exclusão
        ' Primeiro, remover as associações do grupo na tabela Usuario_Grupo
        objConn.Execute "DELETE FROM Usuario_Grupo WHERE ID_Grupo = " & GrupoIDExcluir
        If Err.Number <> 0 Then
            MsgRedirect = "errErro ao remover associações de usuários: " & Err.Description
        Else
            ' Se as associações foram removidas, exclui o grupo
            objConn.Execute "DELETE FROM Grupo WHERE ID_Grupo = " & GrupoIDExcluir
            If Err.Number = 0 Then
                MsgRedirect = "sucssoGrupo e suas associações removidos com sucesso!"
            Else
                MsgRedirect = "errErro ao excluir grupo: " & Err.Description
            End If
        End If
        On Error GoTo 0 ' Desativa tratamento de erro

        Call CloseConnection(objConn) ' Fecha a conexão após a operação GET
    Else
        MsgRedirect = "errID de grupo inválido para exclusão."
    End If

    Response.Redirect "usr_grupo_list.asp?msg=" & Server.URLEncode(MsgRedirect)
    Response.End ' Importante para parar o script
End If

' A conexão objConn não precisa ser fechada aqui, pois as funções GetOpenConnection
' e CloseConnection garantem que ela seja gerenciada por operação.
%>