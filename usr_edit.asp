<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp" -->
<!--#include file="funcoes.inc" -->

<%
' ====================================================================================
' BLOC0 VBSCRIPT: Processamento inicial (carrega dados do usuário para edição)
' ====================================================================================

' Verifica se foi passado o ID do usuário a ser editado
Dim UserId, currentGroupId
UserId = Request.QueryString("Id")

' Se não tiver UserId ou não for numérico, redireciona para a lista
If UserId = "" Or Not IsNumeric(UserId) Then
    Response.Redirect("usr_listar.asp")
End If

Dim objConn_Load, rsUsuario_Load, rsUsuarioGrupo_Load
Set objConn_Load = Server.CreateObject("ADODB.Connection")
objConn_Load.Open StrConn

' Busca os dados do usuário no banco de dados
Set rsUsuario_Load = Server.CreateObject("ADODB.Recordset")
rsUsuario_Load.Open "SELECT * FROM Usuarios WHERE UserId = " & CInt(UserId), objConn_Load

' Verifica se o usuário existe
If rsUsuario_Load.EOF Then
    rsUsuario_Load.Close
    Set rsUsuario_Load = Nothing
    objConn_Load.Close
    Set objConn_Load = Nothing
    Response.Redirect("usr_listar.asp")
End If

' Busca o ID do grupo atual do usuário
' Assume que um usuário está em apenas um grupo. Se puder estar em múltiplos, a lógica deve mudar.
Set rsUsuarioGrupo_Load = Server.CreateObject("ADODB.Recordset")
rsUsuarioGrupo_Load.Open "SELECT ID_Grupo FROM Usuario_Grupo WHERE UserId = " & CInt(UserId), objConn_Load

If Not rsUsuarioGrupo_Load.EOF Then
    currentGroupId = rsUsuarioGrupo_Load("ID_Grupo")
Else
    currentGroupId = "" ' Usuário não tem grupo associado
End If
rsUsuarioGrupo_Load.Close
Set rsUsuarioGrupo_Load = Nothing


' Função para converter código de permissão em texto
Function GetPermissaoTexto(codigo)
    Select Case codigo
        Case 1: GetPermissaoTexto = "SUPERADMIN"
        Case 2: GetPermissaoTexto = "ADMIN"
        Case 3: GetPermissaoTexto = "GESTOR-EDITOR"
        Case 4: GetPermissaoTexto = "EDITOR"
        Case 5: GetPermissaoTexto = "CORRETOR"
        Case 6: GetPermissaoTexto = "CORRETOR-EDITOR"
        Case 7: GetPermissaoTexto = "VISUALIZADOR"
        Case Else: GetPermissaoTexto = "DESCONHECIDO"
    End Select
End Function

' ====================================================================================
' BLOC0 VBSCRIPT: Processamento do formulário (executado ao submeter o POST)
' ====================================================================================
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' Recupera o ID do campo hidden
    UserId = Request.Form("UserId")

    ' Recupera os novos dados do formulário
    Dim Usuario, Senha, Email, Telefones, Acoes, Permissao, Funcao, Ativo, New_ID_Grupo
    Usuario = Trim(Request.Form("Usuario"))
    Senha = Trim(Request.Form("SenhaAtual"))
    Email = Trim(Request.Form("Email"))         ' Novo campo
    Telefones = Trim(Request.Form("Telefones")) ' Novo campo
    Acoes = Trim(Request.Form("Acoes"))
    Permissao = Request.Form("Permissao")
    Funcao = Trim(Request.Form("Funcao"))
    Ativo = Request.Form("Ativo")
    New_ID_Grupo = Request.Form("ID_Grupo")     ' Novo campo

    ' 
    Nome = Request.Form("Nome")         ' Novo campo: Nome
    CRECI = Request.Form("CRECI")       ' Novo campo: CRECI    
    
    ' Converte para booleano (Access usa -1 para True e 0 para False)
    If Ativo = "1" Then
        Ativo = -1
    Else
        Ativo = 0
    End If

    ' Validação do ID
    If UserId = "" Or Not IsNumeric(UserId) Then
        Session("MsgErro") = "ID de usuário inválido!"
        Response.Redirect("usr_listar.asp")
    End If

    ' Restante das validações
    If Usuario = "" Or Acoes = "" Or Permissao = "" Or Funcao = "" Or New_ID_Grupo = "" Then
        Session("MsgErro") = "Todos os campos obrigatórios (Usuário, Nível de Permissão, Função, Ações e Grupo) devem ser preenchidos!"
        Response.Redirect("usr_edit.asp?Id=" & UserId) ' Usa "Id" para a QueryString
    End If

    If Not IsNumeric(Permissao) Or (CInt(Permissao) < 1 Or CInt(Permissao) > 7) Then
        Session("MsgErro") = "Nível de permissão inválido!"
        Response.Redirect("usr_edit.asp?Id=" & UserId)
    End If

    If Not IsNumeric(New_ID_Grupo) Then
        Session("MsgErro") = "Grupo inválido!"
        Response.Redirect("usr_edit.asp?Id=" & UserId)
    End If

    ' Verifica se o nome de usuário já existe (exceto para o próprio usuário)
    Dim rsCheck
    Set rsCheck = Server.CreateObject("ADODB.Recordset")
    Set objConn_Check = Server.CreateObject("ADODB.Connection") ' Nova conexão para verificação
    objConn_Check.Open StrConn
    rsCheck.Open "SELECT UserId FROM Usuarios WHERE Usuario = '" & Replace(Usuario, "'", "''") & "' AND UserId <> " & CInt(UserId), objConn_Check

    If Not rsCheck.EOF Then
        Session("MsgErro") = "Este nome de usuário já está em uso por outro usuário!"
        rsCheck.Close
        Set rsCheck = Nothing
        objConn_Check.Close
        Set objConn_Check = Nothing
        Response.Redirect("usr_edit.asp?Id=" & UserId)
    End If
    rsCheck.Close
    Set rsCheck = Nothing
    objConn_Check.Close
    Set objConn_Check = Nothing


    Dim cmd_Update, objConn_Update
    Set objConn_Update = Server.CreateObject("ADODB.Connection")
    objConn_Update.Open StrConn
    Set cmd_Update = Server.CreateObject("ADODB.Command")
    cmd_Update.ActiveConnection = objConn_Update

    'On Error Resume Next ' Inicia tratamento de erro para transação

    ' Inicia a transação
    objConn_Update.BeginTrans

    ' 1. Atualização na tabela Usuarios (incluindo Email e Telefones)
    If Senha <> "" Then
        cmd_Update.CommandText = "UPDATE Usuarios SET Usuario = ?, SenhaAtual = ?, Nome=?, CRECI=?, Email = ?, Telefones = ?, Acoes = ?, Permissao = ?, Funcao = ?, Ativo = ? WHERE UserId = ?"
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Usuario", 202, 1, 50, Usuario)
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("SenhaAtual", 202, 1, 50, Senha)
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Nome", 202, 1, 255, Nome)      ' Novo parâmetro: Nome
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("CRECI", 202, 1, 50, CRECI)    ' Novo parâmetro: CRECI
        ' 

        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Email", 202, 1, 255, Email)
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Telefones", 202, 1, 255, Telefones)
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Acoes", 202, 1, 255, Acoes)
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Permissao", 3, 1, , CInt(Permissao))
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Funcao", 202, 1, 50, Funcao)
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Ativo", 11, 1, , CBool(Ativo))
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("UserId", 3, 1, , CInt(UserId))
    Else
        cmd_Update.CommandText = "UPDATE Usuarios SET Usuario = ?, Nome=?, CRECI=?, Email = ?, Telefones = ?, Acoes = ?, Permissao = ?, Funcao = ?, Ativo = ? WHERE UserId = ?"
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Usuario", 202, 1, 50, Usuario)

        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Nome", 202, 1, 255, Nome)      ' Novo parâmetro: Nome
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("CRECI", 202, 1, 50, CRECI)    ' Novo parâmetro: CRECI

        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Email", 202, 1, 255, Email)
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Telefones", 202, 1, 255, Telefones)
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Acoes", 202, 1, 255, Acoes)
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Permissao", 3, 1, , CInt(Permissao))
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Funcao", 202, 1, 50, Funcao)
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("Ativo", 11, 1, , CBool(Ativo))
        cmd_Update.Parameters.Append cmd_Update.CreateParameter("UserId", 3, 1, , CInt(UserId))
    End If

    cmd_Update.Execute ' Executa a atualização de Usuários

    'montar o sql para passar para o token ----------------------------'
Dim meuSQL
meuSQL = Replace(sqlComParametros, "?", "'" & Usuario & "'", 1, 1)
meuSQL = Replace(meuSQL, "?", "'" & Senha & "'", 1, 1)
meuSQL = Replace(meuSQL, "?", "'" & Nome & "'", 1, 1)
meuSQL = Replace(meuSQL, "?", "'" & CRECI & "'", 1, 1)
meuSQL = Replace(meuSQL, "?", "'" & Email & "'", 1, 1)
meuSQL = Replace(meuSQL, "?", "'" & Telefones & "'", 1, 1)
meuSQL = Replace(meuSQL, "?", "'" & Acoes & "'", 1, 1)
meuSQL = Replace(meuSQL, "?", CInt(Permissao), 1, 1)
meuSQL = Replace(meuSQL, "?", "'" & Funcao & "'", 1, 1)
meuSQL = Replace(meuSQL, "?", CBool(Ativo), 1, 1)
meuSQL = Replace(meuSQL, "?", CInt(UserId), 1, 1)    

GravaToken "updUsuario", meuSQL

    '-------------------------------------------------------------'
    If Err.Number <> 0 Then
        objConn_Update.RollbackTrans
        Session("MsgErro") = "Erro ao atualizar usuário: " & Err.Description
        objConn_Update.Close
        Set objConn_Update = Nothing
        Set cmd_Update = Nothing
        Response.Redirect("usr_edit.asp?Id=" & UserId)
    End If
    
    Set cmd_Update = Nothing ' Libera o objeto Command

    ' 2. Atualizar a relação na tabela Usuario_Grupo
    ' Primeiro, exclui a associação existente para este usuário (se houver)
    Set cmd_Update = Server.CreateObject("ADODB.Command")
    cmd_Update.ActiveConnection = objConn_Update
    cmd_Update.CommandText = "DELETE FROM Usuario_Grupo WHERE UserId = ?"
    cmd_Update.Parameters.Append cmd_Update.CreateParameter("UserIdParam", 3, 1, , CInt(UserId))
    cmd_Update.Execute ' Executa a exclusão

    If Err.Number <> 0 Then
        objConn_Update.RollbackTrans
        Session("MsgErro") = "Erro ao remover grupo antigo: " & Err.Description
        objConn_Update.Close
        Set objConn_Update = Nothing
        Set cmd_Update = Nothing
        Response.Redirect("usr_edit.asp?Id=" & UserId)
    End If
    Set cmd_Update = Nothing ' Libera o objeto Command

    ' Segundo, insere a nova associação
    Set cmd_Update = Server.CreateObject("ADODB.Command")
    cmd_Update.ActiveConnection = objConn_Update
    cmd_Update.CommandText = "INSERT INTO Usuario_Grupo (UserId, ID_Grupo) VALUES (?, ?)"
    cmd_Update.Parameters.Append cmd_Update.CreateParameter("UserIdParam", 3, 1, , CInt(UserId))
    cmd_Update.Parameters.Append cmd_Update.CreateParameter("IDGrupoParam", 3, 1, , CInt(New_ID_Grupo))
    cmd_Update.Execute ' Executa a inserção do novo grupo

    If Err.Number <> 0 Then
        objConn_Update.RollbackTrans
        Session("MsgErro") = "Erro ao atribuir novo grupo: " & Err.Description
        objConn_Update.Close
        Set objConn_Update = Nothing
        Set cmd_Update = Nothing
        Response.Redirect("usr_edit.asp?Id=" & UserId)
    End If
    Set cmd_Update = Nothing ' Libera o objeto Command

    ' Finaliza a transação se tudo deu certo
    objConn_Update.CommitTrans
    On Error GoTo 0 ' Retorna ao tratamento de erro padrão

    ' Fecha a conexão após todas as operações
    objConn_Update.Close
    Set objConn_Update = Nothing

    Session("MsgSucesso") = "Usuário atualizado com sucesso!"
    Response.Redirect("usr_listar.asp")
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <title>Sunny - Editar Usuário</title>
    
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
    
    <style>
        body {
            background-color: #f8f9fa;
        }
        .card-form {
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        .form-header {
            background-color: #343a40;
            color: white;
            border-radius: 10px 10px 0 0 !important;
        }
        .btn-custom {
            min-width: 100px;
        }
        .required-field::after {
            content: " *";
            color: red;
        }
        /* Estilos para o switch de ativo/inativo */
        .switch {
            position: relative;
            display: inline-block;
            width: 60px;
            height: 34px;
        }
        .switch input {
            opacity: 0;
            width: 0;
            height: 0;
        }
        .slider {
            position: absolute;
            cursor: pointer;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: #ccc;
            transition: .4s;
            border-radius: 34px;
        }
        .slider:before {
            position: absolute;
            content: "";
            height: 26px;
            width: 26px;
            left: 4px;
            bottom: 4px;
            background-color: white;
            transition: .4s;
            border-radius: 50%;
        }
        input:checked + .slider {
            background-color: #28a745;
        }
        input:checked + .slider:before {
            transform: translateX(26px);
        }
        .status-label {
            margin-left: 10px;
            font-weight: 500;
        }
        .status-ativo {
            color: #28a745;
        }
        .status-inativo {
            color: #dc3545;
        }

        /* ESTILOS PARA O OLHO DA SENHA */
        .input-group.password-toggle {
            position: relative;
        }
        .input-group.password-toggle .form-control {
            padding-right: 40px; /* Garante espaço para o ícone */
        }
        .input-group.password-toggle .input-group-append {
            position: absolute;
            right: 0;
            top: 0;
            bottom: 0;
            z-index: 5; /* Garante que o ícone esteja acima do input */
            display: flex;
            align-items: center;
        }
        .input-group.password-toggle .input-group-text {
            background-color: transparent;
            border: none;
            padding: 0.375rem 0.75rem; /* Ajuste para centralizar visualmente */
            cursor: pointer;
            color: #888;
        }
        .input-group.password-toggle .input-group-text:hover {
            color: #333;
        }
        /* Estilos para o badge de permissão (opcional, para visualização no campo Função) */
        .badge-permissao {
            padding: .25em .5em;
            font-size: 75%;
            font-weight: 700;
            line-height: 1;
            text-align: center;
            white-space: nowrap;
            vertical-align: baseline;
            border-radius: .25rem;
            transition: color .15s ease-in-out,background-color .15s ease-in-out,border-color .15s ease-in-out,box-shadow .15s ease-in-out;
        }
        /* Cores de exemplo para os badges, ajuste conforme sua paleta */
        .superadmin { background-color: #dc3545; color: white; } /* Vermelho */
        .admin { background-color: #ffc107; color: #343a40; } /* Amarelo */
        .gestor-editor, .editor { background-color: #007bff; color: white; } /* Azul */
        .corretor, .corretor-editor { background-color: #6c757d; color: white; } /* Cinza */
        .visualizador { background-color: #17a2b8; color: white; } /* Ciano */
        .desconhecido { background-color: #6f42c1; color: white; } /* Roxo */
    </style>
</head>
<body>
    <div class="container py-5">
        <div class="row justify-content-center">
            <div class="col-md-8 col-lg-6">
                <div class="card card-form">
                    <div class="card-header form-header">
                        <h4 class="form-title"><i class="fas fa-user-edit mr-2"></i>Editar Usuário</h4>
                    </div>
                    
                    <div class="card-body">
                        <% If Session("MsgErro") <> "" Then %>
                            <div class="alert alert-danger alert-dismissible fade show" role="alert">
                                <i class="fas fa-exclamation-circle mr-2"></i><%= Session("MsgErro") %>
                                <button type="button" class="close" data-dismiss="alert" aria-label="Close">
                                    <span aria-hidden="true">&times;</span>
                                </button>
                            </div>
                            <% Session("MsgErro") = "" %>
                        <% End If %>
                        
                        <form method="post" action="">
                            <input type="hidden" name="UserId" value="<%= UserId %>">
                            
                            <div class="form-group">
                                <label for="Usuario" class="required-field">Usuário</label>
                                <div class="input-group">
                                    <div class="input-group-prepend">
                                        <span class="input-group-text"><i class="fas fa-user"></i></span>
                                    </div>
                                    <input type="text" class="form-control" id="Usuario" name="Usuario" 
                                           value="<%= Server.HTMLEncode(rsUsuario_Load("Usuario")) %>" required>
                                </div>
                            </div>
                            
                            <div class="form-group">
                                <label for="Senha">Nova Senha</label>
                                <div class="input-group password-toggle">
                                    <div class="input-group-prepend">
                                        <span class="input-group-text"><i class="fas fa-lock"></i></span>
                                    </div>
                                    <input type="password" class="form-control" id="Senha" name="Senha" 
                                           placeholder="Confirme a senha" value="<%= Server.HTMLEncode(rsUsuario_Load("SenhaAtual")) %>">
                                    <div class="input-group-append">
                                        <span class="input-group-text" id="togglePassword">
                                            <i class="fas fa-eye"></i>
                                        </span>
                                    </div>
                                </div>
                                <small class="text-muted">Preencha apenas se desejar alterar a senha</small>
                            </div>

<!--  -->
                            <div class="form-group">
                                <label for="Nome">Nome Completo</label>
                                <div class="input-group">
                                    <div class="input-group-prepend">
                                        <span class="input-group-text"><i class="fas fa-signature"></i></span>
                                    </div>
                                    <input type="text" class="form-control" id="Nome" name="Nome" placeholder="Digite o nome completo do usuário" value="<%= (rsUsuario_Load("Nome")) %>">
                                </div>
                            </div>

                            <div class="form-group">
                                <label for="CRECI">CRECI</label>
                                <div class="input-group">
                                    <div class="input-group-prepend">
                                        <span class="input-group-text"><i class="fas fa-id-card-alt"></i></span>
                                    </div>
                                    <input type="text" class="form-control" id="CRECI" name="CRECI" placeholder="Digite o CRECI (opcional)" value="<%= (rsUsuario_Load("CRECI")) %>">
                                </div>
                            </div>
<!--  -->


                            <div class="form-group">
                                <label for="Email">Email</label>
                                <div class="input-group">
                                    <div class="input-group-prepend">
                                        <span class="input-group-text"><i class="fas fa-envelope"></i></span>
                                    </div>
                                    <input type="email" class="form-control" id="Email" name="Email" 
                                           value="<%= rsUsuario_Load("Email") %>">
                                </div>
                            </div>

                            <div class="form-group">
                                <label for="Telefones">Telefones</label>
                                <div class="input-group">
                                    <div class="input-group-prepend">
                                        <span class="input-group-text"><i class="fas fa-phone"></i></span>
                                    </div>
                                    <input type="text" class="form-control" id="Telefones" name="Telefones" 
                                           value="<%= (rsUsuario_Load("Telefones")) %>" placeholder="Ex: (XX) XXXX-XXXX, (XX) XXXX-XXXX">
                                </div>
                            </div>
                            <div class="form-group">
                                <label for="ID_Grupo" class="required-field">Grupo</label>
                                <select class="form-control" id="ID_Grupo" name="ID_Grupo" required>
                                    <option value="" <% If currentGroupId = "" Then Response.Write "selected" %> disabled>Selecione um grupo</option>
                                    <%
                                    ' BLOC0 VBSCRIPT: Popula o dropdown com dados da tabela Grupo
                                    ' Reabre a conexão aqui para buscar os grupos se a página ainda não foi submetida
                                    Dim objConn_Groups, rsGrupos_HTML
                                    Set objConn_Groups = Server.CreateObject("ADODB.Connection")
                                    objConn_Groups.Open StrConn

                                    Set rsGrupos_HTML = objConn_Groups.Execute("SELECT ID_Grupo, Nome_Grupo FROM Grupo ORDER BY Nome_Grupo")

                                    If Not rsGrupos_HTML.EOF Then
                                        Do While Not rsGrupos_HTML.EOF
                                    %>
                                            <option value="<%= rsGrupos_HTML("ID_Grupo") %>" 
                                                <% If CStr(rsGrupos_HTML("ID_Grupo")) = CStr(currentGroupId) Then Response.Write "selected" %>>
                                                <%= rsGrupos_HTML("Nome_Grupo") %>
                                            </option>
                                    <%
                                            rsGrupos_HTML.MoveNext
                                        Loop
                                    End If
                                    rsGrupos_HTML.Close
                                    Set rsGrupos_HTML = Nothing
                                    objConn_Groups.Close
                                    Set objConn_Groups = Nothing
                                    %>
                                </select>
                            </div>
                            <div class="form-group">
                                <label for="Permissao" class="required-field">Nível de Permissão</label>
                                <select class="form-control" id="Permissao" name="Permissao" required onchange="atualizarFuncao()">
                                    <option value="1" <% If rsUsuario_Load("Permissao") = 1 Then Response.Write "selected" %>>SUPERADMIN</option>
                                    <option value="2" <% If rsUsuario_Load("Permissao") = 2 Then Response.Write "selected" %>>ADMIN</option>
                                    <option value="3" <% If rsUsuario_Load("Permissao") = 3 Then Response.Write "selected" %>>GESTOR-EDITOR</option>
                                    <option value="4" <% If rsUsuario_Load("Permissao") = 4 Then Response.Write "selected" %>>EDITOR</option>
                                    <option value="5" <% If rsUsuario_Load("Permissao") = 5 Then Response.Write "selected" %>>CORRETOR</option>
                                    <option value="6" <% If rsUsuario_Load("Permissao") = 6 Then Response.Write "selected" %>>CORRETOR-EDITOR</option>
                                    <option value="7" <% If rsUsuario_Load("Permissao") = 7 Then Response.Write "selected" %>>VISUALIZADOR</option>
                                </select>
                            </div>
                            
                            <div class="form-group">
                                <label for="Funcao" class="required-field">Função</label>
                                <div class="input-group">
                                    <input type="text" class="form-control" id="Funcao" name="Funcao" 
                                           value="<%= GetPermissaoTexto(rsUsuario_Load("Permissao")) %>" readonly>
                                    <div class="input-group-append">
                                        <span class="input-group-text badge-permissao <%= LCase(Replace(GetPermissaoTexto(rsUsuario_Load("Permissao")), "-", "")) %>">
                                            <i class="fas fa-user-tag"></i>
                                        </span>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="form-group">
                                <label for="Acoes" class="required-field">Ações Permitidas</label>
                                <textarea class="form-control" id="Acoes" name="Acoes" rows="3" required readonly><%= Server.HTMLEncode(rsUsuario_Load("Acoes")) %></textarea>
                            </div>
                            
                            <div class="form-group">
                                <label>Status do Usuário</label>
                                <div class="d-flex align-items-center">
                                    <label class="switch">
                                        <input type="checkbox" id="Ativo" name="Ativo" value="1" <% If CBool(rsUsuario_Load("Ativo")) Then Response.Write "checked" %>>
                                        <span class="slider"></span>
                                    </label>
                                    <span class="status-label <% If CBool(rsUsuario_Load("Ativo")) Then Response.Write "status-ativo" Else Response.Write "status-inativo" %>">
                                        <% If CBool(rsUsuario_Load("Ativo")) Then Response.Write "ATIVO" Else Response.Write "INATIVO" %>
                                    </span>
                                </div>
                            </div>
                            
                            <div class="form-group text-right mt-4">
                                <button type="submit" class="btn btn-primary btn-custom">
                                    <i class="fas fa-save mr-2"></i>Salvar Alterações
                                </button>
                                <a href="usr_listar.asp" class="btn btn-secondary btn-custom">
                                    <i class="fas fa-times mr-2"></i>Cancelar
                                </a>
                            </div>
                        </form>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.1/dist/umd/popper.min.js"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
    
<script>
    // Função para atualizar a Função e as Ações baseadas na permissão selecionada
    function atualizarFuncao() {
        var permissaoSelecionada = document.getElementById("Permissao").value;
        var campoFuncao = document.getElementById("Funcao");
        var campoAcoes = document.getElementById("Acoes");
        var badge = document.querySelector('.input-group-text.badge-permissao');
        var funcaoTexto = "";
        var acoesTexto = "";
        var badgeClass = "";

        switch (permissaoSelecionada) {
            case "1": // SUPERADMIN
                funcaoTexto = "SUPERADMIN";
                acoesTexto = "/GESTAO/SISTEMA/USUARIOS/BACKUP/CONSULTAR/INCLUIR/EDITAR/EXCLUIR/";
                badgeClass = "superadmin";
                break;
            case "2": // ADMIN
                funcaoTexto = "ADMIN";
                acoesTexto = "/USUARIOS/CONSULTAR/INCLUIR/EDITAR/EXCLUIR/";
                badgeClass = "admin";
                break;
            case "3": // GESTOR-EDITOR
                funcaoTexto = "GESTOR-EDITOR";
                acoesTexto = "/GESTOR/CONSULTAR/INCLUIR/EDITAR/EXCLUIR/";
                badgeClass = "gestor-editor"; // Use a classe específica
                break;
            case "4": // EDITOR
                funcaoTexto = "EDITOR";
                acoesTexto = "/CONSULTAR/INCLUIR/EDITAR/EXCLUIR/";
                badgeClass = "editor";
                break;
            case "5": // CORRETOR
                funcaoTexto = "CORRETOR";
                acoesTexto = "/CONSULTAR/";
                badgeClass = "corretor";
                break;
            case "6": // CORRETOR-EDITOR
                funcaoTexto = "CORRETOR-EDITOR";
                acoesTexto = "/CONSULTAR/EDITAR/";
                badgeClass = "corretor-editor"; // Use a classe específica
                break;          
            case "7": // VISUALIZADOR
                funcaoTexto = "VISUALIZADOR";
                acoesTexto = "/CONSULTAR/";
                badgeClass = "visualizador";
                break;
            default:
                funcaoTexto = "DESCONHECIDO";
                acoesTexto = "";
                badgeClass = "";
                break;
        }
        
        campoFuncao.value = funcaoTexto;
        campoAcoes.value = acoesTexto;
        
        // Remove todas as classes de permissão existentes e adiciona a nova
        badge.className = 'input-group-text badge-permissao'; // Reseta para a base
        badge.classList.add(badgeClass); // Adiciona a classe específica
        badge.innerHTML = '<i class="fas fa-user-tag"></i>'; // Garante o ícone
    }
    
    // Atualiza o status do switch quando ele é alterado
    $(document).ready(function() {
        $('#Ativo').change(function() {
            if($(this).is(':checked')) {
                $('.status-label').removeClass('status-inativo').addClass('status-ativo').text('ATIVO');
            } else {
                $('.status-label').removeClass('status-ativo').addClass('status-inativo').text('INATIVO');
            }
        });
        
        // Lógica para o "olho" de mostrar/esconder senha
        const passwordInput = document.getElementById('Senha');
        const togglePassword = document.getElementById('togglePassword');

        togglePassword.addEventListener('click', function () {
            // Alterna o tipo do input entre 'password' e 'text'
            const type = passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
            passwordInput.setAttribute('type', type);

            // Alterna o ícone do olho (Font Awesome)
            this.querySelector('i').classList.toggle('fa-eye');
            this.querySelector('i').classList.toggle('fa-eye-slash'); // fa-eye-slash é o olho fechado
        });

        // Inicializa a função atualizarFuncao() ao carregar a página para garantir que o badge e ações estejam corretos
        atualizarFuncao(); 
    });
</script>
</body>
</html>

<%
' Libera o recordset e a conexão do bloco de carregamento inicial
If IsObject(rsUsuario_Load) Then
    rsUsuario_Load.Close()
    Set rsUsuario_Load = Nothing
End If

%>