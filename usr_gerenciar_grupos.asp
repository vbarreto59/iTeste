<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp" -->
<%
' Configurações básicas
Response.Buffer = True
Response.Expires = 0

' Conexão com o banco de dados
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

' Processar ações
Dim acao, grupo_id, usuario_id, mensagem

acao = Request.Form("acao")
If acao = "" Then acao = Request.QueryString("acao")

grupo_id = Request.Form("grupo_id")
If grupo_id = "" Then grupo_id = Request.QueryString("id")

usuario_id = Request.Form("usuario_id")
If usuario_id = "" Then usuario_id = Request.QueryString("usuario_id")

' Validar grupo_id
If Not IsNumeric(grupo_id) Or grupo_id = "" Then
    Response.Redirect "grupos.asp?msg=ID do grupo inválido"
    Response.End
End If

' Processar ações
Select Case acao
    Case "adicionar"
        If IsNumeric(usuario_id) Then
            ' Verificar se já existe
            Set rs = conn.Execute("SELECT COUNT(*) FROM Usuario_Grupo WHERE UserId = " & usuario_id & " AND ID_Grupo = " & grupo_id)
            If rs(0) = 0 Then
                conn.Execute "INSERT INTO Usuario_Grupo (UserId, ID_Grupo) VALUES (" & usuario_id & ", " & grupo_id & ")"
                mensagem = "Usuário adicionado com sucesso!"
            Else
                mensagem = "Usuário já está no grupo!"
            End If
            rs.Close
            Set rs = Nothing
        End If
    
    Case "remover"
        If IsNumeric(usuario_id) And IsNumeric(grupo_id) Then
            On Error Resume Next
            conn.Execute "DELETE FROM Usuario_Grupo WHERE UserId = " & usuario_id & " AND ID_Grupo = " & grupo_id
            If Err.Number <> 0 Then
                mensagem = "Erro ao remover usuário: " & Err.Description
                Err.Clear
            Else
                mensagem = "Usuário removido com sucesso!"
            End If
            On Error GoTo 0
        Else
            mensagem = "IDs inválidos para remoção!"
        End If
End Select

' Obter nome do grupo
Set rs = conn.Execute("SELECT Nome_Grupo FROM Grupo WHERE ID_Grupo = " & grupo_id)
If rs.EOF Then
    Response.Redirect "grupos.asp?msg=Grupo não encontrado"
    Response.End
End If
Dim nome_grupo
nome_grupo = rs("Nome_Grupo")
rs.Close
Set rs = Nothing
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Gerenciar Grupo: <%=nome_grupo%></title>
    
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    
    <!-- DataTables CSS -->
    <link href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css" rel="stylesheet">
    
    <style>
        .container {
            margin-top: 20px;
            margin-bottom: 50px;
        }
        .msg-box {
            margin-top: 20px;
            margin-bottom: 20px;
        }
        .action-form {
            margin-bottom: 30px;
            padding: 20px;
            background-color: #f8f9fa;
            border-radius: 5px;
        }
        .table-container {
            margin-top: 20px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="row">
            <div class="col-12">
                <h1 class="mb-4">Gerenciar Grupo: <%=nome_grupo%></h1>

                <a href="usr_listar.asp" class="btn btn-secondary">Voltar para a Lista de Usuários</a>
                
                <% If mensagem <> "" Then %>
                    <div class="alert alert-<% If InStr(mensagem, "sucesso") > 0 Then Response.Write "success" Else Response.Write "danger" End If %> msg-box">
                        <%=mensagem%>
                    </div>
                <% End If %>

                <div class="action-form">
                    <h2 class="h4 mb-3">Adicionar Usuário</h2>
                    <form method="post" class="row g-3">
                        <input type="hidden" name="acao" value="adicionar">
                        <input type="hidden" name="grupo_id" value="<%=grupo_id%>">
                        
                        <div class="col-md-8">
                            <select name="usuario_id" class="form-select" required>
                                <option value="">Selecione um usuário</option>
                                <%
                                ' Listar usuários que não estão no grupo
                                Set rs = conn.Execute("SELECT u.UserId, u.Usuario FROM Usuarios u " & _
                                                      "WHERE u.UserId NOT IN (SELECT ug.UserId FROM Usuario_Grupo ug WHERE ug.ID_Grupo = " & grupo_id & ") " & _
                                                      "ORDER BY u.Usuario")
                                
                                Do While Not rs.EOF
                                    Response.Write "<option value=""" & rs("UserId") & """>" & Server.HTMLEncode(rs("Usuario")) & "</option>"
                                    rs.MoveNext
                                Loop
                                rs.Close
                                Set rs = Nothing
                                %>
                            </select>
                        </div>
                        
                        <div class="col-md-4">
                            <button type="submit" class="btn btn-primary">Adicionar</button>
                        </div>
                    </form>
                </div>

                <div class="table-container">
                    <h2 class="h4 mb-3">Usuários no Grupo</h2>
                    <table id="usuariosTable" class="table table-striped table-bordered" style="width:100%">
                        <thead class="table-dark">
                            <tr>
                                <th>ID</th>
                                <th>Nome</th>
                                <th>Ações</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            ' Listar usuários do grupo
                            Set rs = conn.Execute("SELECT u.UserId, u.Usuario FROM Usuarios u " & _
                                                 "INNER JOIN Usuario_Grupo ug ON u.UserId = ug.UserId " & _
                                                 "WHERE ug.ID_Grupo = " & grupo_id & " ORDER BY u.Usuario")
                            
                            Do While Not rs.EOF
                                Response.Write "<tr>"
                                Response.Write "<td>" & rs("UserId") & "</td>"
                                Response.Write "<td>" & UCase(Server.HTMLEncode(rs("Usuario"))) & "</td>"
                                Response.Write "<td><a href=""?acao=remover&id=" & grupo_id & "&usuario_id=" & rs("UserId") & """ class=""btn btn-danger btn-sm"" onclick=""return confirm('Tem certeza que deseja remover este usuário do grupo?')"">Remover</a></td>"
                                Response.Write "</tr>"
                                rs.MoveNext
                            Loop
                            
                            If rs.EOF And rs.BOF Then
                                Response.Write "<tr><td colspan=""3"" class=""text-center"">Nenhum usuário neste grupo</td></tr>"
                            End If
                            
                            rs.Close
                            Set rs = Nothing
                            %>
                        </tbody>
                    </table>
                </div>
                
                <div class="mt-4">
                    <a href="grupos.asp" class="btn btn-secondary">Voltar para lista de grupos</a>
                </div>
            </div>
        </div>
    </div>

    <!-- jQuery, Bootstrap JS and DataTables JS -->
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
    
    <script>
        $(document).ready(function() {
            $('#usuariosTable').DataTable({
                pageLength: 100,
                order: [[ 1, "asc" ]],
                language: {
                    url: 'https://cdn.datatables.net/plug-ins/1.13.6/i18n/pt-BR.json'
                },
                responsive: true
            });
        });
    </script>
</body>
</html>

<%
' Fechar conexão
conn.Close
Set conn = Nothing
%>