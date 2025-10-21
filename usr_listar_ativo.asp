<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp" -->

<%
' Obter todos os usuários e os grupos que participam
Dim rsUsers, rsGrupos, userId, grupos
Set rsUsers = Server.CreateObject("ADODB.Recordset")
' CERTIFIQUE-SE DE QUE SUA TABELA 'Usuarios' TEM UMA COLUNA CHAMADA 'Ativo' DO TIPO BIT/BOOLEAN
rsUsers.Open "SELECT * FROM Usuarios ORDER BY Usuario ASC", StrConn
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>Sunny - Lista de Usuários</title>
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">

    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.1/css/all.min.css">

    <link rel="stylesheet" href="https://cdn.datatables.net/1.10.22/css/dataTables.bootstrap4.min.css">

    <style>
        body {
            background-color: #f8f9fa;
        }
        .table-responsive {
            background-color: white;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            padding: 20px;
            margin-top: 20px;
        }
        .table-header {
            background-color: #343a40;
            color: white;
            border-radius: 10px 10px 0 0;
            padding: 15px 20px;
            margin-bottom: 0;
        }
        .btn-sm {
            min-width: 70px;
        }
        .table {
            width: 100%;
        }
        .table th {
            white-space: nowrap;
        }
        .badge-permissao {
            font-size: 0.85em;
            padding: 0.35em 0.65em;
        }
        .badge-grupo {
            font-size: 0.8em;
            margin-right: 3px;
            margin-bottom: 3px;
            display: inline-block;
        }
        .grupos-container {
            max-width: 250px;
        }
        .header-actions {
            margin-bottom: 20px;
        }
        /* Custom switch for Active/Inactive */
        .switch {
            position: relative;
            display: inline-block;
            width: 40px; /* Smaller width for better fit */
            height: 23px; /* Smaller height */
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
            -webkit-transition: .4s;
            transition: .4s;
            border-radius: 23px; /* Match height for rounded slider */
        }

        .slider:before {
            position: absolute;
            content: "";
            height: 17px; /* Smaller circle */
            width: 17px; /* Smaller circle */
            left: 3px; /* Adjust for padding */
            bottom: 3px; /* Adjust for padding */
            background-color: white;
            -webkit-transition: .4s;
            transition: .4s;
            border-radius: 50%;
        }

        input:checked + .slider {
            background-color: #28a745; /* Green for active */
        }

        input:focus + .slider {
            box-shadow: 0 0 1px #2196F3;
        }

        input:checked + .slider:before {
            -webkit-transform: translateX(17px); /* Move based on new width */
            -ms-transform: translateX(17px);
            transform: translateX(17px);
        }

        /* Styling for inactive rows */
        .user-inativo {
            opacity: 0.6; /* Dim the row */
            /* Add any other styles you want for inactive users, e.g., text-decoration: line-through; */
        }
    </style>
</head>
<body>

    <div class="container">
        <div class="d-flex justify-content-between align-items-center header-actions">
            <h4 class="mb-0"><i class="fas fa-users mr-2"></i>Lista de Usuários</h4>
            <div>
                <a href="#" class="btn btn-info" onclick="window.close(); return false;">
                    <i class="fas fa-times mr-1"></i> Fechar
                </a>
                <a href="usr_novo.asp" class="btn btn-primary">
                    <i class="fas fa-user-plus mr-1"></i> Novo Usuário
                </a>
                <a href="usr_grupo_list.asp" class="btn btn-primary">
                    <i class="fas fa-users mr-1"></i> Grupos
                </a>
                <a href="usr_grupo_contagem.asp" class="btn btn-primary">
                    <i class="fas fa-users mr-1"></i> Grupos Contagem
                </a>
            </div>
        </div>

        <div class="table-responsive">
            <table id="tabelaUsuarios" class="table table-striped table-bordered table-hover" style="width:100%">
                <thead class="thead-dark">
                    <tr>
                        <th>ID</th>
                        <th>Usuário</th>
                        <th class="text-center">Ativo</th>
                        <th>Função</th>
                        <th>Permissão</th>
                        <th>Grupos</th>
                        <th class="text-center">Ações</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    While Not rsUsers.EOF
                        userId = rsUsers("UserID")

                        ' Obter grupos do usuário
                        Set rsGrupos = Server.CreateObject("ADODB.Recordset")
                        rsGrupos.Open "SELECT g.ID_Grupo, g.Nome_Grupo FROM Grupo g " & _
                                      "INNER JOIN Usuario_Grupo ug ON g.ID_Grupo = ug.ID_Grupo " & _
                                      "WHERE ug.UserId = " & userId & " ORDER BY g.Nome_Grupo", StrConn

                        grupos = ""
                        Do While Not rsGrupos.EOF
                            grupos = grupos & "<span class='badge badge-info badge-grupo'>" & Server.HTMLEncode(rsGrupos("Nome_Grupo")) & "</span>"
                            rsGrupos.MoveNext
                        Loop

                        If grupos = "" Then
                            grupos = "<span class='text-muted'>Nenhum grupo</span>"
                        End If

                        rsGrupos.Close
                        Set rsGrupos = Nothing

                        ' Determinar status do usuário para a classe da linha e o checkbox
                        Dim isUserActive
                        isUserActive = CBool(rsUsers("Ativo"))
                        Dim rowClass
                        If Not isUserActive Then
                            rowClass = "user-inativo"
                        Else
                            rowClass = ""
                        End If
                    %>
                    <tr class="<%=rowClass%>">
                        <td><strong><%=userId%></strong></td>
                        <td>
                            <strong><%=UCase(rsUsers("Usuario"))%></strong><br>
                            <small><%=rsUsers("Email")%></small><br>
                            <small><%=rsUsers("Telefones")%></small>
                            <small>CRECI: <%=rsUsers("CRECI")%></small>
                        </td>
                        <td class="text-center"> 
                            <label class="switch">
                                <%
                                Dim checkedAttribute
                                If isUserActive Then
                                    checkedAttribute = "checked"
                                Else
                                    checkedAttribute = ""
                                End If
                                %>
                                <input type="checkbox" <%= checkedAttribute %> onchange="toggleUserStatus(<%=userId%>, this.checked)">
                                <span class="slider"></span>
                            </label>
                        </td>
                        <td>
                            <%
                            Select Case rsUsers("Permissao")
                                Case 1: badgeClass = "badge-danger"
                                Case 2: badgeClass = "badge-warning"
                                Case 3: badgeClass = "badge-warning"
                                Case 4: badgeClass = "badge-info"
                                Case 5: badgeClass = "badge-secondary"
                                Case 6: badgeClass = "badge-secondary"
                                Case Else: badgeClass = "badge-light"
                            End Select
                            %>
                            <span class="badge <%=badgeClass%> badge-permissao"><%=UCase(rsUsers("Funcao"))%></span>
                        </td>
                        <td><%=rsUsers("Permissao")%></td>
                        <td class="grupos-container"><%=grupos%></td>
                        <td class="text-center">
                            <div class="btn-group btn-group-sm" role="group">
                                <a href="usr_edit.asp?id=<%=userId%>" class="btn btn-warning" title="Editar">
                                    <i class="fas fa-edit"></i>
                                </a>
                                <a href="usr_excluir.asp?id=<%=userId%>" class="btn btn-danger" title="Excluir" onclick="return confirm('Tem certeza que deseja excluir este usuário?')">
                                    <i class="fas fa-trash-alt"></i>
                                </a>
                            </div>
                        </td>
                    </tr>
                    <%
                        rsUsers.MoveNext()
                    Wend
                    %>
                </tbody>
            </table>
        </div>

        <footer class="text-center text-muted small mb-3">
            Sunny System &copy; <%= Year(Now()) %>
        </footer>
    </div>

    <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.1/dist/umd/popper.min.js"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>

    <script src="https://cdn.datatables.net/1.10.22/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.10.22/js/dataTables.bootstrap4.min.js"></script>

    <script>
        $(document).ready(function() {
            $('#tabelaUsuarios').DataTable({
                "order": [[1, "asc"]],
                "pageLength": 100,
                "language": {
                    "sEmptyTable": "Nenhum registro encontrado",
                    "sInfo": "Mostrando de _START_ até _END_ de _TOTAL_ registros",
                    "sInfoEmpty": "Mostrando 0 até 0 de 0 registros",
                    "sInfoFiltered": "(Filtrados de _MAX_ registros)",
                    "sInfoPostFix": "",
                    "sInfoThousands": ".",
                    "sLengthMenu": "_MENU_ resultados por página",
                    "sLoadingRecords": "Carregando...",
                    "sProcessing": "Processando...",
                    "sZeroRecords": "Nenhum registro encontrado",
                    "sSearch": "Pesquisar:",
                    "oPaginate": {
                        "sNext": "Próximo",
                        "sPrevious": "Anterior",
                        "sFirst": "Primeiro",
                        "sLast": "Último"
                    },
                    "oAria": {
                        "sSortAscending": ": Ordenar colunas de forma ascendente",
                        "sSortDescending": ": Ordenar colunas de forma descendente"
                    },
                    "select": {
                        "rows": {
                            "_": "Selecionado %d linhas",
                            "0": "Nenhuma linha selecionada",
                            "1": "Selecionado 1 linha"
                        }
                    },
                    "decimal": ",",
                    "thousands": "."
                },
                "dom": '<"top"f>rt<"bottom"lip><"clear">',
                "responsive": true,
                "initComplete": function() {
                    $('.dataTables_filter input').addClass('form-control').attr('placeholder', 'Pesquisar...');
                    $('.dataTables_length select').addClass('form-control');
                },
                "columnDefs": [
                    { "responsivePriority": 1, "targets": 1 }, // Usuário
                    { "responsivePriority": 2, "targets": -1 }, // Ações
                    { "responsivePriority": 3, "targets": 5 }, // Grupos
                    { "responsivePriority": 4, "targets": 3 }, // Função
                    { "responsivePriority": 5, "targets": 2 }, // Ativo (o slider)
                    { "responsivePriority": 6, "targets": 0 }  // ID
                ]
            });
        });
</script>
<!--  -->
<script>
    // JavaScript function to handle the toggle switch
    function toggleUserStatus(userId) {
        // --- Linha adicionada para depuração ---
        console.log("toggleUserStatus chamado para userId:", userId);
        // --- Fim da linha adicionada ---

        var checkbox = $("input[onchange*='" + userId + "']"); // Captura o checkbox atual
        var currentRow = checkbox.closest("tr"); // Captura a linha da tabela

        $.ajax({
            type: "POST",
            url: "usr_update_ativo.asp", // O arquivo que vai processar
            data: { userId: userId }, // Enviando apenas o userId
            success: function(response) {
                // Remove espaços em branco e verifica a resposta do servidor
                var trimmedResponse = response.trim();

                // --- Linha adicionada para depuração ---
                console.log("Resposta do servidor para userId " + userId + ":", trimmedResponse);
                // --- Fim da linha adicionada ---

                if (trimmedResponse === "success_activated") {
                    // Se o servidor disse que ativou
                    checkbox.prop('checked', true);
                    currentRow.removeClass("user-inativo");
                    console.log("Usuário " + userId + " ativado com sucesso!");
                } else if (trimmedResponse === "success_deactivated") {
                    // Se o servidor disse que desativou
                    checkbox.prop('checked', false);
                    currentRow.addClass("user-inativo");
                    console.log("Usuário " + userId + " desativado com sucesso!");
                } else {
                    // Se o servidor retornou um erro ou uma resposta inesperada
                    alert("Erro ao atualizar o status do usuário: " + trimmedResponse);
                    // Reverte o estado do slider localmente em caso de erro
                    checkbox.prop('checked', !checkbox.prop('checked'));
                }
            },
            error: function(xhr, status, error) {
                // --- Linha adicionada para depuração ---
                console.error("Erro AJAX para userId " + userId + ":", status, error, xhr);
                // --- Fim da linha adicionada ---

                alert("Erro na requisição AJAX: " + error);
                // Reverte o estado do slider em caso de erro de rede/servidor
                checkbox.prop('checked', !checkbox.prop('checked'));
            }
        });
    }
</script>
</body>
</html>

<%
rsUsers.Close()
Set rsUsers = Nothing
' Fechar a conexão com o banco de dados se não for gerenciada globalmente em conexao.asp
' If IsObject(Conn) Then
'     If Conn.State = 1 Then Conn.Close
'     Set Conn = Nothing
' End If
%>