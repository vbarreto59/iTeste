<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp" -->

<%
' Obter todos os usuários e os grupos que participam
Dim rsUsers, rsGrupos, userId, grupos
Set rsUsers = Server.CreateObject("ADODB.Recordset")
rsUsers.Open "SELECT * FROM Usuarios ORDER BY Usuario ASC", StrConn
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <title>Sunny - Lista de Usuários</title>
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
  
  <!-- Bootstrap CSS -->
  <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
  
  <!-- Font Awesome -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.1/css/all.min.css">
  
  <!-- DataTables CSS -->
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
    .badge-status {
      font-size: 0.85em;
      padding: 0.5em 0.75em;
      border-radius: 50px;
      min-width: 70px;
      display: inline-block;
      text-align: center;
    }
    .badge-ativo {
      background-color: #28a745;
      color: white;
    }
    .badge-inativo {
      background-color: #dc3545;
      color: white;
    }
    .user-inativo {
      opacity: 0.7;
    }
    .user-photo {
      width: 40px;
      height: 40px;
      border-radius: 50%;
      object-fit: cover;
      border: 2px solid #dee2e6;
    }
    .photo-column {
      width: 60px;
    }
  </style>
</head>
<body>

  <!-- Modal para upload de foto -->
  <div class="modal fade" id="uploadPhotoModal" tabindex="-1" role="dialog" aria-labelledby="uploadPhotoModalLabel" aria-hidden="true">
    <div class="modal-dialog" role="document">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="uploadPhotoModalLabel">Upload de Foto</h5>
          <button type="button" class="close" data-dismiss="modal" aria-label="Fechar">
            <span aria-hidden="true">&times;</span>
          </button>
        </div>
        <form id="uploadPhotoForm" action="upload_foto.asp" method="post" enctype="multipart/form-data">
          <div class="modal-body">
            <input type="hidden" id="userIdForUpload" name="userId">
            <div class="form-group">
              <label for="userPhoto">Selecione uma foto:</label>
              <input type="file" class="form-control-file" id="userPhoto" name="userPhoto" accept="image/*" required>
              <small class="form-text text-muted">A foto será redimensionada para 400x400 pixels.</small>
            </div>
            <div class="text-center">
              <img id="photoPreview" src="" alt="Pré-visualização" class="user-photo d-none" style="width: 100px; height: 100px;">
            </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-dismiss="modal">Cancelar</button>
            <button type="submit" class="btn btn-primary">Enviar Foto</button>
          </div>
        </form>
      </div>
    </div>
  </div>

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
            <th class="photo-column">Foto</th>
            <th>ID</th>
            <th>Usuário</th>
            <th>Status</th>
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
            
            ' Determinar status do usuário
            Dim statusClass, statusText
            If CBool(rsUsers("Ativo")) Then
              statusClass = "badge-ativo"
              statusText = "ATIVO"
            Else
              statusClass = "badge-inativo"
              statusText = "INATIVO"
            End If
            
            ' Verificar se a foto existe
            Dim fotoPath
            fotoPath = "fotos/" & userId & ".jpg"
            
            ' Verificar se o arquivo existe
            Dim fs, fotoExists
            Set fs = Server.CreateObject("Scripting.FileSystemObject")
            fotoExists = fs.FileExists(Server.MapPath(fotoPath))
            Set fs = Nothing
            
            If Not fotoExists Then
              fotoPath = "fotos/semfoto.jpg"
            End If
          %>
          <tr class="<% If Not CBool(rsUsers("Ativo")) Then Response.Write "user-inativo" %>">
            <td class="text-center">
              <img src="<%=fotoPath%>" alt="Foto do usuário" class="user-photo">
            </td>
            <td><strong><%=userId%></strong></td>
            <td><strong><%=UCase(rsUsers("Usuario"))%></strong></td>
            <td>
              <span class="badge badge-status <%=statusClass%>">
                <%=statusText%>
              </span>
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
                <button type="button" class="btn btn-info" title="Alterar Foto" onclick="openUploadModal(<%=userId%>)">
                  <i class="fas fa-camera"></i>
                </button>
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

  <!-- jQuery first, then Popper.js, then Bootstrap JS -->
  <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.1/dist/umd/popper.min.js"></script>
  <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
  
  <!-- DataTables JS -->
  <script src="https://cdn.datatables.net/1.10.22/js/jquery.dataTables.min.js"></script>
  <script src="https://cdn.datatables.net/1.10.22/js/dataTables.bootstrap4.min.js"></script>
  
  <script>
    $(document).ready(function() {
        $('#tabelaUsuarios').DataTable({
            "order": [[2, "asc"]], // Ordenar pelo nome do usuário (coluna 2)
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
                { "responsivePriority": 1, "targets": 2 }, // Usuário
                { "responsivePriority": 2, "targets": -1 }, // Ações
                { "responsivePriority": 3, "targets": 6 }, // Grupos
                { "responsivePriority": 4, "targets": 4 }, // Função
                { "responsivePriority": 5, "targets": 3 }, // Status
                { "responsivePriority": 6, "targets": 1 }, // ID
                { "orderable": false, "targets": [0, -1] } // Foto e Ações não ordenáveis
            ]
        });
        
        // Pré-visualização da foto antes do upload
        $('#userPhoto').change(function() {
            var input = this;
            if (input.files && input.files[0]) {
                var reader = new FileReader();
                reader.onload = function(e) {
                    $('#photoPreview').attr('src', e.target.result).removeClass('d-none');
                }
                reader.readAsDataURL(input.files[0]);
            }
        });
    });
    
    function openUploadModal(userId) {
        $('#userIdForUpload').val(userId);
        $('#photoPreview').addClass('d-none').attr('src', '');
        $('#userPhoto').val('');
        $('#uploadPhotoModal').modal('show');
    }
  </script>
</body>
</html>

<%
rsUsers.Close()
Set rsUsers = Nothing
%>