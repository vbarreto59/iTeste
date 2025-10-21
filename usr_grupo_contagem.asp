<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp" -->

<%
' Obter todos os grupos e seus usuários
Dim rsGrupos, rsUsuarios, grupoId, grupoNome, qtdUsuarios
Set rsGrupos = Server.CreateObject("ADODB.Recordset")
rsGrupos.Open "SELECT g.ID_Grupo, g.Nome_Grupo, COUNT(ug.UserId) AS QtdUsuarios " & _
              "FROM Grupo g LEFT JOIN Usuario_Grupo ug ON g.ID_Grupo = ug.ID_Grupo " & _
              "GROUP BY g.ID_Grupo, g.Nome_Grupo ORDER BY g.Nome_Grupo", StrConn
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <title>Sunny - Lista de Grupos</title>
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
  
  <!-- Bootstrap CSS -->
  <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
  
  <!-- Font Awesome -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.1/css/all.min.css">
  
  <style>
    body {
      background-color: #f8f9fa;
    }
    .card {
      border-radius: 10px;
      box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
      margin-bottom: 20px;
    }
    .card-header {
      background-color: #343a40;
      color: white;
      border-radius: 10px 10px 0 0 !important;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .badge-qtd {
      font-size: 1em;
      padding: 0.5em 0.8em;
    }
    .user-badge {
      margin-right: 8px;
      margin-bottom: 8px;
      display: inline-flex;
      align-items: center;
    }
    .user-avatar {
      width: 30px;
      height: 30px;
      border-radius: 50%;
      background-color: #6c757d;
      color: white;
      display: flex;
      align-items: center;
      justify-content: center;
      margin-right: 8px;
      font-weight: bold;
    }
    .btn-group-sm .btn {
      padding: 0.25rem 0.5rem;
      font-size: 0.875rem;
    }
    .empty-group {
      color: #6c757d;
      font-style: italic;
    }
  </style>
</head>
<body>
  
  <div class="container py-4">
    <div class="d-flex justify-content-between align-items-center mb-4">
      <h2><i class="fas fa-users mr-2"></i>Lista de Grupos</h2>
      <a href="usr_listar.asp" class="btn btn-primary">
        <i class="fas fa-plus mr-1"></i> Voltar
      </a>
    </div>

    <% 
    Do While Not rsGrupos.EOF 
      grupoId = rsGrupos("ID_Grupo")
      grupoNome = rsGrupos("Nome_Grupo")
      qtdUsuarios = rsGrupos("QtdUsuarios")
    %>
    <div class="card">
      <div class="card-header">
        <h5 class="mb-0"><%=grupoNome%></h5>
        <div>
          <span class="badge badge-primary badge-qtd">
            <i class="fas fa-users mr-1"></i> <%=qtdUsuarios%> usuário(s)
          </span>
        <%if UCase(Session("Usuario")) = "BARRETO" then%>
          <div class="btn-group btn-group-sm ml-2">
            <a href="usr_gerenciar_grupos.asp?id=<%=grupoId%>" class="btn btn-warning" title="Gerenciar Usuários">
              <i class="fas fa-user-edit"></i>
            </a>
            <a href="usr_grupo_edit.asp?id=<%=grupoId%>" class="btn btn-info" title="Editar Grupo">
              <i class="fas fa-edit"></i>
            </a>

            <a href="usr_grupo_excluir.asp?id=<%=grupoId%>" class="btn btn-danger" title="Excluir Grupo" 
               onclick="return confirm('Tem certeza que deseja excluir este grupo?')">
              <i class="fas fa-trash-alt"></i>
            </a>
        <%end if%>
          </div>
        </div>
      </div>
      
      <div class="card-body">
        <%
        ' Obter usuários deste grupo
        Set rsUsuarios = Server.CreateObject("ADODB.Recordset")
        rsUsuarios.Open "SELECT u.UserId, u.Usuario FROM Usuarios u " & _
                        "INNER JOIN Usuario_Grupo ug ON u.UserId = ug.UserId " & _
                        "WHERE ug.ID_Grupo = " & grupoId & " ORDER BY u.Usuario", StrConn
        
        If rsUsuarios.EOF Then
          Response.Write "<div class='empty-group'>Nenhum usuário neste grupo</div>"
        Else
          Do While Not rsUsuarios.EOF
            Dim nomeUsuario, iniciais
            nomeUsuario = UCase(rsUsuarios("Usuario"))
            iniciais = UCase(Left(nomeUsuario, 1))
            If InStr(nomeUsuario, " ") > 0 Then
              iniciais = iniciais & UCase(Mid(nomeUsuario, InStr(nomeUsuario, " ") + 1, 1))
            End If
        %>
        <div class="user-badge">
          <div class="user-avatar"><%=iniciais%></div>
          <span><%=nomeUsuario%></span>
        </div>
        <%
            rsUsuarios.MoveNext
          Loop
        End If
        
        rsUsuarios.Close
        Set rsUsuarios = Nothing
        %>
      </div>
    </div>
    <%
      rsGrupos.MoveNext
    Loop
    
    If rsGrupos.EOF And rsGrupos.BOF Then
      Response.Write "<div class='alert alert-info'>Nenhum grupo cadastrado</div>"
    End If
    %>
  </div>

  <!-- jQuery first, then Popper.js, then Bootstrap JS -->
  <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.1/dist/umd/popper.min.js"></script>
  <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
</body>
</html>

<%
rsGrupos.Close
Set rsGrupos = Nothing
%>