<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="usr_acoes.inc"-->
<!--#include file="gestao_header.inc"-->


<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lista de Diretorias</title>
    
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH" crossorigin="anonymous">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link rel="stylesheet" href="css/gestao_estilo.css">
    
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/jquery.dataTables.min.css">
</head>
<body>
    <nav class="navbar navbar-expand-lg">
        <div class="container">
            <a class="navbar-brand" href="#">
                <i class="fas fa-sun me-2"></i>SunnyImob.
            </a>
            <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav">
                <span class="navbar-toggler-icon"></span>
            </button>
            <div class="collapse navbar-collapse" id="navbarNav">
                <ul class="navbar-nav ms-auto">
                    <li class="nav-item">
                        <a class="nav-link" href="gestao_painel2.asp"><i class="fas fa-home me-1"></i> Início</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="#"><i class="fas fa-cog me-1"></i> Configurações</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="gestao_logout.asp"><i class="fas fa-sign-out-alt me-1"></i> Sair</a>
                    </li>
                </ul>
            </div>
        </div>
    </nav>

    <section class="welcome-section text-center">
        <div class="container">
            <h1 class="display-4 mb-2">Tocca Onze</h1>
            <p class="lead">Gerencie as operações de gestão e vendas</p>
        </div>
    </section>

    <div class="container py-5">
        <h2 class="text-center mb-4">Lista de Diretorias</h2>
        <p><a class="btn btn-primary" href="diretoria_create.asp" target="_blank">
            <i class="fas fa-plus me-1"></i> Nova Diretoria
        </a></p>
        
        <div class="table-responsive">
            <table id="tabelaDiretorias" class="display compact nowrap" style="width:100%">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Diretoria</th>
                        <th>Diretor</th>
                        <th>Telefone</th>
                        <th>Usuário</th>
                        <th>Nome do Usuário</th>
                        <th>Estado</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open strConn
sql = "SELECT d.*, u.Nome, u.Usuario FROM Diretorias d LEFT JOIN Usuarios u ON d.UserId = u.UserId"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
%>
                    <tr>
                        <td><%=rs("DiretoriaId")%></td>
                        <td><span title="<%=rs("NomeDiretoria")%>"><%=rs("NomeDiretoria")%></span></td>
                        <td><span title="<%=rs("NomeDiretor")%>"><%=rs("NomeDiretor")%></span></td>
                        <td><%=rs("TelefoneDiretor")%></td>
                        <td><%=lCase(rs("u.Usuario"))%></td>
                        <td><%=rs("u.Nome")%></td>
                        <td><%=rs("EstadoDiretoria")%></td>
                        <td>
                           <a class="btn btn-sm btn-info" href="diretoria_update.asp?id=<%=rs("DiretoriaId")%>" target="_blank">
                                <i class="fas fa-edit me-1"></i> Editar
                            </a>
                        </td>
                    </tr>
<%
rs.MoveNext
Loop
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
                </tbody>
            </table>
        </div>
    </div>
    
    <footer class="text-center mt-auto">
        <div class="container">
            <div class="row">
                <div class="col-md-6">
                    <h5><i class="fas fa-sun me-2"></i>SunnyImob</h5>
                    <p>Valter Barreto</p>
                </div>
                <div class="col-md-6">
                    <p>&copy; 2023 Todos os direitos reservados</p>
                    <div class="social-icons">
                        <a href="#" class="me-2"><i class="fab fa-facebook-f"></i></a>
                        <a href="#" class="me-2"><i class="fab fa-twitter"></i></a>
                        <a href="#" class="me-2"><i class="fab fa-linkedin-in"></i></a>
                        <a href="#"><i class="fab fa-instagram"></i></a>
                    </div>
                </div>
            </div>
        </div>
    </footer>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-YvpcrYf0tY3lHB60NNkmXc5s9fDVZLESaAA55NDzOxhy9GkcIdslK1eN7N6jIeHz" crossorigin="anonymous"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>

    <script>
        $(document).ready(function() {
            $('#tabelaDiretorias').DataTable({
                "language": {
                    "url": "//cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
                },
                "paging": true,
                "lengthChange": true,
                "searching": true,
                "ordering": true,
                "info": true,
                "autoWidth": true,
                "responsive": true
            });
        });
    </script>
</body>
</html>