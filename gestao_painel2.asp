<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if Trim(StrConn)="" then%>
     <!--#include file="conexao.asp"-->
<%end if%>     
<%if Trim(StrConnSales)="" then%>
     <!--#include file="conSunSales.asp"-->
<%end if%>  
<!--#include file="usr_acoes_v4GVendas.inc"-->
<!--#include file="atualizarVendas.asp"-->
<!--#include file="atualizarVendas2.asp"-->

<%
'============================= LOG ============================================'
if (request.ServerVariables("remote_addr") <> "127.0.0.1") AND (request.ServerVariables("remote_addr") <> "::1") then
    set objMail = server.createobject("CDONTS.NewMail")
        objMail.From = "sendmail@gabnetweb.com.br"
        objMail.To   = "sendmail@gabnetweb.com.br, valterpb@hotmail.com"
    objMail.Subject = "SGVendas-" & Ucase(Session("Usuario")) & " - " & request.serverVariables("REMOTE_ADDR") & " - " & Date & " - " & Time
    objMail.MailFormat = 0
    objMail.Body = "Menu Principal"
    objMail.Send
    set objMail = Nothing
end if 
'----------- fim envio de email'

'============================= ATUALIZANDO O BANCO DE DADOS ==================='
Response.Buffer = True
Response.Expires = -1
'On Error Resume Next ' 
' --- CRIAÇÃO DOS OBJETOS ADO DE CONEXÃO ---
Set conn = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")
conn.Open StrConn
connSales.Open StrConnSales

' Primeiro UPDATE: Associar Vendas.DiretoriaId com Diretorias.DiretoriaId e atualizar campos
sqlUpdate1 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Diretorias INNER JOIN Vendas ON Diretorias.DiretoriaId = Vendas.DiretoriaId) SET Vendas.NomeDiretor = [Diretorias].[Nome], Vendas.UserIdDiretoria = [Diretorias].[UserId];"
connSales.Execute(sqlUpdate1)

' UPDATE Gerencias -> Vendas
sqlUpdate2 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Gerencias INNER JOIN Vendas ON Gerencias.GerenciaId = Vendas.GerenciaId) SET [Vendas].[NomeGerente] = [Gerencias].[Nome], [Vendas].[UserIdGerencia] = [Gerencias].[UserId];"
connSales.Execute(sqlUpdate2)

'Atualizar Nome do Corretor-----------------------------'
sqlUpdateCorretor = "UPDATE (Vendas INNER JOIN [;DATABASE=" & dbSunnyPath & "].Usuarios ON Vendas.CorretorId = Usuarios.UserId) " & _
                   "SET Vendas.Corretor = Usuarios.Nome;"
connSales.Execute(sqlUpdateCorretor)

' Esta é a instrução SQL para atualizar o campo Semestre.
sql = "UPDATE Vendas " & _
      "SET Semestre = SWITCH(" & _
      "    Trimestre IN (1, 2), 1, " & _
      "    Trimestre IN (3, 4), 2" & _
      ") " & _
      "WHERE Trimestre IS NOT NULL;"
On Error Resume Next
connSales.Execute sql

' Verificação de erros.
If Err.Number <> 0 Then
    Response.Write "Ocorreu um erro ao atualizar o campo Semestre: " & Err.Description
Else
   ' Response.Write "O campo Semestre foi atualizado com sucesso para todos os registros."
End If
On Error GoTo 0
' ======================= FINAL ATUALIZAÇÃO DO BANCO DE DADOS ========================'
%>

<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="refresh" content="600">
    <title>Menu Administrativo</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link rel="stylesheet" href="css/gestao_estilo.css">
    <style>
        /* Gradient background for title containers with white text */
        .welcome-section, .card-header, footer .col-md-6:first-child {
            background: linear-gradient(45deg, #800020, #A52A2A, #4B0012);
        }
        .welcome-section h1, .card-header h5, footer .col-md-6:first-child h5 {
            color: white;
        }
    </style>
</head>
<body>

<%
if not UsuarioGestor() and not UsuarioAdmin() then
     Response.Write("<h3>Função habilitada apenas para Gestores do Sistema.</h3>")
     Response.End
End if
%>
    <nav class="navbar navbar-expand-lg">
        <div class="container">
            <a class="navbar-brand" href="#">
                <i class="fas fa-sun me-2"></i>SGVendas.
            </a>
            <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav">
                <span class="navbar-toggler-icon"></span>
            </button>
            <div class="collapse navbar-collapse" id="navbarNav">
                <ul class="navbar-nav ms-auto">
                    <li class="nav-item">
                        <a class="nav-link active" href="dashboard3rand1.asp"><i class="fas fa-home me-1"></i> Início</a>
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

    <div class="container mb-5">
        <div class="row g-4">
            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>Dashboard</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Dashboard.</p>
                        <a href="dashboard3rand1.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-funnel-dollar me-2"></i>Vendas</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Gerenciamento de Vendas</p>
                        <a href="gestao_vendas_list2r.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-file-alt me-2"></i>Relatórios</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Relatórios gerenciais e consolidados.</p>
                        <a href="menu_relatorios.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-sitemap me-2"></i>Diretorias</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Cadastro e gerenciamento das diretorias da empresa.</p>
                        <a href="diretoria_list.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>Gerentes</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Cadastro e acompanhamento dos gerentes de departamento.</p>
                        <a href="gerencia_list.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>

            <div class="col-md-6 col-lg-4">
                <div class="card">
                    <div class="card-header text-center">
                        <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>Usuários</h5>
                    </div>
                    <div class="card-body text-center d-flex flex-column">
                        <p class="card-text">Cadastro de usuários.</p>
                        <a href="usrv_gestao_listar.asp" class="btn btn-primary btn-sm mt-auto" target="_blank">
                            <i class="fas fa-arrow-right me-1"></i> Acessar
                        </a>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <footer class="text-center mt-auto">
        <div class="container">
            <div class="row">
                <div class="col-md-12">
                    <p><small>Valter Barreto</p>
                    <p>&copy; 2025 Todos os direitos reservados</p></smaal>
                    <div class="social-icons">

                    </div>
                </div>
            </div>
        </div>
    </footer>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>