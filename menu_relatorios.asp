<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="usr_acoes.inc"-->
<!--#include file="gestao_header.inc"-->

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Menu de Relatórios</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link rel="stylesheet" href="css/gestao_estilo.css">
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
        <h2 class="text-center mb-4">Menu de Relatórios</h2>
        <div class="row g-4">

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_vendas_kpi3.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-chart-line me-2"></i>Relat. Vendas KPIs</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização do Valor Geral de Vendas.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Acessar
                            </span>
                        </div>
                    </div>
                </a>
            </div>

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_vendas_top10.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-handshake me-2"></i>Relat. TOP 10</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização do Valor Geral de Vendas.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Acessar
                            </span>
                        </div>
                    </div>
                </a>
            </div>            

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_vendas_kpi5comissao.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-handshake me-2"></i>Relat. Comissões Gerais </h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização do Valor Geral de Vendas.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Acessar
                            </span>
                        </div>
                    </div>
                </a>
            </div>            



         <!--   <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_vendas_relatorio1vendas.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>Relat. Vendas</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Detalhes completos sobre as vendas.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Acessar
                            </span>
                        </div>
                    </div>
                </a>
            </div>  -->


            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_vendas_geral.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>Relat. Geral</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Detalhes completos sobre as vendas.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Acessar
                            </span>
                        </div>
                    </div>
                </a>
            </div>     

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_corretores_mapa_vendas.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-map-marked-alt me-2"></i>Corretor - Mapa de Vendas(QTD)</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização da quantidade de unidades vendidas.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Visualizar Mapa de Vendas
                        </div>
                    </div>
                </a>
            </div>     

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_vendas_corretores.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>Corretor - Extrato de Vendas 1</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização das vendas do corretor.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Visualizar Extrato
                            </span>
                        </div>
                    </div>
                </a>
            </div>   


            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_corretores_comissoes.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>Corretor - Extrato de Vendas 2</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização das comissões mensais dos corretores.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Visualizar Comissões
                            </span>
                        </div>
                    </div>
                </a>
            </div>    


            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_corretores_extrato_comissoes.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>Corretor - Comissão Anual</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização das comissões mensais dos corretores.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Visualizar Comissões
                            </span>
                        </div>
                    </div>
                </a>
            </div>               

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gestao_corretores_comissoes_anual.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>Corretor - Vendas Anual</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Visualização das comissões mensais dos corretores.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Visualizar Comissões
                            </span>
                        </div>
                    </div>
                </a>
            </div>                                                

            <div class="col-12 col-md-6 col-lg-4">
                <a href="diretoria_list.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-user-tie me-2"></i>Listagem de Diretores</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Lista e detalhes dos diretores da empresa.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Acessar
                            </span>
                        </div>
                    </div>
                </a>
            </div>

            <div class="col-12 col-md-6 col-lg-4">
                <a href="gerencia_list.asp" class="text-decoration-none" target="_blank">
                    <div class="card h-100">
                        <div class="card-header text-center">
                            <h5 class="mb-0"><i class="fas fa-users me-2"></i>Listagem de Gerentes</h5>
                        </div>
                        <div class="card-body text-center d-flex flex-column">
                            <p class="card-text">Lista e detalhes dos gerentes de departamento.</p>
                            <span class="btn btn-primary btn-sm mt-auto">
                                <i class="fas fa-arrow-right me-1"></i> Acessar
                            </span>
                        </div>
                    </div>
                </a>
            </div>
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

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>