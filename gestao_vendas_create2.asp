<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<%
    Function RemoverNumeros(texto)
        Dim regex
        Set regex = New RegExp
        
        ' Remove números (0-9) e asteriscos (*)
        regex.Pattern = "[0-9*-]"
        regex.Global = True
        
        RemoverNumerosEAsteriscos = regex.Replace(texto, "")
        
        ' Remove espaços extras (opcional)
        RemoverNumeros = Trim(Replace(RemoverNumerosEAsteriscos, "  ", " "))
    End Function    

Function FormatNumberForSQL(sValue)
    ' Remove o separador de milhares (o ponto)
    sValue = Replace(sValue, ".", "")
    ' Substitui o separador decimal (a vírgula) por um ponto
    sValue = Replace(sValue, ",", ".")
    FormatNumberForSQL = sValue
End Function    
%>

<%
' -----------------------------------------------------------------------------------
' INICIALIZAÇÃO E CONEXÃO COM BANCOS DE DADOS
' -----------------------------------------------------------------------------------
' Verifica se as strings de conexão estão configuradas.
If Len(StrConn) = 0 Or Len(StrConnSales) = 0 Then
    Response.Write "Erro: Conexões com bancos de dados não configuradas"
    Response.End
End If

' Cria e abre as conexões com os bancos de dados.
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConn

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' -----------------------------------------------------------------------------------
' PROCESSAMENTO DO FORMULÁRIO (MÉTODO POST)
' -----------------------------------------------------------------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' Declaração das variáveis.
    Dim empreend_id, unidade, corretorId, valorUnidade, comissaoPercentual
    Dim dataVenda, obs, usuario, m2, diretoriaId, gerenciaId
    Dim comissaoDiretoria, comissaoGerencia, comissaoCorretor, trimestre
    Dim nomeEmpreendimento, corretorNome, diretoriaNome, gerenciaNome
    Dim valorComissaoGeral, valorComissaoDiretoria, valorComissaoGerencia, valorComissaoCorretor
    Dim sqlVendas, sqlComissoes, vendaId
    
    ' Coleta e formatação dos dados do formulário.
    ' A função `GetFormattedNumber` centraliza a lógica de formatação.
    empreend_id = Request.Form("empreend_id")
    unidade = Request.Form("unidade")
    corretorId = Request.Form("corretorId")
    diretoriaId = Request.Form("diretoriaId")
    gerenciaId = Request.Form("gerenciaId")
    trimestre = Request.Form("trimestre")
    dataVenda = Request.Form("dataVenda")
    obs = Request.Form("obs")
    m2 = GetFormattedNumber(Request.Form("m2"))
    valorUnidade = GetFormattedNumber(Request.Form("valorUnidade"))
    comissaoPercentual = GetFormattedNumber(Request.Form("comissaoPercentual"))
    comissaoDiretoria = GetFormattedNumber(Request.Form("comissaoDiretoria"))
    comissaoGerencia = GetFormattedNumber(Request.Form("comissaoGerencia"))
    comissaoCorretor = GetFormattedNumber(Request.Form("comissaoCorretor"))
    usuario = Session("Usuario")
    
    ' Validação de dados numéricos essenciais.
    If Not IsNumeric(valorUnidade) Or Not IsNumeric(comissaoPercentual) Then
        Response.Write "<script>alert('Valores inválidos!');history.back();</script>"
        Response.End
    End If

    ' A função `GetDataFromDB` centraliza a busca de dados no banco,
    ' evitando a repetição de código para cada Recordset.
    nomeEmpreendimento = GetDataFromDB(conn, "Empreendimento", "NomeEmpreendimento", "Empreend_ID", empreend_id)
    corretorNome = GetDataFromDB(conn, "Usuarios", "Nome", "UserId", corretorId)
    diretoriaNome = GetDataFromDB(conn, "Diretorias", "NomeDiretoria", "DiretoriaID", diretoriaId)
    
    ' Trata o caso onde a gerência pode não ser selecionada.
    If gerenciaId <> "" And IsNumeric(gerenciaId) Then
        gerenciaNome = GetDataFromDB(conn, "Gerencias", "NomeGerencia", "GerenciaID", gerenciaId)
    Else
        gerenciaNome = "Não aplicável"
        gerenciaId = 0
    End If
    
    ' Extrai ano, mês, dia e calcula o trimestre da data de venda.
    Dim anoVenda, mesVenda, diaVenda
    If Trim(dataVenda) <> "" Then
        anoVenda = Year(dataVenda)
        mesVenda = Month(dataVenda)
        diaVenda = Day(dataVenda)
        If Trim(trimestre) = "" Then trimestre = Int((mesVenda - 1) / 3) + 1
    Else
        ' Se a data de venda está vazia, usa a data e hora atuais.
        dataVenda = Now()
        anoVenda = Year(dataVenda)
        mesVenda = Month(dataVenda)
        diaVenda = Day(dataVenda)
        trimestre = Int((mesVenda - 1) / 3) + 1
    End If

    ' Formatação da data para o SQL.
    dataVendaSQL = FormatDateTimeForSQL(dataVenda)
    dataRegistroSQL = FormatDateTimeForSQL(Now())

    ' Cálculo das comissões.
    vFatorDivisao = 10000
    valorComissaoGeral = valorUnidade * (comissaoPercentual / vFatorDivisao)
    valorComissaoDiretoria = valorComissaoGeral * (comissaoDiretoria / vFatorDivisao)
    valorComissaoGerencia = valorComissaoGeral * (comissaoGerencia / vFatorDivisao)
    valorComissaoCorretor = valorComissaoGeral * (comissaoCorretor / vFatorDivisao)

    ' -----------------------------------------------------------------------------------
    ' INSERÇÃO NO BANCO DE DADOS
    ' -----------------------------------------------------------------------------------
    ' Inserção na tabela Vendas.
    sqlVendas = "INSERT INTO Vendas (" & _
    "Empreend_ID, NomeEmpreendimento, Unidade, UnidadeM2, Corretor, CorretorId, " & _
    "ValorUnidade, ComissaoPercentual, ValorComissaoGeral, DataVenda, " & _
    "DiaVenda, MesVenda, AnoVenda, Trimestre, Obs, Usuario, " & _
    "DiretoriaId, Diretoria, GerenciaId, Gerencia, " & _
    "ComissaoDiretoria, ValorDiretoria, " & _
    "ComissaoGerencia, ValorGerencia, " & _
    "ComissaoCorretor, ValorCorretor) " & _
    "VALUES (" & empreend_id & ", '" & SanitizeSQL(nomeEmpreendimento) & "', " & _
    "'" & SanitizeSQL(unidade) & "', " & m2 & ", " & _
    "'" & SanitizeSQL(corretorNome) & "', " & corretorId & ", " & _
    valorUnidade & ", " & comissaoPercentual & ", " & valorComissaoGeral & ", " & dataVendaSQL & ", " & _
    diaVenda & ", " & mesVenda & ", " & anoVenda & ", " & trimestre & ", " & _
    "'" & SanitizeSQL(obs) & "', '" & SanitizeSQL(usuario) & "', " & _
    diretoriaId & ", '" & SanitizeSQL(diretoriaNome) & "', " & gerenciaId & ", " & _
    "'" & SanitizeSQL(gerenciaNome) & "', " & comissaoDiretoria & ", " & valorComissaoDiretoria & ", " & _
    comissaoGerencia & ", " & valorComissaoGerencia & ", " & comissaoCorretor & ", " & valorComissaoCorretor & ")"

    connSales.Execute(sqlVendas)

    ' Obtém o ID da venda recém-inserida.
    Set rsLastID = connSales.Execute("SELECT MAX(ID) AS NewID FROM Vendas")
    If Not rsLastID.EOF Then vendaId = rsLastID("NewID")
    rsLastID.Close
    
    '-------- Inserção na tabela COMISSOES_A_PAGAR.
    sqlComissoes = "INSERT INTO COMISSOES_A_PAGAR (" & _
    "ID_Venda, Empreendimento, Unidade, DataVenda, UserIdDiretoria, NomeDiretor, " & _
    "UserIdGerencia, NomeGerente, UserIdCorretor, NomeCorretor, PercDiretoria, ValorDiretoria, " & _
    "PercGerencia, ValorGerencia, PercCorretor, ValorCorretor, TotalComissao, StatusPagamento, Usuario) " & _
    "VALUES (" & vendaId & ", '" & SanitizeSQL(nomeEmpreendimento) & "', '" & SanitizeSQL(unidade) & "', " & _
    dataVendaSQL & ", " & diretoriaId & ", '" & SanitizeSQL(diretoriaNome) & "', " & gerenciaId & ", " & _
    "'" & SanitizeSQL(gerenciaNome) & "', " & corretorId & ", '" & SanitizeSQL(corretorNome) & "', " & _
    comissaoDiretoria & ", " & valorComissaoDiretoria & ", " & comissaoGerencia & ", " & valorComissaoGerencia & ", " & _
    comissaoCorretor & ", " & valorComissaoCorretor & ", " & valorComissaoGeral & ", 'Pendente', '" & SanitizeSQL(usuario) & "')"

    connSales.Execute(sqlComissoes)
    
    ' Redireciona para a página de sucesso após a inserção.
    Response.Redirect "gestao_vendas_list2r.asp?mensagem=Venda cadastrada com sucesso!"
End If

' -----------------------------------------------------------------------------------
' BUSCA DE DADOS PARA DROPDOWNS (MÉTODO GET)
' -----------------------------------------------------------------------------------
' Busca e popula os recordsets para os dropdowns na página.
Set rsEmpreend = conn.Execute("SELECT Empreend_ID, NomeEmpreendimento, ComissaoVenda FROM Empreendimento ORDER BY NomeEmpreendimento")
Set rsDiretorias = conn.Execute("SELECT DiretoriaID, NomeDiretoria FROM Diretorias ORDER BY NomeDiretoria")
Set rsCorretores = conn.Execute("SELECT UserId, Nome FROM Usuarios WHERE Funcao = 'Corretor' AND Nome <> '' ORDER BY Nome")
%>

<% ' -----------------------------------------------------------------------------------
' FUNÇÕES AUXILIARES
' ----------------------------------------------------------------------------------- %>
<%
' Função para formatar números, removendo pontos e substituindo vírgulas por pontos.
Function GetFormattedNumber(sValue)
    GetFormattedNumber = Replace(Replace(sValue, ".", ""), ",", ".")
End Function

' Função para buscar dados de uma tabela com base em um critério.
Function GetDataFromDB(oConn, sTable, sField, sWhereField, sWhereValue)
    Dim sResult
    Set rs = oConn.Execute("SELECT " & sField & " FROM " & sTable & " WHERE " & sWhereField & " = " & sWhereValue)
    If Not rs.EOF Then
        sResult = rs(sField)
    Else
        sResult = "Desconhecido"
    End If
    rs.Close
    Set rs = Nothing
    GetDataFromDB = sResult
End Function

' Função para formatar a data para o formato SQL.
Function FormatDateTimeForSQL(dDate)
    If Trim(dDate & "") = "" Then
        FormatDateTimeForSQL = "NULL"
    Else
        FormatDateTimeForSQL = "'" & FormatDateTime(dDate, 2) & "'"
    End If
End Function

' Função para sanitizar strings, escapando aspas simples.
Function SanitizeSQL(sValue)
    SanitizeSQL = Replace(sValue, "'", "''")
End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="refresh" content="600">
    <title>Nova Venda</title>
    
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    
    <!-- Select2 para selects com busca -->
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    
    <style>
        body {
            background-color: #807777;
            color: #fff;
            padding: 20px;
        }
        .card {
            background-color: #fff;
            color: #000;
            margin-bottom: 20px;
        }
        .card-header {
            background-color: #f8f9fa;
            font-weight: bold;
        }
        .btn-maroon {
            background-color: #800000;
            color: white;
        }
        .btn-maroon:hover {
            background-color: #a00;
            color: white;
        }
        .comissao-result {
            font-weight: bold;
            color: #17a2b8;
        }
        .comissao-dist {
            font-size: 0.9rem;
            color: #6c757d;
        }
        .error-message {
            color: #dc3545;
            font-size: 0.875em;
        }
        
        /* CORREÇÃO PARA OS SELECTS */
        .select2-container--default .select2-selection--single,
        .select2-container--default .select2-selection--multiple {
            background-color: #fff;
            color: #000;
            border: 1px solid #ced4da;
        }
        .select2-container--default .select2-selection--single .select2-selection__rendered {
            color: #000;
        }
        .select2-container--default .select2-selection--single .select2-selection__placeholder {
            color: #6c757d;
        }
        .select2-dropdown {
            background-color: #fff;
            color: #000;
        }
        .select2-container--default .select2-results__option[aria-selected=true] {
            background-color: #f8f9fa;
            color: #000;
        }
        .select2-container--default .select2-results__option--highlighted[aria-selected] {
            background-color: #007bff;
            color: #fff;
        }
    </style>
</head>
<body>
    <div class="container" style="padding-top: 70px;">
        <h2 class="mt-4 mb-4"><i class="fas fa-plus-circle"></i> Nova Venda</h2>
        
        <button type="button" onclick="window.close();" class="btn btn-success">
            <i class="fas fa-times me-2"></i>Fechar
        </button><br>

        <form method="post" id="formVenda">
            <!-- Campos hidden para dia, mês e ano -->
            <input type="hidden" id="diaVenda" name="diaVenda">
            <input type="hidden" id="mesVenda" name="mesVenda">
            <input type="hidden" id="anoVenda" name="anoVenda">
            
            <!-- Card Empreendimento -->
            <div class="card">
                <div class="card-header">Empreendimento</div>
                <div class="card-body">
                    <div class="row mb-3">
                        <div class="col-md-6">
                            <label for="empreend_id" class="form-label">Empreendimento *</label>
                            <select class="form-select select2" id="empreend_id" name="empreend_id" required>
                                <option value="">Selecione...</option>
                                <% 
                                If Not rsEmpreend.EOF Then
                                    rsEmpreend.MoveFirst
                                    Do While Not rsEmpreend.EOF 
                                %>
                                    <option value="<%= rsEmpreend("Empreend_ID") %>" data-comissao="<%= rsEmpreend("ComissaoVenda") %>">
                                        <%= RemoverNumeros(rsEmpreend("NomeEmpreendimento")) %>
                                    </option>
                                <%
                                        rsEmpreend.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                        <div class="col-md-3">
                            <label for="unidade" class="form-label">Unidade *</label>
                            <input type="text" class="form-control" id="unidade" name="unidade" required>
                        </div>
                        <div class="col-md-3">
                            <label for="m2" class="form-label">M² *</label>
                            <input type="text" class="form-control" id="m2" name="m2" required>
                        </div>
                    </div>
                    
                    <div class="row">
                        <div class="col-md-6">
                            <label for="valorUnidade" class="form-label">Valor da Unidade (R$) *</label>
                            <input type="text" class="form-control" id="valorUnidade" name="valorUnidade" required>
                        </div>
                        <div class="col-md-3">
                            <label for="comissaoPercentual" class="form-label">% Comissão *</label>
                            <div class="input-group">
                                <input type="text" class="form-control" id="comissaoPercentual" name="comissaoPercentual" required>
                                <span class="input-group-text">%</span>
                            </div>
                        </div>
                        <div class="col-md-3">
                            <label class="form-label">Valor Comissão</label>
                            <div class="form-control comissao-result" id="valorComissaoText">R$ 0,00</div>
                            <input type="hidden" id="valorComissaoHidden" name="valorComissao">
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Card Vendido Por -->
            <div class="card">
                <div class="card-header">Vendido Por</div>
                <div class="card-body">
                    <div class="row mb-3">
                        <div class="col-md-4">
                            <label for="diretoriaId" class="form-label">Diretoria *</label>
                            <select class="form-select" id="diretoriaId" name="diretoriaId" required>
                                <option value="">Selecione...</option>
                                <% 
                                If Not rsDiretorias.EOF Then
                                    rsDiretorias.MoveFirst
                                    Do While Not rsDiretorias.EOF 
                                %>
                                    <option value="<%= rsDiretorias("DiretoriaID") %>"><%= rsDiretorias("NomeDiretoria") %></option>
                                <%
                                        rsDiretorias.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                        <div class="col-md-4">
                            <label for="gerenciaId" class="form-label">Gerência *</label>
                            <select class="form-select" id="gerenciaId" name="gerenciaId" required disabled>
                                <option value="">Selecione uma diretoria primeiro</option>
                            </select>
                        </div>
                        <div class="col-md-4">
                            <label for="corretorId" class="form-label">Corretor *</label>
                            <select class="form-select select2" id="corretorId" name="corretorId" required>
                                <option value="">Selecione...</option>
                                <% 
                                If Not rsCorretores.EOF Then
                                    rsCorretores.MoveFirst
                                    Do While Not rsCorretores.EOF 
                                %>
                                    <option value="<%= rsCorretores("UserId") %>"><%= rsCorretores("Nome") %></option>
                                <%
                                        rsCorretores.MoveNext
                                    Loop
                                End If
                                %>
                            </select>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Card Comissão -->
            <div class="card">
                <div class="card-header">Distribuição da Comissão</div>
                <div class="card-body">
                    <div class="row mb-3">
                        <div class="col-md-3">
                            <label for="comissaoDiretoria" class="form-label">% Diretoria</label>
                            <div class="input-group">
                                <input type="text" class="form-control" id="comissaoDiretoria" name="comissaoDiretoria" value="5,00">
                                <span class="input-group-text">%</span>
                            </div>
                            <input type="text" class="form-control mt-2" id="valorComissaoDiretoria" name="valorComissaoDiretoria" readonly>
                        </div>
                        <div class="col-md-3">
                            <label for="comissaoGerencia" class="form-label">% Gerência</label>
                            <div class="input-group">
                                <input type="text" class="form-control" id="comissaoGerencia" name="comissaoGerencia" value="10,00">
                                <span class="input-group-text">%</span>
                            </div>
                            <input type="text" class="form-control mt-2" id="valorComissaoGerencia" name="valorComissaoGerencia" readonly>
                        </div>
                        <div class="col-md-3">
                            <label for="comissaoCorretor" class="form-label">% Corretor</label>
                            <div class="input-group">
                                <input type="text" class="form-control" id="comissaoCorretor" name="comissaoCorretor" value="35,00">
                                <span class="input-group-text">%</span>
                            </div>
                            <input type="text" class="form-control mt-2" id="valorComissaoCorretor" name="valorComissaoCorretor" readonly>
                        </div>
                        <div class="col-md-3">
                            <label class="form-label">Total Comissão</label>
                            <input type="text" class="form-control" id="valorComissaoSoma" name="valorComissaoSoma" readonly>
                            <div id="comissaoError" class="error-message mt-2"></div>
                        </div>
                    </div>
                </div>
            </div>
            
            <div class="card">
                <div class="card-header">Outras Informações</div>
                <div class="card-body">
                    <div class="row mb-3">
                        <div class="col-md-3">
                            <label for="dataVenda" class="form-label">Data da Venda *</label>
                            <input type="date" class="form-control" id="dataVenda" name="dataVenda" required>
                        </div>
                        <div class="col-md-3">
                            <label for="trimestre" class="form-label">Trimestre</label>
                            <select class="form-select" id="trimestre" name="trimestre">
                                <option value="">Selecione...</option>
                                <option value="1">1º Trimestre</option>
                                <option value="2">2º Trimestre</option>
                                <option value="3">3º Trimestre</option>
                                <option value="4">4º Trimestre</option>
                            </select>
                        </div>
                        <div class="col-md-6">
                            <label for="obs" class="form-label">Observações</label>
                            <textarea class="form-control" id="obs" name="obs" rows="3"></textarea>
                        </div>
                    </div>
                    
                    <div class="d-grid gap-2 d-md-flex justify-content-md-end">
                        <a href="gestao_vendas_list2r.asp" class="btn btn-secondary me-md-2"><i class="fas fa-times"></i> Cancelar</a>
                        <button type="submit" class="btn btn-success"><i class="fas fa-save"></i> Salvar Venda</button>
                    </div>
                </div>
            </div>
        </form>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    
    <!-- jQuery e jQuery Mask -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.16/jquery.mask.min.js"></script>
    
    <!-- Select2 -->
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/i18n/pt-BR.js"></script>
    
    <script>
        $(document).ready(function() {
            // Inicializa select2 nos selects
            $('.select2').select2({
                language: "pt-BR",
                placeholder: "Selecione...",
                allowClear: true
            });
            
            // Máscaras para os campos
            $('#valorUnidade').mask('#.##0,00', {reverse: true});
            $('#comissaoPercentual, #comissaoDiretoria, #comissaoGerencia, #comissaoCorretor').mask('##0,00', {reverse: true});
            $('#m2').mask('#0,00', {reverse: true});
            
            // Formata campos monetários como somente leitura
            $('#valorComissaoDiretoria, #valorComissaoGerencia, #valorComissaoCorretor').mask('#.##0,00', {reverse: true, readonly: true});
            
            // Carrega gerencias quando seleciona diretoria
            $('#diretoriaId').change(function() {
                var diretoriaId = $(this).val();
                if (diretoriaId) {
                    $('#gerenciaId').prop('disabled', false);
                    $.getJSON('get_gerencias.asp', {diretoriaId: diretoriaId}, function(data) {
                        var options = '<option value="">Selecione...</option>';
                        $.each(data, function(key, val) {
                            options += '<option value="' + val.GerenciaID + '">' + val.NomeGerencia + '</option>';
                        });
                        $('#gerenciaId').html(options);
                    });
                } else {
                    $('#gerenciaId').prop('disabled', true).html('<option value="">Selecione uma diretoria primeiro</option>');
                }
            });
            
            // Preenche comissão padrão quando seleciona empreendimento
            $('#empreend_id').change(function() {
                var selected = $(this).find('option:selected');
                var comissao = selected.data('comissao');
                if (comissao) {
                    $('#comissaoPercentual').val(comissao.toString().replace('.', ',')).trigger('input');
                }
            });
            
            // Atualiza dia, mês, ano e trimestre quando seleciona data
            $('#dataVenda').change(function() {
                var data = new Date($(this).val());
                if (!isNaN(data.getTime())) {
                    $('#diaVenda').val(data.getDate());
                    $('#mesVenda').val(data.getMonth() + 1);
                    $('#anoVenda').val(data.getFullYear());
                    
                    // Calcula o trimestre
                    var mes = data.getMonth() + 1;
                    var trimestre = Math.floor((mes - 1) / 3) + 1;
                    $('#trimestre').val(trimestre);
                }
            });
            
            // Função para validar números
            function validarNumero(valor) {
                valor = valor.replace(/\./g, '').replace(',', '.');
                return !isNaN(parseFloat(valor)) && isFinite(valor);
            }
            
            // Calcula a comissão
            function calcularComissoes() {
                try {
                    // Valores principais
                    var valorInput = $('#valorUnidade').val();
                    var percentualInput = $('#comissaoPercentual').val();
                    
                    // Remove pontos e substitui vírgula por ponto para cálculo
                    var valor = parseFloat(valorInput.replace(/\./g, '').replace(',', '.')) || 0;
                    var percentual = parseFloat(percentualInput.replace(',', '.')) || 0;
                    
                    // Cálculo da comissão total
                    var comissaoTotal = valor * (percentual / 100);
                    
                    // Valores das comissões parciais
                    var percDiretoria = parseFloat($('#comissaoDiretoria').val().replace(',', '.')) || 0;
                    var percGerencia = parseFloat($('#comissaoGerencia').val().replace(',', '.')) || 0;
                    var percCorretor = parseFloat($('#comissaoCorretor').val().replace(',', '.')) || 0;
                    
                    // Cálculo dos valores parciais
                    var comissaoDiretoria = comissaoTotal * (percDiretoria / 100);
                    var comissaoGerencia = comissaoTotal * (percGerencia / 100);
                    var comissaoCorretor = comissaoTotal * (percCorretor / 100);
                    
                    // Soma das comissões parciais
                    var totalDistribuido = comissaoDiretoria + comissaoGerencia + comissaoCorretor;
                    
                    // Validação do total distribuído
                    var diferenca = Math.abs(comissaoTotal - totalDistribuido);
                    if (diferenca > 0.01) {
                        $('#comissaoError').text('');
                    } else {
                        $('#comissaoError').text('');
                    }
                    
                    // Formata os valores para exibição
                    $('#valorComissaoText').text('R$ ' + comissaoTotal.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}));
                    $('#valorComissaoHidden').val(comissaoTotal.toFixed(2));
                    
                    $('#valorComissaoDiretoria').val(
                        'R$ ' + comissaoDiretoria.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})
                    );
                    
                    $('#valorComissaoGerencia').val(
                        'R$ ' + comissaoGerencia.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})
                    );
                    
                    $('#valorComissaoCorretor').val(
                        'R$ ' + comissaoCorretor.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})
                    );
                    
                    $('#valorComissaoSoma').val(
                        'R$ ' + totalDistribuido.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})
                    );

                } catch(e) {
                    console.error("Erro no cálculo:", e);
                }
            }
            
            // Configura os eventos para cálculo automático
            $('#valorUnidade, #comissaoPercentual').on('input change', calcularComissoes);
            $('#comissaoDiretoria, #comissaoGerencia, #comissaoCorretor').on('input change', calcularComissoes);
            
            // Calcula a comissão inicial
            calcularComissoes();
        });
    </script>
</body>
</html>
<%
' Fecha conexões e recordsets
If IsObject(rsEmpreend) Then
    rsEmpreend.Close
    Set rsEmpreend = Nothing
End If

If IsObject(rsDiretorias) Then
    rsDiretorias.Close
    Set rsDiretorias = Nothing
End If

If IsObject(rsCorretores) Then
    rsCorretores.Close
    Set rsCorretores = Nothing
End If

If IsObject(conn) Then
    conn.Close
    Set conn = Nothing
End If

If IsObject(connSales) Then
    connSales.Close
    Set connSales = Nothing
End If
%>