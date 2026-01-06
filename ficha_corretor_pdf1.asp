<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                  -->
<!-- Data: 04/12/2025                       -->
<!-- CODIGO_ARQUIVO: FICHA_CORRETOR_PDF     -->
<!-- OBS: Gerar PDF do Dossiê do Corretor   -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->
<!--#include file="conexao.asp"-->

<%
' ===============================================
' CONFIGURAÇÃO UTF-8
' ===============================================
Response.CodePage = 65001 ' UTF-8
Response.CharSet = "UTF-8"
Response.ContentType = "text/html; charset=UTF-8"

' ===============================================
' INICIALIZAÇÃO DE VARIÁVEIS GLOBAIS
' ===============================================
' Garantir que todas as variáveis de coleção existam
Dim mesesComVendas, mesesSemVendas
Set mesesComVendas = Server.CreateObject("Scripting.Dictionary")
Set mesesSemVendas = Server.CreateObject("Scripting.Dictionary")

' Inicializar Dictionaries novos
Dim empreendimentosList, localidadesList, topVendasList
Set empreendimentosList = Server.CreateObject("Scripting.Dictionary")
Set localidadesList = Server.CreateObject("Scripting.Dictionary")
Set topVendasList = Server.CreateObject("Scripting.Dictionary")

' ===============================================
' OBTER PARÂMETROS DE FILTRO
' ===============================================
Dim filtroAno, filtroCorretor
filtroAno = Trim(Request.QueryString("ano"))
filtroCorretor = Trim(Request.QueryString("corretor"))

If filtroAno = "" Then 
    Response.Write "<h2>Erro: Ano não especificado.</h2>"
    Response.Write "<p>Por favor, selecione um ano.</p>"
    Response.End
End If

If filtroCorretor = "" Or LCase(filtroCorretor) = "todos" Then 
    Response.Write "<h2>Erro: Corretor não especificado.</h2>"
    Response.Write "<p>Por favor, selecione um corretor específico.</p>"
    Response.End
End If

' ===============================================
' CONFIGURAÇÃO DE BANCO DE DADOS
' ===============================================
Dim connSales
Set connSales = Server.CreateObject("ADODB.Connection")

On Error Resume Next
connSales.Open StrConnSales

If Err.Number <> 0 Then
    Response.Write "<h2>Erro ao conectar ao banco de dados: " & Err.Description & "</h2>"
    Response.End
End If

' ===============================================
' FUNÇÕES UTILITÁRIAS
' ===============================================
Function ConverterValor(valorString)
    On Error Resume Next
    
    Dim valorConvertido
    valorConvertido = 0
    
    If Not IsNull(valorString) And Trim(valorString) <> "" Then
        Dim valorTemp
        valorTemp = Trim(valorString)
        
        ' Remove caracteres não numéricos
        valorTemp = Replace(valorTemp, "R$", "")
        valorTemp = Replace(valorTemp, "$", "")
        valorTemp = Replace(valorTemp, ".", "")
        valorTemp = Replace(valorTemp, ",", ".")
        valorTemp = Trim(valorTemp)
        
        ' Tenta converter
        If IsNumeric(valorTemp) Then
            valorConvertido = CDbl(valorTemp)
        End If
    End If
    
    If Err.Number <> 0 Then
        valorConvertido = 0
        Err.Clear
    End If
    
    On Error GoTo 0
    ConverterValor = valorConvertido
End Function

Function IsNumericValue(valor)
    On Error Resume Next
    Dim resultado
    resultado = False
    
    If Not IsNull(valor) Then
        If IsNumeric(valor) Then
            resultado = True
        Else
            ' Tenta limpar e verificar
            Dim valorTemp
            valorTemp = CStr(valor)
            valorTemp = Replace(valorTemp, "R$", "")
            valorTemp = Replace(valorTemp, "$", "")
            valorTemp = Replace(valorTemp, ".", "")
            valorTemp = Replace(valorTemp, ",", ".")
            valorTemp = Trim(valorTemp)
            
            If IsNumeric(valorTemp) Then
                resultado = True
            End If
        End If
    End If
    
    IsNumericValue = resultado
    On Error GoTo 0
End Function

' Função para verificar se um objeto Dictionary existe e tem itens
Function HasItems(objDict)
    On Error Resume Next
    Dim result
    result = False
    
    If Not objDict Is Nothing Then
        If objDict.Count > 0 Then
            result = True
        End If
    End If
    
    If Err.Number <> 0 Then
        result = False
        Err.Clear
    End If
    
    HasItems = result
    On Error GoTo 0
End Function

' Array com nomes dos meses
Dim arrMesesNome(12)
arrMesesNome(1) = "Janeiro"
arrMesesNome(2) = "Fevereiro"
arrMesesNome(3) = "Março"
arrMesesNome(4) = "Abril"
arrMesesNome(5) = "Maio"
arrMesesNome(6) = "Junho"
arrMesesNome(7) = "Julho"
arrMesesNome(8) = "Agosto"
arrMesesNome(9) = "Setembro"
arrMesesNome(10) = "Outubro"
arrMesesNome(11) = "Novembro"
arrMesesNome(12) = "Dezembro"

' ===============================================
' OBTER DADOS DO CORRETOR
' ===============================================
Dim sqlSafeCorretor
sqlSafeCorretor = Replace(filtroCorretor, "'", "''")

' 1. DADOS GERAIS
Dim sqlDadosGerais, rsDadosGerais
Dim totalVendas, totalVGV, totalComissao, mediaPorVenda, percentualComissao

sqlDadosGerais = "SELECT " & _
                 "COUNT(*) as QtdVendas, " & _
                 "SUM(ValorUnidade) as TotalVGV, " & _
                 "SUM(ValorCorretor) as TotalComissao " & _
                 "FROM Vendas " & _
                 "WHERE Excluido = 0 AND AnoVenda = " & filtroAno & " AND Corretor = '" & sqlSafeCorretor & "'"

Set rsDadosGerais = Server.CreateObject("ADODB.Recordset")
On Error Resume Next
rsDadosGerais.Open sqlDadosGerais, connSales

If Err.Number <> 0 Then
    Response.Write "<h2>Erro na consulta: " & Err.Description & "</h2>"
    Response.Write "<p>SQL: " & Server.HTMLEncode(sqlDadosGerais) & "</p>"
    Response.End
End If

totalVendas = 0
totalVGV = 0
totalComissao = 0

If Not rsDadosGerais.EOF Then
    ' Tratamento seguro para totalVendas
    If Not IsNull(rsDadosGerais("QtdVendas")) Then
        totalVendas = CLng(rsDadosGerais("QtdVendas"))
    End If
    
    ' Tratamento seguro para totalVGV
    If Not IsNull(rsDadosGerais("TotalVGV")) Then
        If IsNumericValue(rsDadosGerais("TotalVGV")) Then
            totalVGV = ConverterValor(rsDadosGerais("TotalVGV"))
        Else
            totalVGV = 0
        End If
    End If
    
    ' Tratamento seguro para totalComissao
    If Not IsNull(rsDadosGerais("TotalComissao")) Then
        If IsNumericValue(rsDadosGerais("TotalComissao")) Then
            totalComissao = ConverterValor(rsDadosGerais("TotalComissao"))
        Else
            totalComissao = 0
        End If
    End If
End If

rsDadosGerais.Close
Set rsDadosGerais = Nothing

' Calcular médias com tratamento de erro
On Error Resume Next
If totalVendas > 0 And IsNumeric(totalVGV) And IsNumeric(totalVendas) Then
    mediaPorVenda = totalVGV / totalVendas
Else
    mediaPorVenda = 0
End If

If totalVGV > 0 And IsNumeric(totalComissao) And IsNumeric(totalVGV) Then
    percentualComissao = (totalComissao / totalVGV) * 100
Else
    percentualComissao = 0
End If

If Err.Number <> 0 Then
    mediaPorVenda = 0
    percentualComissao = 0
    Err.Clear
End If
On Error GoTo 0

' 2. DADOS MENSAIS
Dim sqlMensal, rsMensal

sqlMensal = "SELECT " & _
            "MesVenda, " & _
            "COUNT(*) as QtdVendas, " & _
            "SUM(ValorUnidade) as TotalVGV, " & _
            "SUM(ValorCorretor) as TotalComissao " & _
            "FROM Vendas " & _
            "WHERE Excluido = 0 AND AnoVenda = " & filtroAno & " AND Corretor = '" & sqlSafeCorretor & "' " & _
            "GROUP BY MesVenda " & _
            "ORDER BY MesVenda"

Set rsMensal = Server.CreateObject("ADODB.Recordset")
rsMensal.Open sqlMensal, connSales

Do While Not rsMensal.EOF
    Dim mesNum, qtdVendasMes, totalVGVMes, totalComissaoMes
    mesNum = CStr(rsMensal("MesVenda"))
    
    ' Tratamento seguro para qtdVendasMes
    If Not IsNull(rsMensal("QtdVendas")) Then
        qtdVendasMes = CLng(rsMensal("QtdVendas"))
    Else
        qtdVendasMes = 0
    End If
    
    ' Tratamento seguro para totalVGVMes
    totalVGVMes = 0
    If Not IsNull(rsMensal("TotalVGV")) Then
        If IsNumericValue(rsMensal("TotalVGV")) Then
            totalVGVMes = ConverterValor(rsMensal("TotalVGV"))
        End If
    End If
    
    ' Tratamento seguro para totalComissaoMes
    totalComissaoMes = 0
    If Not IsNull(rsMensal("TotalComissao")) Then
        If IsNumericValue(rsMensal("TotalComissao")) Then
            totalComissaoMes = ConverterValor(rsMensal("TotalComissao"))
        End If
    End If
    
    mesesComVendas.Add mesNum, Array(qtdVendasMes, totalVGVMes, totalComissaoMes)
    
    rsMensal.MoveNext
Loop

rsMensal.Close
Set rsMensal = Nothing

' Identificar meses sem vendas
For i = 1 To 12
    If Not mesesComVendas.Exists(CStr(i)) Then
        mesesSemVendas.Add CStr(i), arrMesesNome(i)
    End If
Next

' 3. EMPREENDIMENTOS - USANDO DICTIONARY
Dim sqlEmpreendimentos, rsEmpreendimentos

sqlEmpreendimentos = "SELECT " & _
                    "NomeEmpreendimento, " & _
                    "COUNT(*) as QtdVendas, " & _
                    "SUM(ValorUnidade) as TotalVGV " & _
                    "FROM Vendas " & _
                    "WHERE Excluido = 0 AND AnoVenda = " & filtroAno & " AND Corretor = '" & sqlSafeCorretor & "' " & _
                    "AND NomeEmpreendimento IS NOT NULL AND NomeEmpreendimento <> '' " & _
                    "GROUP BY NomeEmpreendimento " & _
                    "ORDER BY SUM(ValorUnidade) DESC"

Set rsEmpreendimentos = Server.CreateObject("ADODB.Recordset")
rsEmpreendimentos.Open sqlEmpreendimentos, connSales

Do While Not rsEmpreendimentos.EOF
    Dim nomeEmp, qtdEmp, vgvEmp, percentEmp
    
    ' Nome do empreendimento
    nomeEmp = "Não informado"
    If Not IsNull(rsEmpreendimentos("NomeEmpreendimento")) Then
        nomeEmp = Trim(CStr(rsEmpreendimentos("NomeEmpreendimento")))
    End If
    
    ' Quantidade de vendas
    qtdEmp = 0
    If Not IsNull(rsEmpreendimentos("QtdVendas")) And IsNumeric(rsEmpreendimentos("QtdVendas")) Then
        qtdEmp = CLng(rsEmpreendimentos("QtdVendas"))
    End If
    
    ' VGV Total
    vgvEmp = 0
    If Not IsNull(rsEmpreendimentos("TotalVGV")) Then
        vgvEmp = ConverterValor(rsEmpreendimentos("TotalVGV"))
    End If
    
    ' Percentual do VGV total
    percentEmp = 0
    If IsNumeric(totalVGV) And totalVGV > 0 And IsNumeric(vgvEmp) And vgvEmp > 0 Then
        percentEmp = (vgvEmp / totalVGV) * 100
    End If
    
    ' Armazenar no Dictionary
    empreendimentosList.Add empreendimentosList.Count, Array(nomeEmp, qtdEmp, vgvEmp, percentEmp)
    
    rsEmpreendimentos.MoveNext
Loop

rsEmpreendimentos.Close
Set rsEmpreendimentos = Nothing

' 4. LOCALIDADES - USANDO DICTIONARY
Dim sqlLocalidades, rsLocalidades

sqlLocalidades = "SELECT " & _
                "Localidade, " & _
                "COUNT(*) as QtdVendas, " & _
                "SUM(ValorUnidade) as TotalVGV " & _
                "FROM Vendas " & _
                "WHERE Excluido = 0 AND AnoVenda = " & filtroAno & " AND Corretor = '" & sqlSafeCorretor & "' " & _
                "AND Localidade IS NOT NULL AND Localidade <> '' " & _
                "GROUP BY Localidade " & _
                "ORDER BY SUM(ValorUnidade) DESC"

Set rsLocalidades = Server.CreateObject("ADODB.Recordset")
rsLocalidades.Open sqlLocalidades, connSales

Do While Not rsLocalidades.EOF
    Dim nomeLocal, qtdLocal, vgvLocal, percentLocal
    
    ' Localidade
    nomeLocal = "Não informado"
    If Not IsNull(rsLocalidades("Localidade")) Then
        nomeLocal = Trim(CStr(rsLocalidades("Localidade")))
    End If
    
    ' Quantidade de vendas
    qtdLocal = 0
    If Not IsNull(rsLocalidades("QtdVendas")) And IsNumeric(rsLocalidades("QtdVendas")) Then
        qtdLocal = CLng(rsLocalidades("QtdVendas"))
    End If
    
    ' VGV Total
    vgvLocal = 0
    If Not IsNull(rsLocalidades("TotalVGV")) Then
        vgvLocal = ConverterValor(rsLocalidades("TotalVGV"))
    End If
    
    ' Percentual do VGV total
    percentLocal = 0
    If IsNumeric(totalVGV) And totalVGV > 0 And IsNumeric(vgvLocal) And vgvLocal > 0 Then
        percentLocal = (vgvLocal / totalVGV) * 100
    End If
    
    ' Armazenar no Dictionary
    localidadesList.Add localidadesList.Count, Array(nomeLocal, qtdLocal, vgvLocal, percentLocal)
    
    rsLocalidades.MoveNext
Loop

rsLocalidades.Close
Set rsLocalidades = Nothing

' 5. TOP VENDAS - USANDO DICTIONARY
Dim sqlTopVendas, rsTopVendas

sqlTopVendas = "SELECT TOP 5 " & _
              "DataVenda, " & _
              "NomeEmpreendimento, " & _
              "ValorUnidade " & _
              "FROM Vendas " & _
              "WHERE Excluido = 0 AND AnoVenda = " & filtroAno & " AND Corretor = '" & sqlSafeCorretor & "' " & _
              "ORDER BY ValorUnidade DESC"

Set rsTopVendas = Server.CreateObject("ADODB.Recordset")
rsTopVendas.Open sqlTopVendas, connSales

Do While Not rsTopVendas.EOF
    Dim dataVenda, nomeTop, valorVenda, valorFormatado
    
    ' Data
    dataVenda = ""
    If Not IsNull(rsTopVendas("DataVenda")) Then
        dataVenda = CStr(rsTopVendas("DataVenda"))
    End If
    
    ' Empreendimento
    nomeTop = ""
    If Not IsNull(rsTopVendas("NomeEmpreendimento")) Then
        nomeTop = CStr(rsTopVendas("NomeEmpreendimento"))
    End If
    
    ' Valor
    valorVenda = 0
    If Not IsNull(rsTopVendas("ValorUnidade")) Then
        valorVenda = ConverterValor(rsTopVendas("ValorUnidade"))
    End If
    
    ' Valor formatado
    valorFormatado = "R$ " & FormatNumber(valorVenda, 2)
    
    ' Armazenar no Dictionary
    topVendasList.Add topVendasList.Count, Array(dataVenda, nomeTop, valorVenda, valorFormatado)
    
    rsTopVendas.MoveNext
Loop

rsTopVendas.Close
Set rsTopVendas = Nothing

' ===============================================
' FECHAR CONEXÃO
' ===============================================
connSales.Close
Set connSales = Nothing
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dossiê - <%= filtroCorretor %> - <%= filtroAno %></title>
    <style>
        @media print {
            body {
                font-size: 11pt;
                line-height: 1.4;
                margin: 0.5cm;
                padding: 0;
            }
            .no-print {
                display: none !important;
            }
            .container {
                width: 100%;
                padding: 0;
                margin: 0;
            }
            table {
                page-break-inside: avoid;
                font-size: 10pt;
            }
            .section-title {
                page-break-after: avoid;
            }
        }
        
        body {
            font-family: 'Arial', 'Helvetica', sans-serif;
            color: #333;
            background-color: #fff;
            padding: 15px;
            margin: 0;
            font-size: 13px;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: #fff;
            padding: 20px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
        }
        
        .header {
            text-align: center;
            border-bottom: 3px solid #800000;
            padding-bottom: 15px;
            margin-bottom: 25px;
        }
        
        .header h1 {
            color: #800000;
            font-size: 24px;
            margin: 10px 0 5px 0;
            font-weight: bold;
        }
        
        .header h2 {
            color: #333;
            font-size: 18px;
            margin: 5px 0;
            font-weight: normal;
        }
        
        .corretor-info {
            background: #f5f5f5;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 25px;
            border-left: 5px solid #800000;
        }
        
        .kpi-container {
            display: flex;
            flex-wrap: wrap;
            justify-content: space-between;
            gap: 15px;
            margin-bottom: 25px;
        }
        
        .kpi-box {
            flex: 1;
            min-width: 150px;
            padding: 15px;
            border-radius: 8px;
            color: white;
            text-align: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .kpi-box h3 {
            font-size: 22px;
            margin: 0 0 5px 0;
            font-weight: bold;
        }
        
        .kpi-box p {
            margin: 0;
            font-size: 13px;
            opacity: 0.9;
        }
        
        .bg-vendas { background: linear-gradient(135deg, #007bff, #0056b3); }
        .bg-vgv { background: linear-gradient(135deg, #28a745, #1e7e34); }
        .bg-comissao { background: linear-gradient(135deg, #ffc107, #e0a800); color: #000; }
        .bg-media { background: linear-gradient(135deg, #17a2b8, #138496); }
        
        .section-title {
            color: #800000;
            border-bottom: 2px solid #800000;
            padding-bottom: 8px;
            margin: 30px 0 20px 0;
            font-size: 18px;
            font-weight: bold;
            page-break-after: avoid;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
            font-size: 12px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        
        table th {
            background-color: #800000;
            color: white;
            padding: 10px 12px;
            text-align: left;
            font-weight: bold;
            border: 1px solid #700000;
        }
        
        table td {
            padding: 8px 12px;
            border: 1px solid #ddd;
            vertical-align: top;
        }
        
        table tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        
        table tr:hover {
            background-color: #f1f1f1;
        }
        
        .mes-com-venda {
            display: inline-block;
            background-color: #d4edda;
            color: #155724;
            padding: 5px 10px;
            margin: 3px;
            border-radius: 4px;
            font-size: 11px;
            border: 1px solid #c3e6cb;
        }
        
        .mes-sem-venda {
            display: inline-block;
            background-color: #f8d7da;
            color: #721c24;
            padding: 5px 10px;
            margin: 3px;
            border-radius: 4px;
            font-size: 11px;
            border: 1px solid #f5c6cb;
        }
        
        .valor-destaque {
            font-size: 24px;
            font-weight: bold;
            color: #28a745;
            margin: 5px 0;
        }
        
        .card {
            border: 1px solid #ddd;
            border-radius: 6px;
            margin-bottom: 20px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .card-header {
            background-color: #f8f9fa;
            padding: 12px 15px;
            border-bottom: 1px solid #ddd;
            font-weight: bold;
            font-size: 15px;
            color: #800000;
        }
        
        .card-body {
            padding: 15px;
        }
        
        .footer {
            margin-top: 40px;
            padding-top: 20px;
            border-top: 2px solid #ddd;
            text-align: center;
            color: #666;
            font-size: 12px;
            page-break-before: avoid;
        }
        
        .btn-print {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 1000;
            background: white;
            padding: 8px 15px;
            border-radius: 5px;
            box-shadow: 0 3px 10px rgba(0,0,0,0.2);
            border: 1px solid #ddd;
        }
        
        .btn {
            padding: 8px 15px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 13px;
            font-weight: bold;
            transition: all 0.3s;
            margin: 0 5px;
        }
        
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        }
        
        .btn-danger { 
            background: linear-gradient(135deg, #dc3545, #c82333);
            color: white; 
        }
        
        .btn-secondary { 
            background: linear-gradient(135deg, #6c757d, #545b62);
            color: white; 
        }
        
        .alert {
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
            border: 1px solid transparent;
        }
        
        .alert-warning {
            background-color: #fff3cd;
            border-color: #ffeaa7;
            color: #856404;
        }
        
        .alert-info {
            background-color: #d1ecf1;
            border-color: #bee5eb;
            color: #0c5460;
        }
        
        .text-center { text-align: center; }
        .text-end { text-align: right; }
        .text-start { text-align: left; }
        .text-success { color: #28a745; }
        .text-warning { color: #ffc107; }
        .text-danger { color: #dc3545; }
        .text-info { color: #17a2b8; }
        .text-muted { color: #6c757d; }
        .text-primary { color: #007bff; }
        
        .row {
            display: flex;
            flex-wrap: wrap;
            margin: 0 -10px;
        }
        
        .col {
            padding: 0 10px;
            box-sizing: border-box;
        }
        
        .col-12 { width: 100%; }
        .col-md-8 { width: 66.66%; }
        .col-md-6 { width: 50%; }
        .col-md-4 { width: 33.33%; }
        .col-md-3 { width: 25%; }
        
        .mb-1 { margin-bottom: 5px; }
        .mb-2 { margin-bottom: 10px; }
        .mb-3 { margin-bottom: 15px; }
        .mb-4 { margin-bottom: 20px; }
        .mb-5 { margin-bottom: 25px; }
        
        .mt-1 { margin-top: 5px; }
        .mt-2 { margin-top: 10px; }
        .mt-3 { margin-top: 15px; }
        .mt-4 { margin-top: 20px; }
        .mt-5 { margin-top: 25px; }
        
        .p-1 { padding: 5px; }
        .p-2 { padding: 10px; }
        .p-3 { padding: 15px; }
        
        .badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: bold;
            text-align: center;
            white-space: nowrap;
            vertical-align: baseline;
        }
        
        .badge-success { background-color: #28a745; color: white; }
        .badge-primary { background-color: #007bff; color: white; }
        .badge-warning { background-color: #ffc107; color: #000; }
        .badge-info { background-color: #17a2b8; color: white; }
        
        @media (max-width: 768px) {
            .container {
                padding: 10px;
            }
            
            .kpi-container {
                flex-direction: column;
            }
            
            .kpi-box {
                width: 100%;
                margin-bottom: 10px;
            }
            
            .col-md-3, .col-md-4, .col-md-6, .col-md-8 {
                width: 100%;
            }
            
            table {
                display: block;
                overflow-x: auto;
                white-space: nowrap;
            }
            
            .btn-print {
                top: 10px;
                right: 10px;
                padding: 6px 12px;
                font-size: 12px;
            }
        }
        
        @media print {
            .kpi-box {
                break-inside: avoid;
                box-shadow: none;
                border: 1px solid #ddd;
            }
            
            .card {
                break-inside: avoid;
                box-shadow: none;
                border: 1px solid #ddd;
            }
        }
    </style>
</head>
<body>
    <!-- Botão para impressão -->
    <div class="btn-print no-print">
        <button class="btn btn-danger" onclick="window.print()">
            🖨️ Imprimir / PDF
        </button>
        <button class="btn btn-secondary" onclick="window.close()">
            ✕ Fechar
        </button>
    </div>
    
    <div class="container">
        <!-- Cabeçalho -->
        <div class="header">
            <h1>🏢 SGVendas - Dossiê do Corretor</h1>
            <h2><%= filtroCorretor %></h2>
            <p class="text-muted">Ano: <strong><%= filtroAno %></strong> | Gerado em: <%= FormatDateTime(Date(), 1) %> às <%= Time() %></p>
            <p class="text-muted mb-0">Código do Relatório: DOSSIE-<%= Year(Date()) %>-<%= filtroAno %>-<%= Replace(Left(filtroCorretor, 10), " ", "-") %></p>
        </div>
        
        <!-- Informações do Corretor -->
        <div class="corretor-info">
            <div class="row">
                <div class="col-md-8">
                    <h3 class="mb-1">👤 <%= filtroCorretor %></h3>
                    <p class="text-muted mb-0">Análise de Desempenho - Ano <%= filtroAno %></p>
                </div>
                <div class="col-md-4 text-end">
                    <p class="mb-0"><strong>Status:</strong> <span class="badge badge-success">Relatório Gerado</span></p>
                    <p class="mb-0 text-muted"><small>Documento para uso interno</small></p>
                </div>
            </div>
        </div>
        
        <!-- KPIs Principais -->
        <div class="kpi-container">
            <div class="kpi-box bg-vendas">
                <h3><%= totalVendas %></h3>
                <p>Unidades Vendidas</p>
            </div>
            <div class="kpi-box bg-vgv">
                <h3>R$ <%= FormatNumber(totalVGV, 2) %></h3>
                <p>VGV Total</p>
            </div>
            <div class="kpi-box bg-comissao">
                <h3>R$ <%= FormatNumber(totalComissao, 2) %></h3>
                <p>Comissão Total</p>
            </div>
            <div class="kpi-box bg-media">
                <h3>R$ <%= FormatNumber(mediaPorVenda, 2) %></h3>
                <p>Média por Venda</p>
            </div>
        </div>
        
        <!-- Estatísticas Adicionais -->
        <div class="row mb-5">
            <div class="col-md-4">
                <div class="card">
                    <div class="card-body text-center">
                        <h5 class="card-title">Percentual de Comissão</h5>
                        <div class="valor-destaque"><%= FormatNumber(percentualComissao, 1) %>%</div>
                        <p class="text-muted mb-0">Média sobre o VGV</p>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card">
                    <div class="card-body text-center">
                        <h5 class="card-title">Meses com Vendas</h5>
                        <div class="valor-destaque"><%= mesesComVendas.Count %></div>
                        <p class="text-muted mb-0">de 12 meses no ano</p>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card">
                    <div class="card-body text-center">
                        <h5 class="card-title">Diversificação</h5>
                        <div class="valor-destaque"><%= empreendimentosList.Count %></div>
                        <p class="text-muted mb-0">empreendimentos diferentes</p>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Seção 1: Desempenho Mensal -->
        <h3 class="section-title">📅 Desempenho Mensal</h3>
        
        <div class="alert alert-info mb-4">
            <strong>Distribuição por Mês:</strong> Clique no botão de impressão para gerar um PDF profissional deste relatório.
        </div>
        
        <div class="mb-4">
            <h5 class="mb-2">Visão Geral dos Meses</h5>
            <%
            For i = 1 To 12
                If mesesComVendas.Exists(CStr(i)) Then
                    Dim dadosMes
                    dadosMes = mesesComVendas(CStr(i))
            %>
            <div class="mes-com-venda" title="<%= dadosMes(0) %> vendas - R$ <%= FormatNumber(dadosMes(1), 2) %>">
                <strong><%= arrMesesNome(i) %></strong><br>
                <small><%= dadosMes(0) %> vendas</small>
            </div>
            <%
                Else
            %>
            <div class="mes-sem-venda" title="Sem vendas neste mês">
                <%= arrMesesNome(i) %>
            </div>
            <%
                End If
            Next
            %>
        </div>
        
        <div class="table-responsive">
            <table>
                <thead>
                    <tr>
                        <th>Mês</th>
                        <th class="text-center">Quantidade de Vendas</th>
                        <th class="text-end">VGV Total (R$)</th>
                        <th class="text-end">Comissão (R$)</th>
                        <th class="text-center">% do VGV Anual</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    For i = 1 To 12
                        If mesesComVendas.Exists(CStr(i)) Then
                            dadosMes = mesesComVendas(CStr(i))
                            Dim percentualMes
                            If totalVGV > 0 Then
                                percentualMes = (dadosMes(1) / totalVGV) * 100
                            Else
                                percentualMes = 0
                            End If
                    %>
                    <tr>
                        <td><strong><%= arrMesesNome(i) %></strong></td>
                        <td class="text-center"><span class="badge badge-primary"><%= dadosMes(0) %></span></td>
                        <td class="text-end"><strong>R$ <%= FormatNumber(dadosMes(1), 2) %></strong></td>
                        <td class="text-end">R$ <%= FormatNumber(dadosMes(2), 2) %></td>
                        <td class="text-center">
                            <div style="background-color: #e9ecef; border-radius: 5px; height: 20px; position: relative;">
                                <div style="background-color: #28a745; border-radius: 5px; height: 100%; width: <%= percentualMes %>%"></div>
                                <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; line-height: 20px; text-align: center; font-size: 11px; font-weight: bold;">
                                    <%= FormatNumber(percentualMes, 1) %>%
                                </div>
                            </div>
                        </td>
                    </tr>
                    <%
                        Else
                    %>
                    <tr>
                        <td><strong><%= arrMesesNome(i) %></strong></td>
                        <td class="text-center"><span class="badge badge-secondary">0</span></td>
                        <td class="text-end">R$ 0,00</td>
                        <td class="text-end">R$ 0,00</td>
                        <td class="text-center">
                            <div style="background-color: #e9ecef; border-radius: 5px; height: 20px; position: relative;">
                                <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; line-height: 20px; text-align: center; font-size: 11px; color: #6c757d;">
                                    0%
                                </div>
                            </div>
                        </td>
                    </tr>
                    <%
                        End If
                    Next
                    %>
                </tbody>
                <tfoot>
                    <tr style="background-color: #800000; color: white; font-weight: bold;">
                        <td><strong>TOTAL ANUAL</strong></td>
                        <td class="text-center"><strong><%= totalVendas %></strong></td>
                        <td class="text-end"><strong>R$ <%= FormatNumber(totalVGV, 2) %></strong></td>
                        <td class="text-end"><strong>R$ <%= FormatNumber(totalComissao, 2) %></strong></td>
                        <td class="text-center"><strong>100%</strong></td>
                    </tr>
                </tfoot>
            </table>
        </div>
        
        <!-- Seção 2: Empreendimentos -->
        <% If HasItems(empreendimentosList) Then %>
        <h3 class="section-title">🏢 Empreendimentos Vendidos</h3>
        
        <div class="table-responsive">
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Nome do Empreendimento</th>
                        <th class="text-center">Quantidade de Vendas</th>
                        <th class="text-end">VGV Total (R$)</th>
                        <th class="text-center">% do VGV Total</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    Dim contadorEmp, empKey
                    contadorEmp = 0
                    
                    For Each empKey In empreendimentosList.Keys
                        contadorEmp = contadorEmp + 1
                        Dim empreendInfo
                        empreendInfo = empreendimentosList(empKey)
                        
                        If contadorEmp <= 20 Then
                    %>
                    <tr>
                        <td><%= contadorEmp %></td>
                        <td><strong><%= empreendInfo(0) %></strong></td>
                        <td class="text-center">
                            <span class="badge badge-success"><%= empreendInfo(1) %></span>
                        </td>
                        <td class="text-end"><strong>R$ <%= FormatNumber(empreendInfo(2), 2) %></strong></td>
                        <td class="text-center">
                            <span class="badge badge-info"><%= FormatNumber(empreendInfo(3), 1) %>%</span>
                        </td>
                    </tr>
                    <%
                        End If
                    Next
                    
                    If contadorEmp > 20 Then
                    %>
                    <tr style="background-color: #f8f9fa;">
                        <td colspan="5" class="text-center">
                            <em>... e mais <%= contadorEmp - 20 %> empreendimentos (total de <%= contadorEmp %> diferentes)</em>
                        </td>
                    </tr>
                    <%
                    End If
                    %>
                </tbody>
                <tfoot>
                    <tr style="background-color: #f5f5f5; font-weight: bold;">
                        <td colspan="2"><strong>TOTAL / MÉDIA</strong></td>
                        <td class="text-center"><strong><%= totalVendas %></strong></td>
                        <td class="text-end"><strong>R$ <%= FormatNumber(totalVGV, 2) %></strong></td>
                        <td class="text-center"><strong>100%</strong></td>
                    </tr>
                </tfoot>
            </table>
        </div>
        <% Else %>
        <h3 class="section-title">🏢 Empreendimentos Vendidos</h3>
        <div class="alert alert-warning">
            <strong>Atenção:</strong> Nenhum empreendimento registrado para este corretor no ano de <%= filtroAno %>.
        </div>
        <% End If %>
        
        <!-- Seção 3: Localidades -->
        <% If HasItems(localidadesList) Then %>
        <h3 class="section-title">📍 Atuação por Localidade</h3>
        
        <div class="table-responsive">
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Localidade</th>
                        <th class="text-center">Vendas Realizadas</th>
                        <th class="text-end">VGV (R$)</th>
                        <th class="text-center">Participação</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    Dim contadorLocal, locKey
                    contadorLocal = 0
                    
                    For Each locKey In localidadesList.Keys
                        contadorLocal = contadorLocal + 1
                        Dim localInfo
                        localInfo = localidadesList(locKey)
                        
                        Dim mediaLocal
                        If localInfo(1) > 0 Then
                            mediaLocal = localInfo(2) / localInfo(1)
                        Else
                            mediaLocal = 0
                        End If
                        
                        If contadorLocal <= 15 Then
                    %>
                    <tr>
                        <td><%= contadorLocal %></td>
                        <td><strong><%= localInfo(0) %></strong></td>
                        <td class="text-center">
                            <span class="badge badge-primary"><%= localInfo(1) %></span>
                        </td>
                        <td class="text-end">
                            <strong>R$ <%= FormatNumber(localInfo(2), 2) %></strong><br>
                            <small class="text-muted">(R$ <%= FormatNumber(mediaLocal, 2) %>/venda)</small>
                        </td>
                        <td class="text-center">
                            <span class="badge badge-warning"><%= FormatNumber(localInfo(3), 1) %>%</span>
                        </td>
                    </tr>
                    <%
                        End If
                    Next
                    
                    If contadorLocal > 15 Then
                    %>
                    <tr style="background-color: #f8f9fa;">
                        <td colspan="5" class="text-center">
                            <em>... e mais <%= contadorLocal - 15 %> localidades (atuou em <%= contadorLocal %> diferentes)</em>
                        </td>
                    </tr>
                    <%
                    End If
                    %>
                </tbody>
                <tfoot>
                    <tr style="background-color: #f5f5f5; font-weight: bold;">
                        <td colspan="2"><strong>TOTAL</strong></td>
                        <td class="text-center"><strong><%= totalVendas %></strong></td>
                        <td class="text-end"><strong>R$ <%= FormatNumber(totalVGV, 2) %></strong></td>
                        <td class="text-center"><strong>100%</strong></td>
                    </tr>
                </tfoot>
            </table>
        </div>
        <% Else %>
        <h3 class="section-title">📍 Atuação por Localidade</h3>
        <div class="alert alert-warning">
            <strong>Atenção:</strong> Nenhuma localidade registrada para este corretor no ano de <%= filtroAno %>.
        </div>
        <% End If %>
        
        <!-- Seção 4: Top Vendas -->
        <% If HasItems(topVendasList) Then %>
        <h3 class="section-title">⭐ Maiores Vendas do Ano</h3>
        
        <div class="table-responsive">
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Data da Venda</th>
                        <th>Empreendimento</th>
                        <th class="text-end">Valor da Venda (R$)</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    Dim contadorTopVenda, topKey
                    contadorTopVenda = 0
                    Dim totalTop5
                    totalTop5 = 0
                    
                    For Each topKey In topVendasList.Keys
                        contadorTopVenda = contadorTopVenda + 1
                        Dim topVenda
                        topVenda = topVendasList(topKey)
                        totalTop5 = totalTop5 + topVenda(2)
                    %>
                    <tr>
                        <td><%= contadorTopVenda %></td>
                        <td><%= topVenda(0) %></td>
                        <td><%= topVenda(1) %></td>
                        <td class="text-end">
                            <strong style="color: #28a745; font-size: 14px;">
                                R$ <%= FormatNumber(topVenda(2), 2) %>
                            </strong>
                        </td>
                    </tr>
                    <%
                    Next
                    %>
                </tbody>
                <tfoot>
                    <tr style="background-color: #f5f5f5; font-weight: bold;">
                        <td colspan="3" class="text-end"><strong>Valor Total das 5 Maiores Vendas:</strong></td>
                        <td class="text-end">
                            <strong style="color: #dc3545; font-size: 15px;">
                                R$ <%= FormatNumber(totalTop5, 2) %>
                            </strong>
                        </td>
                    </tr>
                </tfoot>
            </table>
        </div>
        <% End If %>
        
        <!-- Seção 5: Resumo Executivo -->
        <h3 class="section-title">📊 Análise de Desempenho</h3>
        
        <div class="card">
            <div class="card-header">
                Resumo Executivo
            </div>
            <div class="card-body">
                <div class="row">
                    <div class="col-md-6">
                        <h5 class="text-success mb-3">✅ Pontos Fortes</h5>
                        <ul class="mb-4">
                            <%
                            On Error Resume Next
                            
                            If totalVendas >= 10 Then
                                Response.Write "<li class='mb-2'><strong>Alto volume de vendas:</strong> " & totalVendas & " unidades comercializadas.</li>"
                            ElseIf totalVendas >= 5 Then
                                Response.Write "<li class='mb-2'><strong>Volume satisfatório:</strong> " & totalVendas & " unidades vendidas.</li>"
                            Else
                                Response.Write "<li class='mb-2'><strong>Volume básico:</strong> " & totalVendas & " unidades comercializadas.</li>"
                            End If
                            
                            If mediaPorVenda > 500000 Then
                                Response.Write "<li class='mb-2'><strong>Alto ticket médio:</strong> R$ " & FormatNumber(mediaPorVenda, 2) & " por unidade.</li>"
                            ElseIf mediaPorVenda > 300000 Then
                                Response.Write "<li class='mb-2'><strong>Ticket médio satisfatório:</strong> R$ " & FormatNumber(mediaPorVenda, 2) & " por unidade.</li>"
                            End If
                            
                            If mesesComVendas.Count >= 8 Then
                                Response.Write "<li class='mb-2'><strong>Boa consistência mensal:</strong> Vendas em " & mesesComVendas.Count & " dos 12 meses.</li>"
                            ElseIf mesesComVendas.Count >= 6 Then
                                Response.Write "<li class='mb-2'><strong>Regularidade moderada:</strong> Vendas em " & mesesComVendas.Count & " meses.</li>"
                            End If
                            
                            If empreendimentosList.Count >= 3 Then
                                Response.Write "<li class='mb-2'><strong>Diversificação de portfólio:</strong> Atuou em " & empreendimentosList.Count & " empreendimentos diferentes.</li>"
                            End If
                            
                            If localidadesList.Count >= 2 Then
                                Response.Write "<li class='mb-2'><strong>Ampla atuação geográfica:</strong> Presença em " & localidadesList.Count & " localidades.</li>"
                            End If
                            
                            If Err.Number <> 0 Then
                                Err.Clear
                            End If
                            On Error GoTo 0
                            %>
                        </ul>
                    </div>
                    
                    <div class="col-md-6">
                        <h5 class="text-warning mb-3">📈 Oportunidades de Melhoria</h5>
                        <ul>
                            <%
                            On Error Resume Next
                            
                            If mesesComVendas.Count < 6 Then
                                Response.Write "<li class='mb-2'><strong>Aumentar regularidade:</strong> Apenas " & mesesComVendas.Count & " meses com vendas.</li>"
                            End If
                            
                            If (12 - mesesComVendas.Count) > 6 Then
                                Response.Write "<li class='mb-2'><strong>Reduzir meses sem vendas:</strong> " & (12 - mesesComVendas.Count) & " meses sem comercialização.</li>"
                            End If
                            
                            If empreendimentosList.Count < 2 Then
                                Response.Write "<li class='mb-2'><strong>Diversificar empreendimentos:</strong> Atuação concentrada em poucos projetos.</li>"
                            End If
                            
                            If percentualComissao < 2.5 Then
                                Response.Write "<li class='mb-2'><strong>Melhorar margem de comissão:</strong> Atualmente em " & FormatNumber(percentualComissao, 1) & "% sobre VGV.</li>"
                            End If
                            
                            If totalVendas > 0 Then
                                Dim taxaConversaoEstimada
                                taxaConversaoEstimada = totalVendas * 10 ' Estimativa: cada venda = 10 visitas/prospecções
                                Response.Write "<li class='mb-2'><strong>Otimizar taxa de conversão:</strong> Estimativa de " & taxaConversaoEstimada & " prospecções realizadas.</li>"
                            End If
                            
                            If Err.Number <> 0 Then
                                Err.Clear
                            End If
                            On Error GoTo 0
                            %>
                        </ul>
                    </div>
                </div>
                
                <div class="mt-4 pt-3 border-top">
                    <h5 class="text-primary mb-3">🎯 Recomendações Estratégicas</h5>
                    <ol>
                        <li class="mb-2"><strong>Foco em Empreendimentos Premium:</strong> Considerando o ticket médio atual de R$ <%= FormatNumber(mediaPorVenda, 2) %>, há espaço para aumentar o foco em empreendimentos de alto valor agregado.</li>
                        
                        <li class="mb-2"><strong>Expansão Geográfica:</strong> 
                            <%
                            On Error Resume Next
                            If localidadesList.Count < 3 Then
                                Response.Write "Explorar novas localidades para reduzir dependência regional e aumentar a base de clientes."
                            Else
                                Response.Write "Manter a diversificação geográfica já estabelecida, explorando oportunidades nas localidades de melhor performance."
                            End If
                            On Error GoTo 0
                            %>
                        </li>
                        
                        <li class="mb-2"><strong>Consistência Operacional:</strong> Trabalhar para aumentar a regularidade mensal, visando atingir vendas em pelo menos 8 dos 12 meses do ano.</li>
                        
                        <li class="mb-2"><strong>Valorização do Portfólio:</strong> 
                            <%
                            If percentualComissao < 3 Then
                                Response.Write "Buscar negociação de percentuais de comissão mais atrativos, especialmente para empreendimentos de alto valor."
                            Else
                                Response.Write "Manter os bons percentuais de comissão já conquistados, focando em aumentar o volume de negócios."
                            End If
                            %>
                        </li>
                        
                        <li class="mb-2"><strong>Desenvolvimento de Habilidades:</strong> Investir em treinamento em técnicas de vendas complexas e negociação de alto valor para melhorar o desempenho em empreendimentos premium.</li>
                    </ol>
                </div>
            </div>
        </div>
        
        <!-- Rodapé -->
        <div class="footer">
            <p><strong>🏢 SGVendas - Sistema de Gerenciamento de Vendas</strong></p>
            <p class="mb-1">Relatório técnico gerado automaticamente pelo sistema. Dados referentes ao ano de <%= filtroAno %>.</p>
            <p class="mb-1 text-muted">Código do Documento: DOSSIE-<%= Year(Date()) %>-<%= filtroAno %>-<%= Replace(Left(filtroCorretor, 10), " ", "-") %></p>
            <p class="mb-0"><strong>Documento confidencial</strong> - Destinado exclusivamente para uso interno e análise gerencial.</p>
            <p class="mt-2 text-muted"><small>Versão do Relatório: 1.0 | Última atualização: <%= Now() %></small></p>
        </div>
    </div>
    
    <script>
    // Configurações para impressão/PDF
    window.onbeforeprint = function() {
        // Adicionar informações extras para impressão
        var header = '<div style="text-align:center;margin-bottom:20px;border-bottom:3px solid #800000;padding-bottom:15px;page-break-before:always;">' +
                    '<h1 style="color:#800000;font-size:20px;margin:10px 0 5px 0;font-weight:bold;">DOSSIE COMPLETO DO CORRETOR</h1>' +
                    '<h2 style="font-size:16px;margin:5px 0;color:#333;"><%= filtroCorretor %></h2>' +
                    '<p style="margin:5px 0;font-size:13px;color:#666;">Ano: <%= filtroAno %> | Impresso em: ' + new Date().toLocaleDateString('pt-BR') + ' às ' + new Date().toLocaleTimeString('pt-BR', {hour: '2-digit', minute:'2-digit'}) + '</p>' +
                    '</div>';
        
        // Adicionar antes do container
        var container = document.querySelector('.container');
        if (container) {
            container.insertAdjacentHTML('afterbegin', header);
        }
    };
    
    window.onafterprint = function() {
        // Remover o header adicionado
        var header = document.querySelector('div[style*="text-align:center"]');
        if (header && header.innerHTML.includes('DOSSIE COMPLETO')) {
            header.remove();
        }
    };
    
    // Melhorar a experiência de impressão
    document.addEventListener('DOMContentLoaded', function() {
        // Adicionar quebras de página lógicas
        var sectionTitles = document.querySelectorAll('.section-title');
        sectionTitles.forEach(function(title, index) {
            if (index > 0) {
                title.style.pageBreakBefore = 'always';
            }
        });
        
        // Garantir que tabelas não quebrem no meio
        var tables = document.querySelectorAll('table');
        tables.forEach(function(table) {
            table.style.pageBreakInside = 'avoid';
        });
    });
    
    // Focar na janela para facilitar a impressão
    window.focus();
    </script>
</body>
</html>