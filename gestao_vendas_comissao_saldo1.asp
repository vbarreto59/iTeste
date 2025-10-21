<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<%
' =======================================================
' Configuração da conexão com o banco de dados e consulta
' =======================================================

' Declaração das variáveis para a tabela
Dim conn, rs, sql

' Cria o objeto de conexão ADODB
Set conn = Server.CreateObject("ADODB.Connection")

' Abre a conexão usando a string de conexão do arquivo 'conSunSales.asp'
conn.Open StrConnSales

' Constrói a consulta SQL para totalizar as colunas por Nome, incluindo o Saldo
' O nome do campo "TotalPago" foi alterado para "TotalComissaoPago"
sql = "SELECT Nome, " & _
      "Sum(TotalVenda) AS SomaDeTotalVenda, " & _
      "Sum(TotalComissao) AS SomaDeTotalComissao, " & _
      "Sum(TotalComissaoPago) AS SomaDeTotalComissaoPago, " & _
      "Sum(TotalComissao) - Sum(TotalComissaoPago) AS Saldo " & _
      "FROM ComissaoSaldo GROUP BY Nome ORDER BY Nome;"

' Cria o Recordset e executa a consulta
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn

' =======================================================
' Fim das consultas
' =======================================================
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Saldo Comissões</title>
    
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap5.min.css">
</head>
<body>

<div class="container mt-5">
    <h1 class="mb-4">Saldo Comissões a Pagar</h1>

    <table id="myTable" class="table table-striped table-bordered" style="width:100%">
        <thead>
            <tr>
                <th>Nome</th>
                <th>Total Venda</th>
                <th>Total Comissao</th>
                <th>Total Pago</th>
                <th>Saldo</th>
            </tr>
        </thead>
        <tbody>
            <% 
            If Not rs.EOF Then
                Do While Not rs.EOF
            %>
            <tr>
                <td><%= rs("Nome") %></td>
                <td class="text-end">
                    <%
                    If Not IsNull(rs("SomaDeTotalVenda")) And rs("SomaDeTotalVenda") <> 0 Then
                        Response.Write "R$ " & FormatNumber(rs("SomaDeTotalVenda"), 2)
                    Else
                        Response.Write "R$ 0,00"
                    End If
                    %>
                </td>
                <td class="text-end">
                    <%
                    If Not IsNull(rs("SomaDeTotalComissao")) And rs("SomaDeTotalComissao") <> 0 Then
                        Response.Write "R$ " & FormatNumber(rs("SomaDeTotalComissao"), 2)
                    Else
                        Response.Write "R$ 0,00"
                    End If
                    %>
                </td>
                <td class="text-end">
                    <%
                    ' A referência ao campo foi alterada para "SomaDeTotalComissaoPago"
                    If Not IsNull(rs("SomaDeTotalComissaoPago")) And rs("SomaDeTotalComissaoPago") <> 0 Then
                        Response.Write "R$ " & FormatNumber(rs("SomaDeTotalComissaoPago"), 2)
                    Else
                        Response.Write "R$ 0,00"
                    End If
                    %>
                </td>
                <td class="text-end">
                    <%
                    Dim vSaldo
                    vSaldo = 0
                    vComisPago = 0
                     If Not IsNull(rs("SomaDeTotalComissaoPago")) And rs("SomaDeTotalComissaoPago") <> 0 Then
                        vComisPago = rs("SomaDeTotalComissaoPago")
                     end if   
                     vComisTotal = 0
                     If Not IsNull(rs("SomaDeTotalComissao")) And rs("SomaDeTotalComissao") <> 0 Then
                        vComisTotal = rs("SomaDeTotalComissao")
                     end if   
                    
                    vSaldo = vComisTotal-vComisPago

                    If vSaldo < 0 Then%>
                           <span style="background-color: #B9C7E7; display: block;">
                           <%Response.Write "R$ " & FormatNumber(vSaldo, 2)
                             Response.Write "<br><small>Não há venda relacionada.</small></br>"
                    else
                            If Not IsNull(vSaldo) And vSaldo <> 0 Then%>
                                <span style="background-color: #ffcdd2; display: block;">
                                <%Response.Write "R$ " & FormatNumber(vSaldo, 2)
                            Else
                                    Response.Write "R$ 0,00"
                            End If
                    End If
                    %>
                </td>
            </tr>
            <% 
                    rs.MoveNext
                Loop
            End If
            
            ' Fecha o Recordset
            If Not rs Is Nothing Then
                If rs.State = 1 Then rs.Close
            End If
            Set rs = Nothing
            
            ' Fecha a conexão com o banco de dados
            If Not conn Is Nothing Then
                If conn.State = 1 Then conn.Close
            End If
            Set conn = Nothing
            %>
        </tbody>
    </table>
</div>

<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/1.11.5/js/dataTables.bootstrap5.min.js"></script>

<script>
$(document).ready(function() {
    // Inicialização da tabela
    $('#myTable').DataTable({
        "order": [[ 0, "asc" ]], 
        "pageLength": 100,
        "language": {
            "url": "https://cdn.datatables.net/plug-ins/1.11.5/i18n/pt-BR.json"
        }
    });
});
</script>

</body>
</html>