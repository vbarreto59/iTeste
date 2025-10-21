
<%if Trim(StrConn)="" then%>
     <!--#include file="conexao.asp"-->
<%end if%>     
<%if Trim(StrConnSales)="" then%>
     <!--#include file="conSunSales.asp"-->
<%end if%>  

<%
' Configuração
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, must-revalidate"
Response.ContentType = "text/html"
Response.Charset = "UTF-8"

On Error Resume Next

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set cmd = Server.CreateObject("ADODB.Command")

' Abrir conexão
conn.Open StrConnSales

If Err.Number <> 0 Then
    Response.Write "Erro ao conectar ao banco de dados: " & Err.Description
    Response.End
End If

' Iniciar transação para garantir integridade dos dados
conn.BeginTrans

' 1. Excluir registros de Vendas onde Obs = 'Massa 2025 - auto'
Dim sql, affectedRows
sql = "DELETE FROM Vendas WHERE Obs = 'Massa 2025 - auto' AND Excluido = 0"

Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandText = sql
cmd.Execute affectedRows

If Err.Number <> 0 Then
    conn.RollbackTrans
    Response.Write "Erro ao excluir registros de Vendas: " & Err.Description
    Response.End
End If

' 2. Verificar registros órfãos em COMISSOES_A_PAGAR
Dim comissoesOrfas, pagamentosOrfos
sql = "SELECT CP.* FROM COMISSOES_A_PAGAR CP " & _
      "LEFT JOIN Vendas V ON CP.ID_VENDA = V.ID " & _
      "WHERE V.ID IS NULL"

Set rs = conn.Execute(sql)
comissoesOrfas = 0

If Not rs.EOF Then
    Response.Write "<h2>Registros Órfãos em COMISSOES_A_PAGAR</h2>"
    Response.Write "<table border='1' style='border-collapse: collapse; width: 100%;'>"
    Response.Write "<tr><th>ID</th><th>ID_VENDA</th><th>Valor</th><th>Status</th></tr>"
    
    Do While Not rs.EOF
        comissoesOrfas = comissoesOrfas + 1
        Response.Write "<tr>"
        Response.Write "<td>" & rs("ID") & "</td>"
        Response.Write "<td>" & rs("ID_VENDA") & "</td>"
        Response.Write "<td>R$ " & FormatNumber(rs("Valor"), 2) & "</td>"
        Response.Write "<td>" & rs("Status") & "</td>"
        Response.Write "</tr>"
        rs.MoveNext
    Loop
    Response.Write "</table>"
Else
    Response.Write "<p>Nenhum registro órfão encontrado em COMISSOES_A_PAGAR</p>"
End If
rs.Close

' 3. Verificar registros órfãos em PAGAMENTOS_COMISSOES
sql = "SELECT PC.* FROM PAGAMENTOS_COMISSOES PC " & _
      "LEFT JOIN Vendas V ON PC.ID_VENDA = V.ID " & _
      "WHERE V.ID IS NULL"

Set rs = conn.Execute(sql)
pagamentosOrfos = 0

If Not rs.EOF Then
    Response.Write "<h2>Registros Órfãos em PAGAMENTOS_COMISSOES</h2>"
    Response.Write "<table border='1' style='border-collapse: collapse; width: 100%;'>"
    Response.Write "<tr><th>ID</th><th>ID_VENDA</th><th>Valor Pago</th><th>Data Pagamento</th></tr>"
    
    Do While Not rs.EOF
        pagamentosOrfos = pagamentosOrfos + 1
        Response.Write "<tr>"
        Response.Write "<td>" & rs("ID") & "</td>"
        Response.Write "<td>" & rs("ID_VENDA") & "</td>"
        Response.Write "<td>R$ " & FormatNumber(rs("ValorPago"), 2) & "</td>"
        Response.Write "<td>" & FormatDateTime(rs("DataPagamento"), 2) & "</td>"
        Response.Write "</tr>"
        rs.MoveNext
    Loop
    Response.Write "</table>"
Else
    Response.Write "<p>Nenhum registro órfão encontrado em PAGAMENTOS_COMISSOES</p>"
End If
rs.Close

' 4. Opção para excluir os registros órfãos (com confirmação)
If (comissoesOrfas > 0 Or pagamentosOrfos > 0) And Request.Form("confirmar") = "1" Then
    ' Excluir registros órfãos de COMISSOES_A_PAGAR
    If comissoesOrfas > 0 Then
        sql = "DELETE FROM COMISSOES_A_PAGAR WHERE ID IN (" & _
              "SELECT CP.ID FROM COMISSOES_A_PAGAR CP " & _
              "LEFT JOIN Vendas V ON CP.ID_VENDA = V.ID " & _
              "WHERE V.ID IS NULL)"
        conn.Execute sql
        Response.Write "<p>" & comissoesOrfas & " registros órfãos excluídos de COMISSOES_A_PAGAR</p>"
    End If
    
    ' Excluir registros órfãos de PAGAMENTOS_COMISSOES
    If pagamentosOrfos > 0 Then
        sql = "DELETE FROM PAGAMENTOS_COMISSOES WHERE ID IN (" & _
              "SELECT PC.ID FROM PAGAMENTOS_COMISSOES PC " & _
              "LEFT JOIN Vendas V ON PC.ID_VENDA = V.ID " & _
              "WHERE V.ID IS NULL)"
        conn.Execute sql
        Response.Write "<p>" & pagamentosOrfos & " registros órfãos excluídos de PAGAMENTOS_COMISSOES</p>"
    End If
End If

' Commit da transação
conn.CommitTrans

' Resumo da execução
Response.Write "<h2>Resumo da Execução</h2>"
Response.Write "<p><strong>Registros excluídos de Vendas:</strong> " & affectedRows & "</p>"
Response.Write "<p><strong>Registros órfãos em COMISSOES_A_PAGAR:</strong> " & comissoesOrfas & "</p>"
Response.Write "<p><strong>Registros órfãos em PAGAMENTOS_COMISSOES:</strong> " & pagamentosOrfos & "</p>"

' Botão para confirmar exclusão dos órfãos (se houver)
If (comissoesOrfas > 0 Or pagamentosOrfos > 0) And Request.Form("confirmar") <> "1" Then
    Response.Write "<form method='post' style='margin-top: 20px; padding: 15px; background: #fff3cd; border: 1px solid #ffeaa7;'>"
    Response.Write "<h3>⚠️ Atenção!</h3>"
    Response.Write "<p>Foram encontrados registros órfãos. Deseja excluí-los?</p>"
    Response.Write "<input type='hidden' name='confirmar' value='1'>"
    Response.Write "<input type='submit' value='Confirmar Exclusão dos Registros Órfãos' style='background: #dc3545; color: white; padding: 10px; border: none; cursor: pointer;'>"
    Response.Write "</form>"
End If

' Fechar conexões
If rs.State = 1 Then rs.Close
If conn.State = 1 Then conn.Close
Set cmd = Nothing
Set rs = Nothing
Set conn = Nothing

If Err.Number <> 0 Then
    Response.Write "<p style='color: red;'>Ocorreu um erro: " & Err.Description & "</p>"
End If
%>

<!DOCTYPE html>
<html>
<head>
    <title>Limpeza de Dados - Massa 2025</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f8f9fa; }
        h1 { color: #343a40; }
        h2 { color: #495057; margin-top: 20px; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; background: white; }
        th, td { border: 1px solid #dee2e6; padding: 8px; text-align: left; }
        th { background-color: #e9ecef; }
        form { margin: 15px 0; }
        input[type="submit"] { background: #007bff; color: white; padding: 10px 15px; border: none; cursor: pointer; }
        input[type="submit"]:hover { background: #0056b3; }
    </style>
</head>
<body>
    <h1>Rotina de Limpeza - Massa 2025</h1>
    <p><strong>Data da execução:</strong> <%=Now()%></p>
    
    <%
    ' Botão para executar novamente
    Response.Write "<form method='post'>"
    Response.Write "<input type='submit' value='Executar Verificação Novamente'>"
    Response.Write "</form>"
    %>
</body>
</html>