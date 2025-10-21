<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conSunSales.asp"-->

<%
' ====================================================================
' Script para Excluir um Registro de Comissão
' ====================================================================

Dim connSales
Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

Dim comissaoId
comissaoId = Request.QueryString("id")

' Verifica se o ID foi passado e é um número válido
If Not IsNumeric(comissaoId) Or comissaoId = "" Then
    Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=ID de comissão inválido para exclusão."
    Response.End
End If

' Inicia uma transação para garantir que ambas as exclusões (se houver) sejam atômicas
connSales.BeginTrans

On Error Resume Next ' Habilita tratamento de erro para operações de banco de dados

' ====================================================================
' 1. Obter o ID_Venda da comissão a ser excluída
' Isso é necessário para excluir os pagamentos relacionados na tabela PAGAMENTOS_COMISSOES
' ====================================================================
Dim rsVendaId
Set rsVendaId = Server.CreateObject("ADODB.Recordset")
rsVendaId.Open "SELECT ID_Venda FROM COMISSOES_A_PAGAR WHERE ID_Comissoes = " & CInt(comissaoId), connSales

Dim vendaIdToDelete
If Not rsVendaId.EOF Then
    vendaIdToDelete = rsVendaId("ID_Venda")
Else
    ' Se a comissão não for encontrada, redireciona com erro
    Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Comissão não encontrada para exclusão."
    rsVendaId.Close
    connSales.RollbackTrans ' Reverte a transação se algo der errado
    Response.End
End If
rsVendaId.Close

' ====================================================================
' 2. Excluir registros relacionados na tabela PAGAMENTOS_COMISSOES
' Isso evita erros de integridade referencial.
' ====================================================================
Dim sqlDeletePagamentos
sqlDeletePagamentos = "DELETE FROM PAGAMENTOS_COMISSOES WHERE ID_Venda = " & CInt(vendaIdToDelete)
connSales.Execute sqlDeletePagamentos

If Err.Number <> 0 Then
    ' Se houver erro na exclusão de pagamentos, reverte a transação
    connSales.RollbackTrans
    Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Erro ao excluir pagamentos relacionados: " & Replace(Err.Description, "'", "\'")
    Response.End
End If

' ====================================================================
' 3. Excluir o registro principal da tabela COMISSOES_A_PAGAR
' ====================================================================
Dim sqlDeleteComissao
sqlDeleteComissao = "DELETE FROM COMISSOES_A_PAGAR WHERE ID_Comissoes = " & CInt(comissaoId)
connSales.Execute sqlDeleteComissao

If Err.Number <> 0 Then
    ' Se houver erro na exclusão da comissão, reverte a transação
    connSales.RollbackTrans
    Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Erro ao excluir comissão: " & Replace(Err.Description, "'", "\'")
    Response.End
Else
    ' Se tudo correr bem, confirma a transação
    connSales.CommitTrans
    Response.Redirect "gestao_vendas_gerenc_comissoes.asp?mensagem=Comissão excluída com sucesso!"
End If

On Error GoTo 0 ' Desabilita tratamento de erro

' ====================================================================
' Fecha a conexão
' ====================================================================
If IsObject(connSales) Then
    connSales.Close
    Set connSales = Nothing
End If
%>