<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 18/12/2025               -->
<!-- CODIGO_ARQUIVO: HUKESINDEK          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="conexao.asp"-->           
<!--#include file="conSunSales.asp"-->       
<!--#include file="usr_acoes_v4GVendas.inc"-->

<%
' =============================================
'      META – GERENCIAMENTO (BLOCO ANTES DO HTML)
' =============================================
Response.CodePage = 65001
Response.Charset = "utf-8"

Const ANO_MINIMO = 2026

' -----------------------------
'  ABRIR CONEXÕES
' -----------------------------
Dim connOrg, connSales
Set connOrg = Server.CreateObject("ADODB.Connection")
Set connSales = Server.CreateObject("ADODB.Connection")

connOrg.Open strConn
connSales.Open strConnSales

Function AtualizarMetaDiretoria()

    Dim sql, rs, sqlInsert, usuarioLogado
    Dim DiretoriaId, Ano, Mes, TotalMetas
    
    usuarioLogado = Session("Usuario")
    If usuarioLogado = "" Then usuarioLogado = "Sistema"
    
    On Error Resume Next
    
    ' Limpar tabela
    connSales.Execute "DELETE FROM MetaDiretoria"
    If Err.Number <> 0 Then Exit Function
    
    ' Buscar soma das metas agrupadas por diretoria/ano/mês
    'sql = "SELECT g.DiretoriaID, mg.Ano, mg.Mes, SUM(mg.ValorMeta) AS Total " & _
       ''   "FROM MetaGerencia mg " & _
      ''    "INNER JOIN Gerencias g ON mg.GerenciaID = g.GerenciaID " & _
       ''   "WHERE mg.ValorMeta > 0 " & _
      ''    "GROUP BY g.DiretoriaID, mg.Ano, mg.Mes"

    sql = "SELECT MetaGerencia.DiretoriaId, MetaGerencia.Ano, MetaGerencia.Mes, Sum(MetaGerencia.ValorMeta) AS SomaDeValorMeta FROM MetaGerencia GROUP BY MetaGerencia.DiretoriaId, MetaGerencia.Ano, MetaGerencia.Mes HAVING (((Sum(MetaGerencia.ValorMeta))>0));"      
    
    Set rs = connSales.Execute(sql)
    
    Do Until rs.EOF
        DiretoriaId = rs("DiretoriaID")
        Ano = rs("Ano")
        Mes = rs("Mes")
        TotalMetas = rs("SomaDeValorMeta")
        
        
        sqlInsert = "INSERT INTO MetaDiretoria " & _
                        "(DiretoriaID, Ano, Mes, TotalMetas, Usuario) " & _
                        "VALUES (" & DiretoriaId & "," & Ano & "," & Mes & "," & _
                        TotalMetas & "," & "'" & Replace(usuarioLogado, "'", "''") & "')"
        'Response.write sqlInsert
        'Response.end                 
        
        connSales.Execute sqlInsert
    
        
        rs.MoveNext
    Loop
    
    ' Atualizar meta da empresa
    Call AtualizarMetaEmpresa()
    
    On Error GoTo 0
    
End Function


Function AtualizarMetaEmpresa()

    Dim sql, rs, sqlInsert
    Dim Ano, Mes, TotalEmpresa
    
    On Error Resume Next
    
    ' Limpar tabela (apenas anos a partir de 2026)
    connSales.Execute "DELETE FROM MetaEmpresa WHERE Ano >= 2026"
    If Err.Number <> 0 Then Exit Function
    
    ' Buscar soma das metas da empresa agrupadas por ano/mês
    sql = "SELECT " & _
          "MetaDiretoria.Ano, " & _
          "MetaDiretoria.Mes, " & _
          "Sum(MetaDiretoria.TotalMetas) AS SomaDeTotalMetas " & _
          "FROM MetaDiretoria " & _
          "WHERE MetaDiretoria.Ano >= 2026 " & _
          "GROUP BY MetaDiretoria.Ano, MetaDiretoria.Mes " & _
          "HAVING Sum(MetaDiretoria.TotalMetas) > 0"
          'response.write sql 
          'response.end  
    
    Set rs = connSales.Execute(sql)
    
    Do Until rs.EOF
        Ano = rs("Ano")
        Mes = rs("Mes")
        TotalEmpresa = rs("SomaDeTotalMetas")
        
        ' Verificar se o valor não é nulo
        If Not IsNull(TotalEmpresa) And TotalEmpresa > 0 Then
            sqlInsert = "INSERT INTO MetaEmpresa " & _
                        "(Ano, Mes, Meta) " & _
                        "VALUES (" & Ano & "," & Mes & "," & _
                        TotalEmpresa & ")"
          response.write sqlInsert 
          'response.end                         
            
            connSales.Execute sqlInsert
        End If
        
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    On Error GoTo 0
    
End Function

AtualizarMetaDiretoria()
%>

