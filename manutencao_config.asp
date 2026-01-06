<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: LOUHKFOAFC          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<%
' Função para verificar se está em manutenção
Function EstaEmManutencao()
    Dim fs, arquivoManutencao
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    arquivoManutencao = Server.MapPath("manut.txt")
    
    EstaEmManutencao = fs.FileExists(arquivoManutencao)
    
    Set fs = Nothing
End Function

' Função para verificar se é o usuário admin
Function EAdmin()
    EAdmin = (LCase(Session("Usuario")) = "barreto")
End Function

' Redirecionar para manutenção se necessário
If EstaEmManutencao() And Not EAdmin() Then
    If Request.ServerVariables("SCRIPT_NAME") <> "manutencao.asp" Then
        Response.Redirect "manutencao.asp"
        Response.End
    End If
End If
%>


<%
' Função para verificar se o diretório de vendas está acessível
Function DiretorioVendasAcessivel()
    Dim fs
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    DiretorioVendasAcessivel = fs.FileExists(Server.MapPath("DIRETORIAS_VENDAS.TXT"))
    Set fs = Nothing
End Function

' Exemplo de uso em outras páginas:
If Not DiretorioVendasAcessivel() Then
    'Response.Write "O diretório de vendas está temporariamente indisponível para manutenção."
    'Response.End
End If
%>

<%
' Função para verificar se o diretório de vendas está BLOQUEADO
Function DiretorioVendasBloqueado()
    Dim fs
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    DiretorioVendasBloqueado = fs.FileExists(Server.MapPath("DIRETORIAS_VENDAS.TXT"))
    Set fs = Nothing
End Function

' Exemplo de uso - Bloquear acesso se o arquivo existir
If DiretorioVendasBloqueado() Then
    'Response.Write "<div style='padding:20px; background:#f8d7da; border:1px solid #f5c6cb; color:#721c24; margin:20px; border-radius:5px;'>"
    'Response.Write "<h3><i class='fas fa-exclamation-triangle'></i> Acesso Restrito</h3>"
    'Response.Write "<p>O diretório de vendas está temporariamente bloqueado para manutenção.</p>"
    'Response.Write "<p><small>Por favor, tente novamente mais tarde.</small></p>"
    'Response.Write "</div>"
    'Response.End
End If
%>