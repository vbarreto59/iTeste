<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: AGFESXXFZJ          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<%
' Limpar a sessão do usuário
Session.Abandon()

' Redirecionar para a página de login
Response.Redirect("gestao_login.asp")
%>