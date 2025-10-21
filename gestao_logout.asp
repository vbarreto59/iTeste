<%
' Limpar a sessão do usuário
Session.Abandon()

' Redirecionar para a página de login
Response.Redirect("gestao_login.asp")
%>