<%

Function UserFuncao()
  Dim strAcoesPermitidas
  strAcoesPermitidas = "" ' Inicializa a variável

  ' Verifica se o UserId na sessão existe e é um número válido
  If Session("UserId") <> "" And IsNumeric(Session("UserId")) Then
    Dim objConn, objRS, strSQL, intUserId

    ' Recupera o UserId da sessão
    intUserId = Session("UserId")

    ' Cria e abre a conexão com o banco de dados
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open StrConn ' Certifique-se que 'StrConn' está definida com a sua string de conexão

    ' Define a instrução SQL para buscar as ações permitidas do usuário
    strSQL = "SELECT * FROM Usuarios WHERE UserId = " & intUserId

    ' Cria e executa o Recordset
    Set objRS = Server.CreateObject("ADODB.Recordset")
    objRS.Open strSQL, objConn

    ' Verifica se algum registro foi encontrado
    If Not objRS.EOF Then
      strAcoesPermitidas = objRS("Funcao")
    End If

    ' Fecha e libera os objetos ADO
    objRS.Close
    Set objRS = Nothing
    objConn.Close
    Set objConn = Nothing
  Else
    ' Caso o UserId na sessão seja inválido ou não exista, retorna uma string vazia ou um valor indicativo de erro
    strAcoesPermitidas = "" ' Ou poderia ser algo como "ERRO: Usuário não logado"
  End If

  ' Retorna a string de ações permitidas
  UserFuncao = strAcoesPermitidas
End Function


Function AcoesPermitidas()
  Dim strAcoesPermitidas
  strAcoesPermitidas = "" ' Inicializa a variável

  ' Verifica se o UserId na sessão existe e é um número válido
  If Session("UserId") <> "" And IsNumeric(Session("UserId")) Then
    Dim objConn, objRS, strSQL, intUserId

    ' Recupera o UserId da sessão
    intUserId = Session("UserId")

    ' Cria e abre a conexão com o banco de dados
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open StrConn ' Certifique-se que 'StrConn' está definida com a sua string de conexão

    ' Define a instrução SQL para buscar as ações permitidas do usuário
    strSQL = "SELECT * FROM Usuarios WHERE UserId = " & intUserId

    ' Cria e executa o Recordset
    Set objRS = Server.CreateObject("ADODB.Recordset")
    objRS.Open strSQL, objConn

    ' Verifica se algum registro foi encontrado
    If Not objRS.EOF Then
      ' Atribui o valor do campo 'ACOES' à variável de retorno
      strAcoesPermitidas = objRS("Acoes")
    End If

    ' Fecha e libera os objetos ADO
    objRS.Close
    Set objRS = Nothing
    objConn.Close
    Set objConn = Nothing
  Else
    ' Caso o UserId na sessão seja inválido ou não exista, retorna uma string vazia ou um valor indicativo de erro
    strAcoesPermitidas = "" ' Ou poderia ser algo como "ERRO: Usuário não logado"
  End If

  ' Retorna a string de ações permitidas
  AcoesPermitidas = strAcoesPermitidas
  Session("Funcao")
End Function
%>

<%
Function UsuarioPodeEditar()
  Dim podeEditar

  ' Inicializa a variável como falso
  podeEditar = False

'response.write("Acoes: " & Session("Acoes"))
'response.end()

  ' Verifica se a variável de sessão "Acoes" existe e não está vazia
  If Session("Acoes") <> "" Then
          ' Verifica se a string de ações contém "/EDITAR/" e "/GESTOR/" (case-insensitive)
          If InStr(1, Session("Acoes"), "/EDITAR/", vbTextCompare) > 0 Then
              podeEditar = True
          End If

          If InStr(1, Session("Acoes"), "/GESTOR/", vbTextCompare) > 0 Then
              podeEditar = True
          End If
  Else
      response.redirect("login2.asp")
  End If

  ' Retorna o valor booleano
  UsuarioPodeEditar = podeEditar
End Function
%>