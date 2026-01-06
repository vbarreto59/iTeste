<%
' =================================================================================================
' FUNÇÃO PARA GRAVAR UM ARQUIVO DE LOG DE OPERAÇÕES DE DADOS
' =================================================================================================
' Esta função cria um arquivo de texto na pasta 'token' contendo o SQL da operação,
' o nome do usuário, IP do cliente e data/hora.
' 
' Parâmetros:
' - p_parametro1 (String): Uma string para identificar a operação (ex: "cadUsr").
' - p_sql (String): A string contendo o comando SQL (ex: "UPDATE ...").
'
' O arquivo é salvo com o nome: TK_<p_parametro1>_YYYYMMDD_HHMMSS.txt
' =================================================================================================
Sub GravaToken(p_parametro1, p_sql)
    
    ' Habilita a manipulação de erros para evitar que o script pare em caso de falha.
    On Error Resume Next
    
    Dim fso
    Dim caminhoPastaToken
    Dim nomeArquivo
    Dim caminhoCompleto
    Dim arquivo
    
    ' Obtém o caminho físico da pasta onde o script está sendo executado.
    Dim servidorCaminho
    servidorCaminho = Server.MapPath(".")
    
    ' Define o caminho completo para a pasta 'token'.
    caminhoPastaToken = "token\"
    
    ' Cria uma instância do objeto FileSystemObject.
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    
    ' Verifica se a pasta 'token' existe. Se não, cria a pasta.
    If Not fso.FolderExists(caminhoPastaToken) Then
        fso.CreateFolder caminhoPastaToken
    End If
    
    ' Verifica se a pasta 'token' foi criada com sucesso ou já existia.
    If fso.FolderExists(caminhoPastaToken) Then
        
        ' Formata a data e hora atual para o nome do arquivo, removendo caracteres inválidos.
        Dim dataParaNome
        dataParaNome = Replace(Replace(Replace(Now, "/", ""), ":", ""), " ", "_")
        
        ' Constrói o nome do arquivo.
        nomeArquivo = "TK_" & p_parametro1 & "_" & dataParaNome & ".txt"
        
        ' Constrói o caminho completo do arquivo.
        caminhoCompleto = caminhoPastaToken & "\" & nomeArquivo
        
        ' Cria o arquivo de texto. True indica que o arquivo será sobrescrito se já existir.
        Set arquivo = fso.CreateTextFile(caminhoCompleto, True)
        
        ' Escreve o conteúdo no arquivo.
        arquivo.WriteLine "========================================================"
        arquivo.WriteLine "  LOG DE OPERAÇÃO DE DADOS"
        arquivo.WriteLine "========================================================"
        arquivo.WriteLine "Usuário: " & Session("usuario")
        arquivo.WriteLine "IP: " & Request.ServerVariables("REMOTE_ADDR")
        arquivo.WriteLine "Data e Hora: " & Now
        arquivo.WriteLine "--------------------------------------------------------"
        arquivo.WriteLine "SQL:"
        arquivo.WriteLine p_sql
        arquivo.WriteLine "========================================================"
        
        ' Fecha o arquivo.
        arquivo.Close
        
    End If
    
    ' Libera os objetos da memória.
    Set arquivo = Nothing
    Set fso = Nothing
    
    ' Desabilita a manipulação de erros.
    On Error Goto 0
    
End Sub
%>