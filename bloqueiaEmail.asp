<!-- ###################################### -->
<!-- SISTEMA: SGVENDAS                      -->
<!-- AUTOR: VALTER BARRETO                    -->
<!-- Data: 04/12/2025               -->
<!-- CODIGO_ARQUIVO: ZDKGYLUODJ          -->
<!-- OBS:                                     -->
<!-- ###################################### -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' Variáveis de configuração
Const ARQUIVO_BLOQUEIO = "NoEmail.txt"
Dim strCaminhoArquivo, objFSO, blnBloqueado
Dim strMensagem, strAcao, strCor ' strCor agora usa classes de alerta do Bootstrap

' Cria o caminho absoluto para o arquivo
strCaminhoArquivo = Server.MapPath(ARQUIVO_BLOQUEIO)

' Inicializa o FileSystemObject
On Error Resume Next
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

If Err.Number <> 0 Then
    ' Se não for possível criar o objeto (permissão ou configuração), exibe erro
    strMensagem = "ERRO: Não foi possível inicializar o FileSystemObject. Verifique as permissões do servidor."
    strCor = "alert-danger"
Else
    ' 1. Verifica a ação da QueryString
    strAcao = LCase(Request.QueryString("acao"))

    Select Case strAcao
        Case "criar"
            ' Tenta criar o arquivo (Bloquear)
            If Not objFSO.FileExists(strCaminhoArquivo) Then
                objFSO.CreateTextFile strCaminhoArquivo
                strMensagem = "Sucesso: O arquivo '" & ARQUIVO_BLOQUEIO & "' foi CRIADO. O envio de e-mail está BLOQUEADO."
                strCor = "alert-success" ' Mapeado de bg-green-600
            Else
                strMensagem = "Aviso: O arquivo '" & ARQUIVO_BLOQUEIO & "' já existe. O envio de e-mail JÁ ESTÁ BLOQUEADO."
                strCor = "alert-warning" ' Mapeado de bg-yellow-600
            End If
        
        Case "excluir"
            ' Tenta excluir o arquivo (Desbloquear)
            If objFSO.FileExists(strCaminhoArquivo) Then
                objFSO.DeleteFile strCaminhoArquivo, True ' True para forçar
                strMensagem = "Sucesso: O arquivo '" & ARQUIVO_BLOQUEIO & "' foi EXCLUÍDO. O envio de e-mail está DESBLOQUEADO."
                strCor = "alert-success" ' Mapeado de bg-green-600
            Else
                strMensagem = "Aviso: O arquivo '" & ARQUIVO_BLOQUEIO & "' não existe. O envio de e-mail JÁ ESTÁ DESBLOQUEADO."
                strCor = "alert-warning" ' Mapeado de bg-yellow-600
            End If
    End Select
    
    ' 2. Verifica o status atual do arquivo após a ação (ou sem ação)
    blnBloqueado = objFSO.FileExists(strCaminhoArquivo)
    
    ' Define a mensagem de status inicial se nenhuma ação foi executada
    If strAcao = "" Then
        If blnBloqueado Then
            strMensagem = "STATUS ATUAL: O envio de e-mail está BLOQUEADO."
            strCor = "alert-danger" ' Mapeado de bg-red-600
        Else
            strMensagem = "STATUS ATUAL: O envio de e-mail está DESBLOQUEADO."
            strCor = "alert-primary" ' Mapeado de bg-blue-600
        End If
    End If

    If blnBloqueado then
       Session("EnviaEmail") = "NoEmail"
    Else
       Session("EnviaEmail") = "**"
    End If 
End If

Set objFSO = Nothing
Err.Clear
On Error GoTo 0 ' Restaura tratamento de erro
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gerenciamento de E-mail</title>
    <!-- Inclusão do Bootstrap 5 CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f8f9fa; }
        /* Centraliza o conteúdo na tela */
        html, body { height: 100%; }
        .full-height-container {
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
        }
    </style>
</head>
<body class="full-height-container p-4">

    <!-- Container principal (substituindo o max-w-lg) -->
    <div class="container">
        <div class="row justify-content-center">
            <div class="col-sm-12 col-md-8 col-lg-6">
                <div class="card shadow-lg border-0 rounded-3 p-4 p-md-5">
                    
                    <h1 class="h3 fw-bold text-dark text-center mb-4">Controle de Envio de E-mails</h1>
                    <p class="text-secondary text-center mb-4">
                        Esta página gerencia a existência do arquivo 
                        <span class="badge bg-light text-dark border border-secondary fw-normal">
                            <%= ARQUIVO_BLOQUEIO %>
                        </span>.
                        Se o arquivo existir, o envio de e-mails de login é BLOQUEADO.
                    </p>
                    
                    <!-- Status Atual (usando classes alert do Bootstrap) -->
                    <div class="alert <%= strCor %> text-center fw-semibold mb-5" role="alert">
                        <%= strMensagem %>
                    </div>

                    <!-- Ações -->
                    <div class="d-grid gap-3">
                        <% If blnBloqueado Then %>
                            <!-- Se está bloqueado, mostra o botão de Desbloquear (btn-primary) -->
                            <a href="?acao=excluir" class="btn btn-primary btn-lg">
                                &#128275; Desbloquear E-mail (Excluir <%= ARQUIVO_BLOQUEIO %>)
                            </a>
                        <% Else %>
                            <!-- Se está desbloqueado, mostra o botão de Bloquear (btn-danger) -->
                            <a href="?acao=criar" class="btn btn-danger btn-lg">
                                &#128274; Bloquear Envio de E-mail (Criar <%= ARQUIVO_BLOQUEIO %>)
                            </a>
                        <% End If %>
                        
                        <!-- Botão para recarregar/verificar status -->
                        <a href="bloqueiaEmail.asp" class="btn btn-outline-secondary btn-lg">
                            &#8635; Recarregar Status
                        </a>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Inclusão do Bootstrap JS (opcional para botões, mas boa prática) -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>