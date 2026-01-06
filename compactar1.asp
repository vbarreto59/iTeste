<%
' Configurações do Banco de Dados
OldDBPath = "C:\inetpub\wwwroot\iTeste\db\SunSales.mdb" ' Caminho completo do banco original
NewDBPath = "C:\inetpub\wwwroot\iTeste\db\SunSales.mdb" ' Caminho completo do novo banco compactado
BackupDBPath = "C:\inetpub\wwwroot\iTeste\db\backup/SunSales.mdb"

' String de conexão (Provider)
Provider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="

' 1. Cria um objeto FileSystemObject (FSO) para manipulação de arquivos
Set FSO = CreateObject("Scripting.FileSystemObject")

' 2. Cria um backup do arquivo original por segurança
If FSO.FileExists(OldDBPath) Then
    FSO.CopyFile OldDBPath, BackupDBPath, True ' O 'True' permite sobrescrever
End If

' 3. Compacta o banco de dados
Set Engine = CreateObject("JRO.JetEngine")

'On Error Resume Next ' Ignora erros, caso o arquivo compactado não possa ser criado

Engine.CompactDatabase Provider & OldDBPath, Provider & NewDBPath

If Err.Number <> 0 Then
    Response.Write "Erro ao compactar o banco de dados: " & Err.Description
    Set Engine = Nothing
    Set FSO = Nothing
    Response.End
End If

On Error GoTo 0 ' Volta a tratar erros

Set Engine = Nothing

' 4. Substitui o arquivo original pelo arquivo compactado
If FSO.FileExists(NewDBPath) Then
    ' Apaga o arquivo original (OldDB)
    FSO.DeleteFile OldDBPath
    
    ' Renomeia o novo arquivo compactado (NewDB) para o nome original (OldDB)
    FSO.MoveFile NewDBPath, OldDBPath
    
    Response.Write "Banco de dados compactado e substituído com sucesso."
Else
    Response.Write "Erro: O arquivo compactado não foi criado."
End If

' Libera os objetos
Set FSO = Nothing
%>