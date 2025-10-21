<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<%
Response.Buffer = True
Randomize Timer

If Len(StrConn) = 0 Or Len(StrConnSales) = 0 Then
    Response.Write "<p style='color:red'>Erro: Conexões não configuradas (StrConn / StrConnSales).</p>"
    Response.End
End If

Dim connMain, connSales
Set connMain  = Server.CreateObject("ADODB.Connection")
connMain.Open StrConn

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

Function SqlQuote(s)
    If IsNull(s) Then
        SqlQuote = "NULL"
    Else
        SqlQuote = "'" & Replace(CStr(s), "'", "''") & "'"
    End If
End Function

Function SqlDate(d)
    SqlDate = "'" & Year(d) & "-" & Right("0" & Month(d),2) & "-" & Right("0" & Day(d),2) & "'"
End Function

Function SqlNum(n)
    SqlNum = Replace(CStr(n), ",", ".")
End Function

Function GetRowsOrNull(sql, oConn)
    Dim rs, arr
    On Error Resume Next
    Set rs = oConn.Execute(sql)
    If Err.Number <> 0 Then
        Response.Write "<div style='color:red'>Erro executando SQL: " & Server.HTMLEncode(sql) & "<br>Detalhe: " & Err.Description & "</div>"
        Err.Clear
        GetRowsOrNull = Null
        Exit Function
    End If
    On Error GoTo 0

    If rs.EOF Then
        rs.Close : Set rs = Nothing
        GetRowsOrNull = Null
        Exit Function
    End If

    arr = rs.GetRows()
    rs.Close : Set rs = Nothing
    GetRowsOrNull = arr
End Function

Function RandIndex(maxVal)
    RandIndex = Int((maxVal + 1) * Rnd)
End Function

Function RandomDate(y, m)
    Dim lastDay
    lastDay = Day(DateSerial(y, m + 1, 0))
    RandomDate = DateSerial(y, m, Int(lastDay * Rnd) + 1)
End Function

Function PickUnidadeCodigo()
    Dim letras, a, b, num
    letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    a = Mid(letras, Int(26 * Rnd) + 1, 1)
    b = Mid(letras, Int(26 * Rnd) + 1, 1)
    num = Int(900 * Rnd) + 100
    PickUnidadeCodigo = "AP-" & a & b & "-" & num
End Function

Sub FlushLine(msg)
    Response.Write msg & "<br>" & vbCrLf
    Response.Flush
End Sub

Dim empreendimentos, diretorias, corretores
empreendimentos = GetRowsOrNull("SELECT empreend_id, NomeEmpreendimento, ComissaoVenda FROM empreendimento ORDER BY NomeEmpreendimento", connMain)
diretorias      = GetRowsOrNull("SELECT DiretoriaID, NomeDiretoria FROM Diretorias WHERE NomeDiretoria IN ('VILARIM','VERAS') ORDER BY NomeDiretoria", connMain)
corretores      = GetRowsOrNull("SELECT UserId, Nome FROM Usuarios WHERE Funcao='Corretor' AND Nome IS NOT NULL AND Nome <> '' ORDER BY Nome", connMain)

If IsNull(empreendimentos) Then
    FlushLine "<span style='color:red'>Nenhum Empreendimento encontrado.</span>"
    Response.End
End If
If IsNull(diretorias) Then
    FlushLine "<span style='color:red'>Nenhuma Diretoria encontrada.</span>"
    Response.End
End If
If IsNull(corretores) Then
    FlushLine "<span style='color:red'>Nenhum Corretor encontrado.</span>"
    Response.End
End If

Dim maxEmp, maxDir, maxCor
maxEmp = UBound(empreendimentos, 2)
maxDir = UBound(diretorias,  2)
maxCor = UBound(corretores,  2)

FlushLine "<strong>Iniciando geração de massa para 2025…</strong>"
FlushLine "Empreendimentos: " & (maxEmp+1) & " | Diretorias: " & (maxDir+1) & " | Corretores: " & (maxCor+1)

Dim anoAlvo:   anoAlvo   = 2025
Dim qtdPorMes: qtdPorMes = 1

Dim percDirDefault: percDirDefault = 5
Dim percGerDefault: percGerDefault = 10
Dim percCorDefault: percCorDefault = 35

Dim usuarioAtual
usuarioAtual = Session("Usuario")
If Trim(usuarioAtual & "") = "" Then usuarioAtual = "SISTEMA"

Dim mes, k
For mes = 1 To 12
    FlushLine "<hr><strong>Mês " & Right("0" & mes,2) & "/" & anoAlvo & "</strong>"

    For k = 1 To qtdPorMes

        Dim idxEmp, idEmp, nomeEmp, comissaoVendaEmp
        idxEmp = RandIndex(maxEmp)
        idEmp  = empreendimentos(0, idxEmp)
        nomeEmp = empreendimentos(1, idxEmp)
        comissaoVendaEmp = empreendimentos(2, idxEmp)
        If IsNull(comissaoVendaEmp) Or comissaoVendaEmp = "" Then comissaoVendaEmp = 3

        Dim idxDir, idDir, nomeDir
        idxDir = RandIndex(maxDir)
        idDir  = diretorias(0, idxDir)
        nomeDir= diretorias(1, idxDir)

        Dim gerArr, idGer, nomeGer
gerArr = GetRowsOrNull("SELECT GerenciaID, NomeGerencia FROM Gerencias WHERE DiretoriaID=" & idDir & " ORDER BY NomeGerencia", connMain)
If IsNull(gerArr) Then
    idGer = 0
    nomeGer = "Não aplicável"
Else
    Dim maxGer, idxGer
    maxGer = UBound(gerArr, 2)
    idxGer = RandIndex(maxGer)
    idGer  = gerArr(0, idxGer)
    nomeGer= gerArr(1, idxGer)
End If


        Dim idxCor, idCor, nomeCor
        idxCor = RandIndex(maxCor)
        idCor  = corretores(0, idxCor)
        nomeCor= corretores(1, idxCor)

        Dim valorUnidade, m2, dataVenda, unidadeCod
        valorUnidade = Int((500000 - 200000 + 1) * Rnd + 200000)
        m2 = Int((50 - 20 + 1) * Rnd + 20)
        dataVenda = RandomDate(anoAlvo, mes)
        unidadeCod = PickUnidadeCodigo()

        Dim diaVenda, mesVenda, anoVenda, trimestre
        diaVenda = Day(dataVenda)
        mesVenda = Month(dataVenda)
        anoVenda = Year(dataVenda)
        trimestre = Int((mesVenda - 1) / 3) + 1

        Dim percTotal, comissaoTotal, valorDir, valorGer, valorCor
        percTotal = CDbl(comissaoVendaEmp)
        comissaoTotal = CDbl(valorUnidade) * (percTotal / 100)

        valorDir = comissaoTotal * (percDirDefault / 100)
        valorGer = comissaoTotal * (percGerDefault / 100)
        valorCor = comissaoTotal * (percCorDefault / 100)

        Dim sqlVendas
        sqlVendas = "INSERT INTO Vendas (" & _
                    "Empreend_ID, NomeEmpreendimento, Unidade, UnidadeM2, " & _
                    "Corretor, CorretorId, ValorUnidade, ComissaoPercentual, ValorComissaoGeral, " & _
                    "DataVenda, DiaVenda, MesVenda, AnoVenda, Trimestre, " & _
                    "Obs, Usuario, DiretoriaId, Diretoria, GerenciaId, Gerencia, " & _
                    "ComissaoDiretoria, ValorDiretoria, ComissaoGerencia, ValorGerencia, ComissaoCorretor, ValorCorretor) VALUES (" & _
                    idEmp & ", " & SqlQuote(nomeEmp) & ", " & SqlQuote(unidadeCod) & ", " & SqlNum(m2) & ", " & _
                    SqlQuote(nomeCor) & ", " & idCor & ", " & SqlNum(valorUnidade) & ", " & SqlNum(percTotal) & ", " & SqlNum(comissaoTotal) & ", " & _
                    SqlDate(dataVenda) & ", " & diaVenda & ", " & mesVenda & ", " & anoVenda & ", " & trimestre & ", " & _
                    SqlQuote("Massa 2025 - auto") & ", " & SqlQuote(usuarioAtual) & ", " & idDir & ", " & SqlQuote(nomeDir) & ", " & idGer & ", " & SqlQuote(nomeGer) & ", " & _
                    SqlNum(percDirDefault) & ", " & SqlNum(valorDir) & ", " & SqlNum(percGerDefault) & ", " & SqlNum(valorGer) & ", " & SqlNum(percCorDefault) & ", " & SqlNum(valorCor) & ")"

        On Error Resume Next
        connSales.Execute sqlVendas
        If Err.Number = 0 Then
            FlushLine "[" & Right("0"&diaVenda,2) & "/" & Right("0"&mesVenda,2) & "/" & anoVenda & "] " & _
                      "Empreendimento: <strong>" & Server.HTMLEncode(nomeEmp) & "</strong> | Diretoria: " & Server.HTMLEncode(nomeDir) & _
                      " | Gerência: " & Server.HTMLEncode(nomeGer) & " | Corretor: " & Server.HTMLEncode(nomeCor) & _
                      " | Valor: R$ " & FormatNumber(valorUnidade, 2) & " | m²: " & m2 & " → <span style='color:green'>OK</span>"
        Else
            FlushLine "<span style='color:red'>Erro inserindo Vendas: " & Err.Description & "</span><br><code>" & Server.HTMLEncode(sqlVendas) & "</code>"
            Err.Clear
        End If
        On Error GoTo 0

    Next
Next

FlushLine "<hr><strong>Inserção de massa concluída!</strong>"

If Not connMain Is Nothing Then If connMain.State = 1 Then connMain.Close : Set connMain = Nothing
If Not connSales Is Nothing Then If connSales.State = 1 Then connSales.Close : Set connSales = Nothing




' ======================= ATUALIZAÇÃO FINAL DO BANCO DE DADOS =======================

'Dim connMain, connSales
Set connMain  = Server.CreateObject("ADODB.Connection")
connMain.Open StrConn

Set connSales = Server.CreateObject("ADODB.Connection")
connSales.Open StrConnSales

' Reaproveitar conexão já aberta: connSales (base Vendas)
On Error Resume Next

Response.Write dbSunnyPath
Response.end 



' --- UPDATE Diretoria ---
sqlUpdate1 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Diretorias INNER JOIN Vendas ON Diretorias.DiretoriaId = Vendas.DiretoriaId) " & _
             "SET Vendas.NomeDiretor = [Diretorias].[Nome], Vendas.UserIdDiretoria = [Diretorias].[UserId];"
connSales.Execute sqlUpdate1

' --- UPDATE Gerência ---
sqlUpdate2 = "UPDATE ([;DATABASE=" & dbSunnyPath & "].Gerencias INNER JOIN Vendas ON Gerencias.GerenciaId = Vendas.GerenciaId) " & _
             "SET Vendas.NomeGerente = [Gerencias].[Nome], Vendas.UserIdGerencia = [Gerencias].[UserId];"
connSales.Execute sqlUpdate2

' --- UPDATE Corretor ---
sqlUpdateCorretor = "UPDATE (Vendas INNER JOIN [;DATABASE=" & dbSunnyPath & "].Usuarios ON Vendas.CorretorId = Usuarios.UserId) " & _
                    "SET Vendas.Corretor = Usuarios.Nome;"
connSales.Execute sqlUpdateCorretor

' --- UPDATE Semestre ---
sqlUpdateSemestre = "UPDATE Vendas " & _
                    "SET Semestre = SWITCH(" & _
                    "    Trimestre IN (1, 2), 1, " & _
                    "    Trimestre IN (3, 4), 2" & _
                    ") " & _
                    "WHERE Trimestre IS NOT NULL;"
connSales.Execute sqlUpdateSemestre

' --- Verificação de erros ---
If Err.Number <> 0 Then
    Response.Write "<span style='color:red'>Ocorreu um erro ao atualizar o banco de dados: " & Err.Description & "</span><br>"
Else
    FlushLine "<span style='color:green'>Atualização final do banco de dados concluída com sucesso!</span>"
End If
On Error GoTo 0

' ======================= FIM DA ATUALIZAÇÃO =======================


%>
