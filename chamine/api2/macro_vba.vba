Sub EnviarParaAPI()
    Dim http As Object
    Dim url As String
    Dim jsonBody As String
    Dim ws As Worksheet
    Dim wsPerfil As Worksheet
    Dim ultimaLinha As Integer
    Dim i As Integer

    ' URL da API
    url = "http://localhost/chamine/api2/apinova.php"

    ' Definir as planilhas
    Set ws = ActiveSheet
    Set wsPerfil = ThisWorkbook.Sheets("Perfil")

    ' Capturar data e hora do momento da execução
    Dim dataHoraExecucao As String
    dataHoraExecucao = Format(Now, "yyyy-MM-dd HH:mm:ss")

    ' Montar JSON com os dados do perfil
    jsonBody = "{""data_hora_execucao"": """ & dataHoraExecucao & """, ""perfil"": {"
    
    ' Capturar e tratar valores do perfil
    jsonBody = jsonBody & """Projeto"": """ & TrataValor(wsPerfil.Range("C5").Value) & """," & _
                         """Fonte"": """ & TrataValor(wsPerfil.Range("E5").Value) & """," & _
                         """Data"": """ & TrataData(wsPerfil.Range("G5").Value) & """," & _
                         """Equipe"": """ & TrataValor(wsPerfil.Range("J5").Value) & """," & _
                         """Empresa"": """ & TrataValor(wsPerfil.Range("C6").Value) & """," & _
                         """Cidade"": """ & TrataValor(wsPerfil.Range("J6").Value) & """," & _
                         """Processo"": """ & TrataValor(wsPerfil.Range("C7").Value) & """," & _
                         """Parametros"": """ & TrataValor(wsPerfil.Range("J7").Value) & """," & _
                         """Montante"": """ & TrataValor(wsPerfil.Range("C8").Value) & """," & _
                         """Jusante"": """ & TrataValor(wsPerfil.Range("E8").Value) & """," & _
                         """DiametroMedio"": """ & TrataValor(wsPerfil.Range("I8").Value) & """," & _
                         """Jdc"": """ & TrataValor(wsPerfil.Range("K8").Value) & """," & _
                         """Mdc"": """ & TrataValor(wsPerfil.Range("K9").Value) & """," & _
                         """AreaChamine"": """ & TrataValor(wsPerfil.Range("K10").Value) & """," & _
                         """Comprimento"": """ & TrataValor(wsPerfil.Range("C9").Value) & """," & _
                         """Diametroequivalente"": """ & TrataValor(wsPerfil.Range("E9").Value) & """," & _
                         """NumeroDePontos"": """ & TrataValor(wsPerfil.Range("C11").Value) & """," & _
                         """NumeroDeEixos"": """ & TrataValor(wsPerfil.Range("E11").Value) & """," & _
                         """NumeroDePontosPorEixo"": """ & TrataValor(wsPerfil.Range("G11").Value) & """," & _
                         """TempoTotalColeta"": """ & TrataValor(wsPerfil.Range("I11").Value) & """," & _
                         """TempoporPonto"": """ & TrataValor(wsPerfil.Range("K11").Value) & """}, ""coletas"": ["

    ' Definir última linha da coleta (se for fixo 25, pode deixar assim)
    ultimaLinha = 25

    ' Construir JSON com os dados das coletas
    For i = 4 To ultimaLinha
        jsonBody = jsonBody & "{""parametro"": """ & TrataValor(ws.Cells(i, 3).Value) & """," & _
                               """unidade"": """ & TrataValor(ws.Cells(i, 4).Value) & """," & _
                               """coleta_1"": """ & TrataValor(ws.Cells(i, 5).Value) & """," & _
                               """coleta_2"": """ & TrataValor(ws.Cells(i, 6).Value) & """," & _
                               """coleta_3"": """ & TrataValor(ws.Cells(i, 7).Value) & """},"
    Next i

    ' Remover a última vírgula extra
    If Right(jsonBody, 1) = "," Then
        jsonBody = Left(jsonBody, Len(jsonBody) - 1)
    End If

    ' Fechar JSON
    jsonBody = jsonBody & "]}"

    ' Criar objeto HTTP para envio da requisição
    Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json"
        .Send jsonBody
    End With

    ' Verificar a resposta da API
    MsgBox "Resposta da API: " & http.responseText

    ' Liberar objeto HTTP
    Set http = Nothing
End Sub




