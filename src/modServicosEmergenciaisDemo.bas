Attribute VB_Name = "modServicosEmergenciaisDemo"
Option Explicit

' EXEMPLO DIDATICO INDEPENDENTE - SOMENTE DADOS FICTICIOS
' Nao e a macro utilizada em ambiente corporativo.
' Nao acessa SAP, Power Automate, rede, arquivos ou credenciais.
' A unica macro publica cria uma NOVA pasta de trabalho, sem salvar.
' O retorno DEMO-LOCAL identifica uma simulacao, nunca uma ordem SAP.

Private Const COL_ID As Long = 1
Private Const COL_RELATO As Long = 2
Private Const COL_EQUIPE As Long = 3
Private Const COL_INICIO As Long = 4
Private Const COL_FIM As Long = 5
Private Const COL_FINALIZADO As Long = 6
Private Const COL_PENDENCIA As Long = 7
Private Const COL_HORAS As Long = 8
Private Const COL_STATUS As Long = 9
Private Const COL_RETORNO As Long = 10
Private Const COL_ERRO As Long = 11

Public Sub ExecutarDemonstracao()
    Dim livroDemo As Workbook
    Dim planilha As Worksheet
    Dim linha As Long
    Dim simulados As Long
    Dim rejeitados As Long
    Dim motivo As String
    Dim cabecalhos As Variant
    Dim coluna As Long
    Dim tabela As ListObject

    On Error GoTo FalhaDemonstracao

    ' Destino separado: nunca grava dados no livro que contem a macro.
    Set livroDemo = Application.Workbooks.Add(xlWBATWorksheet)
    Set planilha = livroDemo.Worksheets(1)
    planilha.Name = "Demonstracao"

    cabecalhos = Array("ItemID", "RelatoFicticio", "EquipeFicticia", _
        "Inicio", "Fim", "ServicoFinalizado", "Pendencia", _
        "DuracaoHoras", "StatusDemo", "RetornoSimulado", "Validacao")

    For coluna = LBound(cabecalhos) To UBound(cabecalhos)
        planilha.Cells(1, coluna + 1).Value = cabecalhos(coluna)
    Next coluna

    CarregarCenarios planilha

    For linha = 2 To 6
        motivo = ValidarCenario(planilha, linha)
        If motivo = "" Then
            planilha.Cells(linha, COL_HORAS).Value = Round( _
                (CDbl(planilha.Cells(linha, COL_FIM).Value) - _
                 CDbl(planilha.Cells(linha, COL_INICIO).Value)) * 24#, 2)
            planilha.Cells(linha, COL_STATUS).Value = "SIMULADO"
            planilha.Cells(linha, COL_RETORNO).Value = _
                "DEMO-LOCAL-" & Format$(linha - 1, "000")
            planilha.Cells(linha, COL_ERRO).Value = "Validacao local aprovada"
            simulados = simulados + 1
        Else
            planilha.Cells(linha, COL_STATUS).Value = "REJEITADO"
            planilha.Cells(linha, COL_ERRO).Value = motivo
            rejeitados = rejeitados + 1
        End If
    Next linha

    Set tabela = planilha.ListObjects.Add( _
        SourceType:=xlSrcRange, Source:=planilha.Range("A1:K6"), _
        XlListObjectHasHeaders:=xlYes)
    tabela.Name = "tbCenariosDemo"
    tabela.TableStyle = "TableStyleMedium2"
    planilha.Range("D2:E6").NumberFormat = "dd/mm/yyyy hh:mm"
    planilha.Range("H2:H6").NumberFormat = "0.00"
    planilha.Columns("A:K").AutoFit
    planilha.Columns("B").ColumnWidth = 38
    planilha.Columns("G").ColumnWidth = 30
    planilha.Columns("K").ColumnWidth = 40
    planilha.Range("A1:K6").WrapText = True
    planilha.Rows("1:6").AutoFit

    ' Autoverificacao dos cenarios fixos; roda apenas no Excel do usuario.
    ConferirResultados planilha, simulados, rejeitados

    MsgBox "Demonstracao concluida: " & simulados & " simulados e " & _
        rejeitados & " rejeitados." & vbCrLf & _
        "Autoverificacao dos cenarios aprovada." & vbCrLf & _
        "Nenhum dado foi enviado ao SAP. Nenhum arquivo foi salvo.", _
        vbInformation, "Exemplo independente"
    Exit Sub

FalhaDemonstracao:
    MsgBox "A demonstracao foi interrompida: " & Err.Description & _
        vbCrLf & "Se criada, a pasta de demonstracao permanece aberta." & _
        vbCrLf & "Nenhum arquivo foi salvo automaticamente.", _
        vbExclamation, "Exemplo independente"
End Sub

Private Sub CarregarCenarios(ByVal destino As Worksheet)
    Dim dia As Date
    dia = DateSerial(2026, 1, 15)

    AdicionarCenario destino, 2, "CASO-001", _
        "Conexao de equipamento ficticio ajustada e testada.", _
        "Equipe Alfa", dia + TimeSerial(8, 0, 0), _
        dia + TimeSerial(9, 30, 0), True, ""

    AdicionarCenario destino, 3, "CASO-002", _
        "Inspecao ficticia; troca de componente pendente.", _
        "Equipe Beta", dia + TimeSerial(10, 0, 0), _
        dia + TimeSerial(12, 0, 0), False, "Aguardando componente ficticio."

    ' Erro intencional: relato vazio.
    AdicionarCenario destino, 4, "CASO-003", "", _
        "Equipe Alfa", dia + TimeSerial(13, 0, 0), _
        dia + TimeSerial(14, 0, 0), True, ""

    ' Erro intencional: horario final anterior ao inicial.
    AdicionarCenario destino, 5, "CASO-004", _
        "Teste ficticio com horarios inconsistentes.", _
        "Equipe Gama", dia + TimeSerial(16, 0, 0), _
        dia + TimeSerial(15, 0, 0), True, ""

    ' Cenario valido atravessando a meia-noite.
    AdicionarCenario destino, 6, "CASO-005", _
        "Verificacao ficticia em mudanca de dia.", _
        "Equipe Gama", dia + TimeSerial(23, 45, 0), _
        dia + 1 + TimeSerial(0, 15, 0), True, ""
End Sub

Private Sub AdicionarCenario(ByVal destino As Worksheet, _
    ByVal linha As Long, ByVal item As String, ByVal relato As String, _
    ByVal equipe As String, ByVal inicio As Date, ByVal fim As Date, _
    ByVal finalizado As Boolean, ByVal pendencia As String)

    destino.Cells(linha, COL_ID).Value = item
    destino.Cells(linha, COL_RELATO).Value = relato
    destino.Cells(linha, COL_EQUIPE).Value = equipe
    destino.Cells(linha, COL_INICIO).Value = inicio
    destino.Cells(linha, COL_FIM).Value = fim
    destino.Cells(linha, COL_FINALIZADO).Value = finalizado
    destino.Cells(linha, COL_PENDENCIA).Value = pendencia
    destino.Cells(linha, COL_STATUS).Value = "AGUARDANDO"
End Sub

Private Function ValidarCenario(ByVal dados As Worksheet, _
    ByVal linha As Long) As String

    ' Entrada interna tipada, criada por CarregarCenarios.
    ' Nao e um importador nem validador completo de planilhas externas.
    If Trim$(CStr(dados.Cells(linha, COL_ID).Value)) = "" Then
        ValidarCenario = "Identificador obrigatorio."
    ElseIf Trim$(CStr(dados.Cells(linha, COL_RELATO).Value)) = "" Then
        ValidarCenario = "Relato obrigatorio."
    ElseIf Trim$(CStr(dados.Cells(linha, COL_EQUIPE).Value)) = "" Then
        ValidarCenario = "Equipe obrigatoria."
    ElseIf dados.Cells(linha, COL_FIM).Value <= _
           dados.Cells(linha, COL_INICIO).Value Then
        ValidarCenario = "Fim deve ser posterior ao inicio."
    ElseIf Not CBool(dados.Cells(linha, COL_FINALIZADO).Value) And _
           Trim$(CStr(dados.Cells(linha, COL_PENDENCIA).Value)) = "" Then
        ValidarCenario = "Servico nao finalizado exige motivo da pendencia."
    End If
End Function

Private Sub ConferirResultados(ByVal dados As Worksheet, _
    ByVal simulados As Long, ByVal rejeitados As Long)

    If simulados <> 3 Or rejeitados <> 2 Then
        Err.Raise vbObjectError + 701, "Demo", "Contagem inesperada."
    End If
    If dados.Cells(2, COL_HORAS).Value <> 1.5 Or _
       dados.Cells(3, COL_HORAS).Value <> 2 Or _
       dados.Cells(6, COL_HORAS).Value <> 0.5 Then
        Err.Raise vbObjectError + 702, "Demo", "Duracao inesperada."
    End If
    If dados.Cells(4, COL_ERRO).Value <> "Relato obrigatorio." Or _
       dados.Cells(5, COL_ERRO).Value <> "Fim deve ser posterior ao inicio." Then
        Err.Raise vbObjectError + 703, "Demo", "Validacao inesperada."
    End If
    If dados.Cells(4, COL_RETORNO).Value <> "" Or _
       dados.Cells(5, COL_RETORNO).Value <> "" Then
        Err.Raise vbObjectError + 704, "Demo", "Retorno em caso rejeitado."
    End If
End Sub
