Option Explicit

'================================================================================
' PONTO DE ENTRADA PRINCIPAL
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo CriticalErrorHandler

    executionStartTime = Now
    formattingCancelled = False
    undoGroupEnabled = False ' Reset inicial

    ' Verificacoes iniciais ANTES de iniciar UndoRecord
    If Not CheckWordVersion() Then
        Application.StatusBar = "Erro: Word 2010 ou superior necessario"
        LogMessage "Versao do Word " & Application.version & " nao suportada. Minimo: " & CStr(MIN_SUPPORTED_VERSION), LOG_LEVEL_ERROR
        MsgBox "Requer Word 2010 ou superior." & vbCrLf & _
               "Versao atual: " & Application.version, vbCritical, "Versao Incompativel"
        Exit Sub
    End If

    Dim doc As Document
    Set doc = Nothing

    On Error Resume Next
    Set doc = ActiveDocument
    If doc Is Nothing Then
        Application.StatusBar = "Erro: Nenhum documento aberto"
        MsgBox "Nenhum documento esta aberto para processamento.", vbCritical, "Erro"
        Exit Sub
    End If
    Err.Clear
    On Error GoTo CriticalErrorHandler
    ' ---------------------------------------------------------------------------

    ' Inicializa sistema de logging ANTES de qualquer LogMessage
    If Not InitializeLogging(doc) Then
        Application.StatusBar = "Aviso: Log desabilitado"
    End If

    ' Inicializa sistema de progresso (18 etapas do pipeline - 2 passagens)
    InitializeProgress 18

    If Not SetAppState(False, "Iniciando...") Then
        LogMessage "Falha ao configurar estado da aplicacao", LOG_LEVEL_WARNING
    End If

    IncrementProgress "Verificando documento"
    If Not PreviousChecking(doc) Then
        GoTo CleanUp
    End If

    If doc.Path = "" Then
        If Not SaveDocumentFirst(doc) Then
            Application.StatusBar = "Cancelado: documento nao salvo"
            LogMessage "Operacao cancelada - documento nao foi salvo", LOG_LEVEL_INFO
            GoTo CleanUp
        End If
    End If

    ' Cria backup do documento antes de qualquer modificacao
    IncrementProgress "Criando backup"
    If Not CreateDocumentBackup(doc) Then
        LogMessage "Falha ao criar backup - continuando sem backup", LOG_LEVEL_WARNING
    End If

    ' Backup das configuracoes de visualizacao originais
    IncrementProgress "Salvando configuracoes"
    If Not BackupViewSettings(doc) Then
        LogMessage "Aviso: Falha no backup das configuracoes de visualizacao", LOG_LEVEL_WARNING
    End If

    ' Backup de imagens antes das formatacoes
    IncrementProgress "Protegendo imagens"
    If Not BackupAllImages(doc) Then
        LogMessage "Aviso: Falha no backup de imagens - continuando com protecao basica", LOG_LEVEL_WARNING
    End If

    ' Backup de formatacoes de lista antes das formatacoes
    IncrementProgress "Protegendo listas"
    If Not BackupListFormats(doc) Then
        LogMessage "Aviso: Falha no backup de listas - formatacoes de lista podem ser perdidas", LOG_LEVEL_WARNING
    End If

    ' ---------------------------------------------------------------------------
    ' INICIO DO GRUPO DE DESFAZER (UndoRecord) - melhor esforco
    ' ---------------------------------------------------------------------------
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "CHAINSAW - Padronizacao"
    If Err.Number = 0 Then
        undoGroupEnabled = True
        LogMessage "UndoRecord iniciado", LOG_LEVEL_INFO
    Else
        undoGroupEnabled = False
        Err.Clear
    End If
    On Error GoTo CriticalErrorHandler
    ' ---------------------------------------------------------------------------

'================================================================================
' CONCLUIR - Copia ementa, salva e fecha com seguranca
'================================================================================
' FUNCOES PUBLICAS DE ACESSO AOS ELEMENTOS ESTRUTURAIS
'================================================================================

'--------------------------------------------------------------------------------
' GetTituloRange - Retorna o Range do titulo
'--------------------------------------------------------------------------------
Public Function GetTituloRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetTituloRange = Nothing

    If tituloParaIndex <= 0 Or tituloParaIndex > doc.Paragraphs.count Then Exit Function
    Set GetTituloRange = doc.Paragraphs(tituloParaIndex).Range
    Exit Function

ErrorHandler:
    Set GetTituloRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetEmentaRange - Retorna o Range da ementa
'--------------------------------------------------------------------------------
Public Function GetEmentaRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetEmentaRange = Nothing

    If ementaParaIndex <= 0 Or ementaParaIndex > doc.Paragraphs.count Then Exit Function
    Set GetEmentaRange = doc.Paragraphs(ementaParaIndex).Range
    Exit Function

ErrorHandler:
    Set GetEmentaRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetProposicaoRange - Retorna o Range da proposicao (conjunto de paragrafos)
'--------------------------------------------------------------------------------
Public Function GetProposicaoRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetProposicaoRange = Nothing

    If proposicaoStartIndex <= 0 Or proposicaoEndIndex <= 0 Then Exit Function
    If proposicaoStartIndex > doc.Paragraphs.count Then Exit Function
    If proposicaoEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long
    Dim endPos As Long

    startPos = doc.Paragraphs(proposicaoStartIndex).Range.Start
    endPos = doc.Paragraphs(proposicaoEndIndex).Range.End

    Set GetProposicaoRange = doc.Range(startPos, endPos)
    Exit Function

ErrorHandler:
    Set GetProposicaoRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetTituloJustificativaRange - Retorna o Range do titulo "Justificativa"
'--------------------------------------------------------------------------------
Public Function GetTituloJustificativaRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetTituloJustificativaRange = Nothing

    If tituloJustificativaIndex <= 0 Or tituloJustificativaIndex > doc.Paragraphs.count Then Exit Function
    Set GetTituloJustificativaRange = doc.Paragraphs(tituloJustificativaIndex).Range
    Exit Function

ErrorHandler:
    Set GetTituloJustificativaRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetJustificativaRange - Retorna o Range da justificativa (conjunto de paragrafos)
'--------------------------------------------------------------------------------
Public Function GetJustificativaRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetJustificativaRange = Nothing

    If justificativaStartIndex <= 0 Or justificativaEndIndex <= 0 Then Exit Function
    If justificativaStartIndex > doc.Paragraphs.count Then Exit Function
    If justificativaEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long
    Dim endPos As Long

    startPos = doc.Paragraphs(justificativaStartIndex).Range.Start
    endPos = doc.Paragraphs(justificativaEndIndex).Range.End

    Set GetJustificativaRange = doc.Range(startPos, endPos)
    Exit Function

ErrorHandler:
    Set GetJustificativaRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetDataRange - Retorna o Range da data (Plenario)
'--------------------------------------------------------------------------------
Public Function GetDataRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetDataRange = Nothing

    If dataParaIndex <= 0 Or dataParaIndex > doc.Paragraphs.count Then Exit Function
    Set GetDataRange = doc.Paragraphs(dataParaIndex).Range
    Exit Function

ErrorHandler:
    Set GetDataRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetAssinaturaRange - Retorna o Range da assinatura (3 paragrafos + imagens)
'--------------------------------------------------------------------------------
Public Function GetAssinaturaRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetAssinaturaRange = Nothing

    If assinaturaStartIndex <= 0 Or assinaturaEndIndex <= 0 Then Exit Function
    If assinaturaStartIndex > doc.Paragraphs.count Then Exit Function
    If assinaturaEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long
    Dim endPos As Long

    startPos = doc.Paragraphs(assinaturaStartIndex).Range.Start
    endPos = doc.Paragraphs(assinaturaEndIndex).Range.End

    Set GetAssinaturaRange = doc.Range(startPos, endPos)
    Exit Function

ErrorHandler:
    Set GetAssinaturaRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetTituloAnexoRange - Retorna o Range do titulo "Anexo" ou "Anexos"
'--------------------------------------------------------------------------------
Public Function GetTituloAnexoRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetTituloAnexoRange = Nothing

    If tituloAnexoIndex <= 0 Or tituloAnexoIndex > doc.Paragraphs.count Then Exit Function
    Set GetTituloAnexoRange = doc.Paragraphs(tituloAnexoIndex).Range
    Exit Function

ErrorHandler:
    Set GetTituloAnexoRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetAnexoRange - Retorna o Range do anexo (todo conteudo abaixo do titulo)
'--------------------------------------------------------------------------------
Public Function GetAnexoRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetAnexoRange = Nothing

    If anexoStartIndex <= 0 Or anexoEndIndex <= 0 Then Exit Function
    If anexoStartIndex > doc.Paragraphs.count Then Exit Function
    If anexoEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long
    Dim endPos As Long

    startPos = doc.Paragraphs(anexoStartIndex).Range.Start
    endPos = doc.Paragraphs(anexoEndIndex).Range.End

    Set GetAnexoRange = doc.Range(startPos, endPos)
    Exit Function

ErrorHandler:
    Set GetAnexoRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetProposituraRange - Retorna o Range de toda a propositura (documento completo)
'--------------------------------------------------------------------------------
Public Function GetProposituraRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetProposituraRange = Nothing

    If doc Is Nothing Then Exit Function
    Set GetProposituraRange = doc.Range
    Exit Function

ErrorHandler:
    Set GetProposituraRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetElementInfo - Retorna informacoes sobre todos os elementos identificados
' REFATORADO: Usa funcoes identificadoras ao inves de acesso direto as variaveis
'--------------------------------------------------------------------------------
Public Function GetElementInfo(doc As Document) As String
    On Error Resume Next

    Dim info As String
    Dim rng As Range

    info = "=== INFORMACOES DOS ELEMENTOS ESTRUTURAIS ===" & vbCrLf

    ' Titulo - usa GetTituloRange
    Set rng = GetTituloRange(doc)
    If Not rng Is Nothing Then
        info = info & "Titulo: Paragrafo " & tituloParaIndex & vbCrLf
    Else
        info = info & "Titulo: Nao identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Ementa - usa GetEmentaRange
    Set rng = GetEmentaRange(doc)
    If Not rng Is Nothing Then
        info = info & "Ementa: Paragrafo " & ementaParaIndex & vbCrLf
    Else
        info = info & "Ementa: Nao identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Proposicao - usa GetProposicaoRange
    Set rng = GetProposicaoRange(doc)
    If Not rng Is Nothing Then
        info = info & "Proposicao: Paragrafos " & proposicaoStartIndex & " a " & proposicaoEndIndex & _
                      " (" & (proposicaoEndIndex - proposicaoStartIndex + 1) & " paragrafos)" & vbCrLf
    Else
        info = info & "Proposicao: Nao identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Titulo Justificativa - ainda usa variavel direta (nao tem funcao Get especifica)
    If tituloJustificativaIndex > 0 Then
        info = info & "Titulo Justificativa: Paragrafo " & tituloJustificativaIndex & vbCrLf
    Else
        info = info & "Titulo Justificativa: Nao identificado" & vbCrLf
    End If

    ' Justificativa - usa GetJustificativaRange
    Set rng = GetJustificativaRange(doc)
    If Not rng Is Nothing Then
        info = info & "Justificativa: Paragrafos " & justificativaStartIndex & " a " & justificativaEndIndex & _
                      " (" & (justificativaEndIndex - justificativaStartIndex + 1) & " paragrafos)" & vbCrLf
    Else
        info = info & "Justificativa: Nao identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Data - usa GetDataRange
    Set rng = GetDataRange(doc)
    If Not rng Is Nothing Then
        info = info & "Data (Plenario): Paragrafo " & dataParaIndex & vbCrLf
    Else
        info = info & "Data (Plenario): Nao identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Assinatura - usa GetAssinaturaRange
    Set rng = GetAssinaturaRange(doc)
    If Not rng Is Nothing Then
        info = info & "Assinatura: Paragrafos " & assinaturaStartIndex & " a " & assinaturaEndIndex & _
                      " (" & (assinaturaEndIndex - assinaturaStartIndex + 1) & " paragrafos)" & vbCrLf
    Else
        info = info & "Assinatura: Nao identificado" & vbCrLf
    End If
    Set rng = Nothing

    If tituloAnexoIndex > 0 Then
        info = info & "Titulo Anexo: Paragrafo " & tituloAnexoIndex & vbCrLf
        If anexoStartIndex > 0 And anexoEndIndex > 0 Then
            info = info & "Anexo: Paragrafos " & anexoStartIndex & " a " & anexoEndIndex & _
                          " (" & (anexoEndIndex - anexoStartIndex + 1) & " paragrafos)" & vbCrLf
        End If
    Else
        info = info & "Anexo: Nao presente" & vbCrLf
    End If

    info = info & "============================================="

    GetElementInfo = info
End Function

'================================================================================
' SUBROTINA PUBLICA: ABRIR REPOSITORIO DO GITHUB
'================================================================================
Public Sub AbrirReadme()
    On Error GoTo ErrorHandler

    Const GITHUB_REPO_URL As String = "https://github.com/chrmsantos/chainsaw"

    ' Abre o repositorio do GitHub no navegador padrao
    Application.StatusBar = "Abrindo repositorio do GitHub..."

    ' Usa o comando Shell com o protocolo http:// para abrir no navegador padrao
    CreateObject("WScript.Shell").Run GITHUB_REPO_URL, 1, False

    ' Log da operacao se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Repositorio do GitHub aberto pelo usuario: " & GITHUB_REPO_URL, LOG_LEVEL_INFO
    End If

    Application.StatusBar = "Repositorio aberto no navegador"

    Exit Sub

ErrorHandler:
    Application.StatusBar = "Erro ao abrir repositorio"
    LogMessage "Erro ao abrir repositorio do GitHub: " & Err.Description, LOG_LEVEL_ERROR

    ' Tenta metodo alternativo
    On Error Resume Next
    shell "explorer.exe """ & GITHUB_REPO_URL & """", vbNormalFocus
End Sub

'================================================================================
' SUBROTINA PUBLICA: CONFIRMAR DESFAZIMENTO DA PADRONIZACAO
'================================================================================
Public Sub ConfirmarDesfazerPadronizacao()
    On Error GoTo ErrorHandler

    ' Verifica se ha um documento ativo
    Dim doc As Document
    Set doc = Nothing

    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        Exit Sub
    End If

    ' Verifica o numero de acoes disponiveis para desfazer
    Dim canUndo As Boolean
    canUndo = False

    On Error Resume Next
    canUndo = Application.CommandBars.ActionControl.enabled
    If Err.Number <> 0 Then canUndo = False
    On Error GoTo ErrorHandler

    ' Armazena informacoes antes do desfazer
    Dim beforeUndoCount As Long
    Dim docName As String
    Dim docPath As String

    beforeUndoCount = doc.Paragraphs.count
    docName = doc.Name
    docPath = doc.Path

    ' Executa o comando Desfazer (Undo)
    Application.StatusBar = "Desfazendo padronizacao..."
    On Error Resume Next
    doc.Undo
    On Error GoTo ErrorHandler

    ' Aguarda o Word processar o desfazer
    DoEvents

    ' Verifica se houve mudanca no documento
    Dim afterUndoCount As Long
    afterUndoCount = doc.Paragraphs.count

    ' Calcula a diferenca
    Dim changeCount As Long
    changeCount = Abs(beforeUndoCount - afterUndoCount)

    ' Cria mensagem informativa
    Dim undoMsg As String

    If changeCount > 0 Then
        undoMsg = "[<<] Padronizacao desfeita com sucesso!" & vbCrLf & vbCrLf & _
                  "[CHART] Alteracoes revertidas:" & vbCrLf & _
                  "    Paragrafos afetados: " & changeCount & vbCrLf & vbCrLf & _
                  "[DIR] Documento:" & vbCrLf & _
                  "   " & docName & vbCrLf & vbCrLf & _
                  "[i] DICA: O backup da padronizacao permanece disponivel." & vbCrLf & _
                  "   Use 'Abrir Pasta de Logs e Backups' para acessa-lo."
    Else
        undoMsg = "[<<] Desfazer executado!" & vbCrLf & vbCrLf & _
                  "[i] O documento foi revertido para o estado anterior." & vbCrLf & vbCrLf & _
                  "[DIR] Documento:" & vbCrLf & _
                  "   " & docName & vbCrLf & vbCrLf & _
                  "[i] DICA: O backup da padronizacao permanece disponivel." & vbCrLf & _
                  "   Use 'Abrir Pasta de Logs e Backups' para acessa-lo."
    End If

    ' Exibe mensagem de confirmacao
    MsgBox undoMsg, vbInformation, "CHAINSAW - Desfazer Padronizacao"

    ' Registra no log se estiver ativo
    If loggingEnabled Then
        LogMessage "Padronizacao desfeita pelo usuario - documento: " & docName, LOG_LEVEL_INFO
    End If

    Application.StatusBar = "Padronizacao desfeita"

    Exit Sub

ErrorHandler:
    Application.StatusBar = "Erro ao desfazer"

    ' Mensagem de erro generica
    MsgBox "Nao foi possivel desfazer a operacao." & vbCrLf & vbCrLf & _
           "[!] Possiveis causas:" & vbCrLf & _
           "    Nao ha operacoes para desfazer" & vbCrLf & _
           "    O documento foi fechado e reaberto" & vbCrLf & _
           "    Limite de desfazer atingido" & vbCrLf & vbCrLf & _
           "[i] SOLUCAO: Restaure manualmente a partir do backup." & vbCrLf & _
           "   Use 'Abrir Pasta de Logs e Backups' para acessar os backups.", _
           vbExclamation, "CHAINSAW - Erro ao Desfazer"

    If loggingEnabled Then
        LogMessage "Erro ao desfazer padronizacao: " & Err.Description, LOG_LEVEL_WARNING
    End If
End Sub

'================================================================================
' SUBROTINA PUBLICA: DESFAZER COM CONFIRMACAO AUTOMATICA
' Esta sub pode ser chamada diretamente ou apos o usuario usar Ctrl+Z
'================================================================================
Public Sub NotificarDesfazerPadronizacao()
    On Error Resume Next

    ' Verifica se ha um documento ativo
    Dim doc As Document
    Set doc = ActiveDocument

    If doc Is Nothing Then Exit Sub

    ' Cria mensagem de confirmacao simplificada
    Dim msg As String
    msg = "[<<] Padronizacao desfeita!" & vbCrLf & vbCrLf & _
          "[OK] Todas as alteracoes da ultima padronizacao foram revertidas." & vbCrLf & vbCrLf & _
          "[DIR] Documento: " & doc.Name & vbCrLf & vbCrLf & _
          "[SAVE] O backup continua disponivel na pasta de backups." & vbCrLf & _
          "   Use 'Abrir Pasta de Logs e Backups' para acessa-lo."

    ' Exibe notificacao
    MsgBox msg, vbInformation, "CHAINSAW - Operacao Desfeita"

    ' Log se disponivel
    If loggingEnabled Then
        LogMessage "Notificacao de desfazer exibida para: " & doc.Name, LOG_LEVEL_INFO
    End If
End Sub

