' Mod4Media.bas
Option Explicit

'================================================================================
' GERENCIAMENTO DE CAMINHO DA IMAGEM DE CABECALHO
'================================================================================
Public Function GetHeaderImagePath() As String
    On Error GoTo ErrorHandler
    Dim headerImagePath As String

    ' Constroi caminho absoluto para a imagem desejada
    headerImagePath = Environ("USERPROFILE") & HEADER_IMAGE_RELATIVE_PATH

    ' Verifica se o arquivo existe
    If Dir(headerImagePath) = "" Then
        LogMessage "Imagem de cabecalho nao encontrada em: " & headerImagePath, LOG_LEVEL_WARNING
        GetHeaderImagePath = ""
        Exit Function
    End If

    GetHeaderImagePath = headerImagePath
    Exit Function

ErrorHandler:
    LogMessage "Erro ao localizar imagem de cabecalho: " & Err.Description, LOG_LEVEL_ERROR
    GetHeaderImagePath = ""
End Function

'================================================================================
' INSERCAO DE IMAGEM DE CABECALHO
'================================================================================
Public Function InsertHeaderstamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim imgFile As String
    Dim username As String
    Dim imgWidth As Single
    Dim imgHeight As Single
    Dim shp As shape
    Dim imgFound As Boolean
    Dim sectionsProcessed As Long

    ' Define o caminho da imagem do cabecalho
    imgFile = GetHeaderImagePath()

    If imgFile = "" Then
        Application.StatusBar = "Aviso: Imagem nao encontrada"
        InsertHeaderstamp = False
        Exit Function
    End If

    ' Size using standard constants
    imgWidth = CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
    imgHeight = imgWidth * HEADER_IMAGE_HEIGHT_RATIO

    For Each sec In doc.Sections
        Set header = sec.Headers(wdHeaderFooterPrimary)
        If header.Exists Then
            header.LinkToPrevious = False
            header.Range.Delete

            ' Define fonte padrao para o cabecalho: Arial 12
            With header.Range.Font
                .Name = STANDARD_FONT  ' Arial
                .size = STANDARD_FONT_SIZE  ' 12
            End With

            Set shp = header.Shapes.AddPicture(fileName:=imgFile, LinkToFile:=False, SaveWithDocument:=msoTrue)

            If shp Is Nothing Then
                LogMessage "Failed to insert header image at section " & sectionsProcessed + 1, LOG_LEVEL_WARNING
            Else
                With shp
                    .LockAspectRatio = msoTrue
                    .Width = imgWidth
                    .Height = imgHeight
                    .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                    .RelativeVerticalPosition = wdRelativeVerticalPositionPage
                    .Left = (doc.PageSetup.PageWidth - .Width) / 2
                    .Top = CentimetersToPoints(HEADER_IMAGE_TOP_MARGIN_CM)
                    .WrapFormat.Type = wdWrapTopBottom
                    .ZOrder msoSendToBack
                End With

                imgFound = True
                sectionsProcessed = sectionsProcessed + 1
            End If
        End If
    Next sec

    If imgFound Then
        ' Log detalhado removido para performance
        InsertHeaderstamp = True
    Else
    LogMessage "No header was inserted", LOG_LEVEL_WARNING
        InsertHeaderstamp = False
    End If

    Exit Function

ErrorHandler:
    LogMessage "Error inserting header: " & Err.Description, LOG_LEVEL_ERROR
    InsertHeaderstamp = False
End Function

'================================================================================
' IMAGE PROTECTION SYSTEM - SISTEMA DE PROTECAO DE IMAGENS
'================================================================================

'================================================================================
' BACKUP DE IMAGENS
'================================================================================
Public Function BackupAllImages(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Protegendo imagens..."

    imageCount = 0
    ReDim savedImages(0)

    Dim para As Paragraph
    Dim i As Long
    Dim j As Long
    Dim shape As InlineShape
    Dim tempImageInfo As ImageInfo

    ' Conta todas as imagens primeiro (com DoEvents para responsividade)
    Dim totalImages As Long
    For i = 1 To doc.Paragraphs.count
        If i Mod 30 = 0 Then DoEvents ' Responsividade
        Set para = doc.Paragraphs(i)
        totalImages = totalImages + para.Range.InlineShapes.count
    Next i

    ' Adiciona shapes flutuantes
    totalImages = totalImages + doc.Shapes.count

    ' Redimensiona array se necessario
    If totalImages > 0 Then
        ReDim savedImages(totalImages - 1)

        ' Backup de imagens inline - apenas propriedades criticas
        For i = 1 To doc.Paragraphs.count
            If i Mod 30 = 0 Then DoEvents ' Responsividade
            Set para = doc.Paragraphs(i)

            For j = 1 To para.Range.InlineShapes.count
                Set shape = para.Range.InlineShapes(j)

                ' Salva apenas propriedades essenciais para protecao
                With tempImageInfo
                    .paraIndex = i
                    .ImageIndex = j
                    .ImageType = "Inline"
                    .Position = shape.Range.Start
                    .Width = shape.Width
                    .Height = shape.Height
                    Set .AnchorRange = shape.Range.Duplicate
                    .ImageData = "InlineShape_Protected"
                End With

                savedImages(imageCount) = tempImageInfo
                imageCount = imageCount + 1

                ' Evita overflow
                If imageCount >= UBound(savedImages) + 1 Then Exit For
            Next j

            ' Evita overflow
            If imageCount >= UBound(savedImages) + 1 Then Exit For
        Next i

        ' Backup de shapes flutuantes - apenas propriedades criticas
        Dim floatingShape As shape
        For i = 1 To doc.Shapes.count
            Set floatingShape = doc.Shapes(i)

            If floatingShape.Type = msoPicture Then
                ' Redimensiona array se necessario
                If imageCount >= UBound(savedImages) + 1 Then
                    ReDim Preserve savedImages(imageCount)
                End If

                With tempImageInfo
                    .paraIndex = -1 ' Indica que e flutuante
                    .ImageIndex = i
                    .ImageType = "Floating"
                    .WrapType = floatingShape.WrapFormat.Type
                    .Width = floatingShape.Width
                    .Height = floatingShape.Height
                    .LeftPosition = floatingShape.Left
                    .TopPosition = floatingShape.Top
                    .ImageData = "FloatingShape_Protected"
                End With

                savedImages(imageCount) = tempImageInfo
                imageCount = imageCount + 1
            End If
        Next i
    End If

    LogMessage "Backup de propriedades de imagens concluido: " & imageCount & " imagens catalogadas"
    BackupAllImages = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao fazer backup de propriedades de imagens: " & Err.Description, LOG_LEVEL_WARNING
    BackupAllImages = False
End Function

'================================================================================
' RESTAURACAO DE IMAGENS
'================================================================================
Public Function RestoreAllImages(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If imageCount = 0 Then
        RestoreAllImages = True
        Exit Function
    End If

    Application.StatusBar = "Verificando integridade das imagens..."

    Dim i As Long
    Dim verifiedCount As Long
    Dim correctedCount As Long

    For i = 0 To imageCount - 1
        On Error Resume Next

        With savedImages(i)
            If .ImageType = "Inline" Then
                ' Verifica se a imagem inline ainda existe na posicao esperada
                If .paraIndex <= doc.Paragraphs.count Then
                    Dim para As Paragraph
                    Set para = doc.Paragraphs(.paraIndex)

                    ' Se ainda ha imagens inline no paragrafo, considera verificada
                    If para.Range.InlineShapes.count > 0 Then
                        verifiedCount = verifiedCount + 1
                    End If
                End If

            ElseIf .ImageType = "Floating" Then
                ' Verifica e corrige propriedades de shapes flutuantes se ainda existem
                If .ImageIndex <= doc.Shapes.count Then
                    Dim targetShape As shape
                    Set targetShape = doc.Shapes(.ImageIndex)

                    ' Verifica se as propriedades foram alteradas e corrige se necessario
                    Dim needsCorrection As Boolean
                    needsCorrection = False

                    If Abs(targetShape.Width - .Width) > 1 Then needsCorrection = True
                    If Abs(targetShape.Height - .Height) > 1 Then needsCorrection = True
                    If Abs(targetShape.Left - .LeftPosition) > 1 Then needsCorrection = True
                    If Abs(targetShape.Top - .TopPosition) > 1 Then needsCorrection = True

                    If needsCorrection Then
                        ' Restaura propriedades originais
                        With targetShape
                            .Width = savedImages(i).Width
                            .Height = savedImages(i).Height
                            .Left = savedImages(i).LeftPosition
                            .Top = savedImages(i).TopPosition
                            .WrapFormat.Type = savedImages(i).WrapType
                        End With
                        correctedCount = correctedCount + 1
                    End If

                    verifiedCount = verifiedCount + 1
                End If
            End If
        End With

        On Error GoTo ErrorHandler
    Next i

    If correctedCount > 0 Then
        LogMessage "Verificacao de imagens concluida: " & verifiedCount & " verificadas, " & correctedCount & " corrigidas"
    Else
        LogMessage "Verificacao de imagens concluida: " & verifiedCount & " imagens integras"
    End If

    RestoreAllImages = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao verificar imagens: " & Err.Description, LOG_LEVEL_WARNING
    RestoreAllImages = False
End Function

'================================================================================
' FORMAT IMAGE PARAGRAPHS INDENTS - Formata recuos de paragrafos com imagens
'================================================================================
Public Function FormatImageParagraphsIndents(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim formattedCount As Long
    formattedCount = 0

    ' Percorre todos os paragrafos
    Dim imgCounter As Long
    imgCounter = 0
    For Each para In doc.Paragraphs
        imgCounter = imgCounter + 1
        If imgCounter Mod 30 = 0 Then DoEvents ' Responsividade

        ' Verifica se o paragrafo contem imagens inline
        If para.Range.InlineShapes.count > 0 Then
            ' Zera o recuo a esquerda e centraliza
            With para.Format
                .leftIndent = 0
                .firstLineIndent = 0
                .alignment = wdAlignParagraphCenter
            End With
            formattedCount = formattedCount + 1
        End If
    Next para

    If formattedCount > 0 Then
        LogMessage "Recuos de paragrafos com imagens formatados: " & formattedCount & " paragrafos", LOG_LEVEL_INFO
    End If

    FormatImageParagraphsIndents = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao formatar recuos de imagens: " & Err.Description, LOG_LEVEL_WARNING
    FormatImageParagraphsIndents = False
End Function

'================================================================================
' BACKUP LIST FORMATS - Salva formatacoes de lista antes do processamento
'================================================================================
Public Function BackupListFormats(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim i As Long
    Dim tempListInfo As ListFormatInfo

    listFormatCount = 0
    ReDim savedListFormats(0)

    ' Conta quantos paragrafos tem formatacao de lista (com DoEvents)
    Dim totalLists As Long
    Dim countIter As Long
    totalLists = 0
    countIter = 0
    For Each para In doc.Paragraphs
        countIter = countIter + 1
        If countIter Mod 30 = 0 Then DoEvents ' Responsividade
        If para.Range.ListFormat.ListType <> 0 Then
            totalLists = totalLists + 1
        End If
    Next para

    If totalLists = 0 Then
        LogMessage "Nenhuma lista encontrada no documento", LOG_LEVEL_INFO
        BackupListFormats = True
        Exit Function
    End If

    ' Aloca array com tamanho adequado
    ReDim savedListFormats(totalLists - 1)

    ' Salva informacoes de cada paragrafo com lista (com DoEvents)
    Dim saveIter As Long
    saveIter = 0
    i = 1
    For Each para In doc.Paragraphs
        saveIter = saveIter + 1
        If saveIter Mod 30 = 0 Then DoEvents ' Responsividade

        If para.Range.ListFormat.ListType <> 0 Then
            With tempListInfo
                .paraIndex = i
                .HasList = True
                .ListType = para.Range.ListFormat.ListType

                ' Salva o nivel da lista se aplicavel
                On Error Resume Next
                .ListLevelNumber = para.Range.ListFormat.ListLevelNumber
                If Err.Number <> 0 Then
                    .ListLevelNumber = 1
                    Err.Clear
                End If
                On Error GoTo ErrorHandler

                ' Salva a string da lista (marcador ou numero)
                On Error Resume Next
                .ListString = para.Range.ListFormat.ListString
                If Err.Number <> 0 Then
                    .ListString = ""
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            End With

            savedListFormats(listFormatCount) = tempListInfo
            listFormatCount = listFormatCount + 1

            If listFormatCount >= UBound(savedListFormats) + 1 Then Exit For
        End If
        i = i + 1
    Next para

    LogMessage "Formatacoes de lista salvas: " & listFormatCount & " paragrafos com lista", LOG_LEVEL_INFO
    BackupListFormats = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao salvar formatacoes de lista: " & Err.Description, LOG_LEVEL_WARNING
    BackupListFormats = False
End Function

'================================================================================
' RESTORE LIST FORMATS - Restaura formatacoes de lista apos o processamento
'================================================================================
Public Function RestoreListFormats(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If listFormatCount = 0 Then
        RestoreListFormats = True
        Exit Function
    End If

    Dim i As Long
    Dim restoredCount As Long
    Dim para As Paragraph

    restoredCount = 0

    For i = 0 To listFormatCount - 1
        On Error Resume Next

        With savedListFormats(i)
            If .HasList And .paraIndex <= doc.Paragraphs.count Then
                Set para = doc.Paragraphs(.paraIndex)

                ' Remove qualquer formatacao de lista existente primeiro
                para.Range.ListFormat.RemoveNumbers

                ' Aplica a formatacao de lista original
                Select Case .ListType
                    Case 2 ' wdListBullet
                        ' Lista com marcadores
                        para.Range.ListFormat.ApplyBulletDefault

                    Case 3, 4 ' wdListSimpleNumbering, wdListListNumOnly
                        ' Lista numerada simples
                        para.Range.ListFormat.ApplyNumberDefault

                    Case 5 ' wdListMixedNumbering
                        ' Lista com numeracao mista
                        para.Range.ListFormat.ApplyNumberDefault

                    Case 6 ' wdListOutlineNumbering
                        ' Lista com numeracao de topicos
                        para.Range.ListFormat.ApplyOutlineNumberDefault

                    Case Else
                        ' Tenta aplicar formatacao padrao
                        If InStr(.ListString, ".") > 0 Or IsNumeric(Left(.ListString, 1)) Then
                            para.Range.ListFormat.ApplyNumberDefault
                        Else
                            para.Range.ListFormat.ApplyBulletDefault
                        End If
                End Select

                ' Tenta restaurar o nivel da lista
                If .ListLevelNumber > 0 And .ListLevelNumber <= 9 Then
                    para.Range.ListFormat.ListLevelNumber = .ListLevelNumber
                End If

                If Err.Number = 0 Then
                    restoredCount = restoredCount + 1
                Else
                    LogMessage "Aviso: Nao foi possivel restaurar lista no paragrafo " & .paraIndex & ": " & Err.Description, LOG_LEVEL_WARNING
                    Err.Clear
                End If
            End If
        End With

        On Error GoTo ErrorHandler
    Next i

    If restoredCount > 0 Then
        LogMessage "Formatacoes de lista restauradas: " & restoredCount & " paragrafos", LOG_LEVEL_INFO
    End If

    ' Limpa o array
    ReDim savedListFormats(0)
    listFormatCount = 0

    RestoreListFormats = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao restaurar formatacoes de lista: " & Err.Description, LOG_LEVEL_WARNING
    RestoreListFormats = False
End Function

'================================================================================
' CENTER IMAGE AFTER PLENARIO - Centraliza imagem entre 5a e 7a linha apos Plenario
'================================================================================
Public Function CenterImageAfterPlenario(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim i As Long
    Dim plenarioIndex As Long
    Dim paraText As String
    Dim paraTextCmp As String
    Dim lineCount As Long
    Dim centeredCount As Long

    plenarioIndex = 0
    centeredCount = 0

    ' Localiza o paragrafo "Plenario Dr. Tancredo Neves"
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(para.Range.text)
        paraTextCmp = NormalizeForComparison(paraText)

        ' Procura por "Plenario" e "Tancredo Neves" com $DATAATUALEXTENSO$
          If InStr(paraTextCmp, "plenario") > 0 And _
              InStr(paraTextCmp, "tancredo neves") > 0 And _
           InStr(paraText, "$DATAATUALEXTENSO$") > 0 Then
            plenarioIndex = i
            Exit For
        End If
    Next i

    ' Se nao encontrou o paragrafo do Plenario, retorna
    If plenarioIndex = 0 Then
        LogMessage "Paragrafo do Plenario nao encontrado para centralizar imagem", LOG_LEVEL_INFO
        CenterImageAfterPlenario = True
        Exit Function
    End If

    ' Verifica as linhas 5, 6 e 7 apos o Plenario (contando em branco e textuais)
    lineCount = 0
    For i = plenarioIndex + 1 To doc.Paragraphs.count
        lineCount = lineCount + 1

        ' Verifica apenas entre a 5 e 7 linha
        If lineCount >= 5 And lineCount <= 7 Then
            Set para = doc.Paragraphs(i)

            ' Se o paragrafo contem imagem, centraliza
            If para.Range.InlineShapes.count > 0 Then
                para.alignment = wdAlignParagraphCenter
                centeredCount = centeredCount + 1
                LogMessage "Imagem centralizada na linha " & lineCount & " apos Plenario", LOG_LEVEL_INFO
            End If
        End If

        ' Para apos a 7 linha
        If lineCount > 7 Then
            Exit For
        End If
    Next i

    If centeredCount > 0 Then
        LogMessage "Imagens centralizadas apos Plenario: " & centeredCount, LOG_LEVEL_INFO
    End If

    CenterImageAfterPlenario = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao centralizar imagem apos Plenario: " & Err.Description, LOG_LEVEL_WARNING
    CenterImageAfterPlenario = False
End Function









'================================================================================
' LIMPEZA DE PROTECAO DE IMAGENS
'================================================================================
Public Sub CleanupImageProtection()
    On Error Resume Next

    ' Limpa arrays de imagens
    If imageCount > 0 Then
        Dim i As Long
        For i = 0 To imageCount - 1
            Set savedImages(i).AnchorRange = Nothing
        Next i
    End If

    imageCount = 0
    ReDim savedImages(0)

    LogMessage "Variaveis de protecao de imagens limpas"
End Sub

