Attribute VB_Name = "carregarArquivoServicos"
' ESTE PROCEDIMENTO ABRE O ARQUIVO CSV ESCOLHIDO E COLA NA PLANILHA PASSADA COMO PARAMETRO
Public Const ERRO_DE_CABECALHO As Long = vbObjectError + 513


Sub carregar_Servicos(ByVal planilha As String)

    On Error GoTo TE
    
    Dim arquivoEscolhido As String
    
      arquivoEscolhido = Application.GetOpenFilename("CSV File (*.csv), *.csv", , "Escolha um arquivo CSV de relat�rio de SERVI�OS", , False)
    
     
     
     Worksheets(planilha).Cells.Clear
     
    If Not (arquivoEscolhido = "Falso") Then
        
        With Worksheets(planilha).QueryTables.Add("TEXT;" + arquivoEscolhido, Worksheets(planilha).Range("A1"))
        
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = xlWindows
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = True
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
            
        End With
        
        Dim cabecalho As Variant
     
        cabecalho = Join(Application.Transpose(Application.Transpose(Worksheets("Configura��es").Range("G5:BM5"))), "")
                
        If Join(Application.Transpose(Application.Transpose(Worksheets("Servicos").UsedRange.Rows(1))), "") <> cabecalho Then

            Worksheets("Servicos").Cells.Clear

            userFormPrincipal.textboxServicos.Text = ""

            Err.Raise ERRO_DE_CABECALHO, "Erro de cabe�alho" _
                    , "Arquivo csv com cabe�alho inv�lido."

        End If
        
        userFormPrincipal.textboxServicos.Text = arquivoEscolhido
    
    
    End If
    
    
Exit Sub
TE:     'Tratamento de Erros

    MsgBox " Erro: " & Err.Description & Chr(13) & Chr(13) & "Local: M�dulo carregarArquivoServicos.carregar_Servicos"

End Sub
