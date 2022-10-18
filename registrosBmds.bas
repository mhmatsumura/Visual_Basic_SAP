Attribute VB_Name = "registrosBmds"
Sub salvaColecaoDeLinhasNoCsvRegistro(ByVal caminhoArquivo As String, ByVal colecaoLinhas As Collection)

    Set fso = New Scripting.FileSystemObject
    Set arquivoCsv = fso.OpenTextFile(caminhoArquivo, ForAppending, True, TristateFalse)

    'inclui todas as  linhas da colecaoLinhas no csv registro
    For Each linha In colecaoLinhas
     arquivoCsv.WriteLine linha(1) & ";" & linha(2) & ";" & linha(3) & ";" & linha(4)
    Next
    
   
    arquivoCsv.Close

End Sub
Sub criaArquivoCsvRegistroVazio(ByVal caminhoArquivo As String)
   
    Set fso = New Scripting.FileSystemObject

    Set arquivoCsv = fso.CreateTextFile(caminhoArquivo, True)
    
    'inclui a duas primeiras linhas do csv vazio
    arquivoCsv.WriteLine "NUM_BMD;SEQUENCIA_BMD;NUM_OS;NUM_PEDIDO"
   
    arquivoCsv.Close

End Sub
Function arquivoExiste(ByVal caminhoArquivo As String)
    If Dir(caminhoArquivo) = "" Then
        arquivoExiste = False
    Else
        arquivoExiste = True
    End If
End Function
Function arquivoRegistroEhFormatoValido(ByVal caminhoArquivo As String)

    Dim numeroArquivo As Integer
    Dim linhaArquivo As String
    Dim linhaEsplitada As Variant
    
    numeroArquivo = FreeFile
    arquivoRegistroEhFormatoValido = True
    
    Open caminhoArquivo For Input As #numeroArquivo
    
    
    'verifica se primeira linha corresponde a "NUM_BMD;SEQUENCIA_BMD;NUM_OS;NUM_PEDIDO"
    Line Input #numeroArquivo, linhaArquivo

    If linhaArquivo <> "NUM_BMD;SEQUENCIA_BMD;NUM_OS;NUM_PEDIDO" Then
     arquivoRegistroEhFormatoValido = False
    End If

    Close #numeroArquivo

End Function
Function consultaUltimoBmd()

    Dim numeroArquivo As Integer
    Dim linhaArquivo As String
 
   
    numeroArquivo = FreeFile
    caminhoArquivo = Application.ActiveWorkbook.Path & "\registroLancamentosBmdsBD.csv"
    
   
    Open caminhoArquivo For Input As #numeroArquivo
   
 
    'percorre linha por linha do arquivo at� o EOF(end of file)
    CONT = 0
    Do Until EOF(numeroArquivo)
        Line Input #numeroArquivo, linhaArquivo
        CONT = CONT + 1
    Loop
    
    linhaEsplitada = Split(linhaArquivo, ";")
    
    
    If CONT = 1 Then
     consultaUltimoBmd = 0
    Else
     consultaUltimoBmd = linhaEsplitada(0)
    End If
    
    
    Close #numeroArquivo

End Function
Function termoExisteNoArquivo(ByVal caminhoArquivo As String, _
                   ByVal indiceColuna As Long, ByVal termoProcurado As String)
                   
    Dim numeroArquivo As Integer
    Dim linhaArquivo As String
    Dim linhaEsplitada As Variant
    
    termoExisteNoArquivo = False
    numeroArquivo = FreeFile
    
    Open caminhoArquivo For Input As #numeroArquivo
    
    contLinha = 0
    'percorre linha por linha do arquivo at� o EOF(end of file)
    Do Until EOF(numeroArquivo)
        Line Input #numeroArquivo, linhaArquivo
        'pula a primeira linha do arquivo
        If contLinha > 0 Then
            linhaEsplitada = Split(linhaArquivo, ";")
            If linhaEsplitada(indiceColuna - 1) = termoProcurado Then
                termoExisteNoArquivo = True
                Exit Do
            End If
        End If
        contLinha = contLinha + 1
    Loop

    Close #numeroArquivo
End Function















