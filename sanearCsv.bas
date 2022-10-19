Attribute VB_Name = "sanearCsv"
'EDITA O ARQUIVO CSV ORIGINAL ELIMINANDO COLUNAS NAO UTEIS AO PROJETO
Public Const ERRO_OS_INEXISTENTE As Long = vbObjectError + 514
Public Const ERRO_PLANILHA_VAZIA As Long = vbObjectError + 515
Public mesAno As String
Dim colecaoPendencias As Collection

Sub sanear_Csv()

   On Error GoTo TE
   
   Set colecaoPendencias = New Collection
   
   'Application.ScreenUpdating = False
   
   
   
If planilhaVazia(Worksheets("Os")) Then

    Err.Raise ERRO_PLANILHA_VAZIA, "Planilha vazia!" _
                        , "Nenhum relat�rio de OS foi carregado."
                        
End If

If planilhaVazia(Worksheets("Servicos")) Then

    Err.Raise ERRO_PLANILHA_VAZIA, "Planilha vazia!" _
                        , "Nenhum relat�rio de SERVICOS foi carregado."
                        
End If

'     'USADO EM TESTES
'                'FAZ UMA C�PIA DA PLANILHA "Servicos" NA PLANILHA "Serv_Cprog"
'                Worksheets("Servicos").UsedRange.Copy Destination:=Worksheets("Serv_Cprog").Range("A1")
    
  
    'AUTOFILTRA PELO TERMO CPROG E COPIA PARA PLANILHA Serv_Cprog
    empreiteira = Worksheets("Configura��es").Range("C7").Value
    Worksheets("Servicos").Range("A1").AutoFilter Field:=7, Criteria1:=empreiteira
    Worksheets("Servicos").AutoFilter.Range.Copy Destination:=Worksheets("Serv_Cprog").Range("A1")
    Worksheets("Servicos").AutoFilterMode = False
    
If planilhaVazia(Worksheets("Serv_Cprog")) Then

    Err.Raise ERRO_PLANILHA_VAZIA, "Planilha vazia!" _
                        , "N�o h� itens da equipe CPROG a serem processados."
                        
End If

'============================================================================

    'AUTOFILTRA PELO TERMO CPROG E COPIA PARA PLANILHA Os_Cprog
    Worksheets("Os").Range("A1").AutoFilter Field:=6, Criteria1:=empreiteira
    Worksheets("Os").AutoFilter.Range.Copy Destination:=Worksheets("Os_Cprog").Range("A1")
    Worksheets("Os").AutoFilterMode = False
    
    'FAZ UMA C�PIA DA PLANILHA "Serv_Cprog" NA PLANILHA "Unificado"
     Worksheets("Serv_Cprog").UsedRange.Copy Destination:=Worksheets("Unificado").Range("A1")
     
     
     
      
            'verificando se todas as os da planilha servico possui correspondente na planilha os de onde vira as informacoes de qtd km, status e data de fechamento da os
             eliminaOsAusente
            
            
   
    
'UNIFICA AS PLANILHAS Serv_Cprog E OS_Cprog INSERINDO A COLUNA ESTADO, DESLOCAMENTO e DATA_FECHAMENTO;
unificaServicosOs

 
    
'FORMATA COLUNA DATA DE FECHAMENTO NA PLANILHA UNIFICADO COMO MES E ANO
 formataComoMes Worksheets("Unificado"), coluna:=3
 formataComoMes Worksheets("Unificado"), coluna:=14

 
    extraiEmBaixa
    extraiNovas
                
    'AUTOFILTRA PLANILHA UNIFICADO PELO TERMO FECHADA E MES/ANO
    Worksheets("Unificado").Range("A1").AutoFilter Field:=2, Criteria1:="FECHADA"
    Worksheets("Unificado").Range("A1").AutoFilter Field:=3, Criteria1:=mesAno
    Worksheets("Unificado").AutoFilter.Range.Copy Destination:=Worksheets("Col_Interesse").Range("A1")
    Worksheets("Unificado").AutoFilterMode = False
    

If planilhaVazia(Worksheets("Col_Interesse")) Then

    Err.Raise ERRO_PLANILHA_VAZIA, "Planilha vazia!" _
                        , "N�o h� OS FECHADA do M�S/ANO especificado para ser processada."

End If
 
'altera a ordem das colunas de interesse para o lado esquerdo na planilha "Col_Interesse".
formataColunasInteresse
    
    'FORMATA COLUNA DATA DE EXECUCAO NA PLANILHA Col_Interesse COMO MES E ANO
     formataComoMes Worksheets("Col_Interesse"), coluna:=1
     formataComoMes Worksheets("Col_Interesse"), coluna:=9
     formataComoMes Worksheets("Col_Interesse"), coluna:=30
     
                
    
    'FAZ UMA C�PIA DA PLANILHA "Col_Interesse" NA PLANILHA "Car"
     Worksheets("Col_Interesse").UsedRange.Copy Destination:=Worksheets("Car").Range("A1")
     
    'verifica se arquivo de registro de lancamentos existe e eh valido e entao filtra as os ja pagas
    caminhoArquivo = Application.ActiveWorkbook.Path & "\registroLancamentosBmdsBD.csv"
    If Not arquivoExiste(caminhoArquivo) Then
        criaArquivoCsvRegistroVazio (caminhoArquivo)
        MsgBox "ARQUIVO DE REGISTROS N�O ENCONTRADO!!" & Chr(13) & Chr(13) & "FOI CRIADO ARQUIVO DE REGISTRO DE LAN�AMENTOS VAZIO PARA CONTROLE DE OS�s J� PAGAS E CONTROLE DE SEQU�NCIA DE N�MERO DOS BMD�s."
        adicionaPendencia "Informa��o!", "Qtd de OS j� pagas e ignoradas: 0", ""
    Else
     If arquivoRegistroEhFormatoValido(caminhoArquivo) Then
     
        'consulta arquivo de registro e elimina todas as linhas da planilha CAR que contem num_os j� paga
        filtraOsPaga

      Else
        Err.Raise vbObjectError + 50, " INV�LIDO!" _
             , "Cabe�alho do arquivo de registro � inv�lido!! Verifique Arquivo CSV de registro de lan�amentos!"
      End If
      
    End If


If planilhaVazia(Worksheets("Car")) Then

    Err.Raise ERRO_PLANILHA_VAZIA, "Planilha vazia!" _
                        , "Todas as OS�s do M�S/ANO especificado j� foram pagas!"

End If
            
           
            
            'DETERMINA O TAMANHO DAS COLUNAS"
            Worksheets("Car").Columns("A:Z").ColumnWidth = 15
            'ABRE A PLANILHA "Car"
            Worksheets("Car").Activate
 
            
    'ISOLA O CAR E CONSULTA NOME DA CIDADE
     converteCidade coluna:=2
     
           
     
            'FAZ UMA C�PIA DA PLANILHA "Car" NA PLANILHA "Atividade"
            Worksheets("Car").UsedRange.Copy Destination:=Worksheets("Atividade").Range("A1")
            'DETERMINA O TAMANHO DAS COLUNAS"
            Worksheets("Atividade").Columns("A:Z").ColumnWidth = 15
            'ABRE A PLANILHA "Atividade"
            Worksheets("Atividade").Activate
    
    'EXTRAI AS ATIVIDADES, REMOVE DUPLICADAS
     extraiAtividade coluna:=4
    

'     unificaDeslocamentoAtividade
    
If planilhaVazia(Worksheets("Atividade")) Then
    MsgBox "N�o h� OS com ATIVIDADE v�lida para ser processada. Verifique planilha de Pend�ncias!"
End If
            
            'ATRIBUI AO CABECALHO DA COLUNA COMENTARIO O VALOR ATIVIDADE
             Worksheets("Atividade").Range("D1").Value = "Atividade"
             

   
             Worksheets("Itens_Boletim").Cells.Clear
            'FAZ UMA C�PIA DA PLANILHA "Atividade" NA PLANILHA "Itens_Boletim"
             Worksheets("Atividade").UsedRange.Copy Destination:=Worksheets("Itens_Boletim").Range("A1")
   
   calculaItensBoletim
   
             Worksheets("Boletim").Cells.Clear
            'FAZ UMA C�PIA DA PLANILHA "Itens_Boletim" NA PLANILHA "Boletim"
             Worksheets("Itens_Boletim").UsedRange.Copy Destination:=Worksheets("Boletim").Range("A1")
             
   geraBoletim
         
    'preenche planilha pendencias
    colecaoParaPlanilha colecaoPendencias, Worksheets("Pendencias")
    classificaPlanilhaSimples "Pendencias", "B1"
    Worksheets("Pendencias").Cells(1, 1).Value = "TIPO"
    Worksheets("Pendencias").Cells(1, 2).Value = "DESCRI��O"
    Worksheets("Pendencias").Cells(1, 3).Value = "NUM_OS"
    Worksheets("Pendencias").Columns.AutoFit
   
    Worksheets("Boletim").Activate
    Worksheets("Boletim").Columns.AutoFit
   
    MsgBox "PROCESSO FINALIZADO"
   
    UserFormBmd.Show
   
   'Application.ScreenUpdating = True
 
Exit Sub
TE:     'Tratamento de Erros

    MsgBox " Erro: " & Err.Description & Chr(13) & Chr(13) & "Local: M�dulo sanearCsv.sanearCsv"

End Sub
Sub eliminaOsAusente()

    Set colecaoLinhas = planilhaParaColecao(Worksheets("Unificado"))
    Set colecaoOsAusente = New Collection
    
    CONT = 1
    For Line = 1 To colecaoLinhas.Count
                   
        If Worksheets("Os_Cprog").Range("A:A").Find(colecaoLinhas(CONT)(2)) Is Nothing Then
        
            If Not existeItemColecao(colecaoOsAusente, colecaoLinhas(CONT)(2)) Then
             colecaoOsAusente.Add CStr(colecaoLinhas(CONT)(2)), CStr(colecaoLinhas(CONT)(2))
             adicionaPendencia "Erro! OS N�o ser� paga!", "SERVI�O ausente no relat�rio de OS", colecaoLinhas(CONT)(2)
            End If
           
            colecaoLinhas.Remove (CONT)
            CONT = CONT - 1
          
        End If

        CONT = CONT + 1
        
    Next

    'devolve a colecaolinhas para a planilha Unificado j� eliminado as linhas com OS ausente no relatorio de OS
    colecaoParaPlanilha colecaoLinhas, Worksheets("Unificado")
        
End Sub
Sub filtraOsPaga()

    On Error GoTo TE

    Set colecaoJaPagas = New Collection
    
    caminhoArquivo = Application.ActiveWorkbook.Path & "\registroLancamentosBmdsBD.csv"

    
    Set colecaoLinhas = planilhaParaColecao(Worksheets("Car"))
    
    
    contador = 1
    For Each elemento In colecaoLinhas
        'procura na coluna 3 do arquivo de registro de lancamentos correlacoes da coluna 5(num_os) da colecaodelinhas
        If termoExisteNoArquivo(caminhoArquivo, 3, elemento(5)) Then
          
           If Not existeItemColecao(colecaoJaPagas, CStr(elemento(5))) Then
            colecaoJaPagas.Add elemento(5), CStr(elemento(5))
           End If
           
          colecaoLinhas.Remove (contador)
          contador = contador - 1
        End If
      
        contador = contador + 1

     Next

    colecaoParaPlanilha colecaoLinhas, Worksheets("Car")

    adicionaPendencia "Informa��o!", "Qtd de OS j� pagas e ignoradas: " & colecaoJaPagas.Count, ""
    
Exit Sub
TE:     'Tratamento de Erros

    MsgBox " Erro: " & Err.Description & Chr(13) & Chr(13) & "Local: M�dulo sanearCsv.filtraOsPaga"

End Sub
Sub extraiEmBaixa()

    'limpa planilha Pendencias
    Worksheets("Pendencias").Cells.Clear

    'AUTOFILTRA PLANILHA UNIFICADO PELO TERMO EM BAIXA E COPIA PARA PLANILHA Pendencias
    Worksheets("Unificado").Range("A1").AutoFilter Field:=2, Criteria1:="EM BAIXA"
    Worksheets("Unificado").Range("A1").AutoFilter Field:=14, Criteria1:=mesAno
    Worksheets("Unificado").AutoFilter.Range.Copy Destination:=Worksheets("Pendencias").Range("A1")
    Worksheets("Unificado").AutoFilterMode = False

    Worksheets("Pendencias").Activate
    'remove linhas duplicadas
    Worksheets("Pendencias").Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(5), Header:=xlYes
    
    qtdLinhas = Worksheets("Pendencias").Range("A1").CurrentRegion.Rows.Count
    
    
    arraycolunas = Worksheets("Pendencias").Range("E1:E" & qtdLinhas)
    
    'limpa planilha Pendencias
    Worksheets("Pendencias").Cells.Clear
 
    For linha = 2 To qtdLinhas
        adicionaPendencia "Erro! OS N�o ser� paga!", "A OS n�o est� fechada(Em Baixa).", arraycolunas(linha, 1)
    Next

     
End Sub
Sub extraiNovas()

    'limpa planilha Pendencias
    Worksheets("Pendencias").Cells.Clear

    'AUTOFILTRA PLANILHA UNIFICADO PELO TERMO EM NOVA E COPIA PARA PLANILHA Pendencias
    Worksheets("Unificado").Range("A1").AutoFilter Field:=2, Criteria1:="NOVA"
    Worksheets("Unificado").Range("A1").AutoFilter Field:=14, Criteria1:=mesAno
    Worksheets("Unificado").AutoFilter.Range.Copy Destination:=Worksheets("Pendencias").Range("A1")
    Worksheets("Unificado").AutoFilterMode = False

    Worksheets("Pendencias").Activate
    
    'remove linhas duplicadas
    Worksheets("Pendencias").Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(5), Header:=xlYes
    
    qtdLinhas = Worksheets("Pendencias").Range("A1").CurrentRegion.Rows.Count
    
    arraycolunas = Worksheets("Pendencias").Range("E1:E" & qtdLinhas)
    
    'limpa planilha Pendencias
    Worksheets("Pendencias").Cells.Clear
    Worksheets("Pendencias").Cells(1, 1).Value = "PENDENCIAS"
 
    For linha = 2 To qtdLinhas
        adicionaPendencia "Erro! OS N�o ser� paga!", "A OS n�o est� fechada(Nova).", arraycolunas(linha, 1)
    Next

     
End Sub
Sub adicionaPendencia(ByVal texto1 As String, ByVal texto2 As String, ByVal texto3 As String)
    ReDim arrayPendencia(1 To 3) As String
    arrayPendencia(1) = texto1
    arrayPendencia(2) = texto2
    arrayPendencia(3) = texto3
    colecaoPendencias.Add arrayPendencia
End Sub

Sub formataColunasInteresse()


  Worksheets("Col_Interesse").Activate
  
  qtdLinhas = Worksheets("Col_Interesse").Range("A1").CurrentRegion.Rows.Count
  
  ' array contendo a nova ordem para as colunas
  'ordem original     1  2  3  4  5 6 7 8  9 101112131415 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36
  novaOrdemColunas = "3 23 26 27 25 5 6 1 29 14 4 7 2 8 9 10 11 12 13 15 16 17 18 19 20 21 22 24 28 30 31 32 33 34 35 36"
 
  
  '� passado um conjunto de linhas e um array de colunas como parametro na func�o index que retorna um array de colunas
  arraycolunas = Application.Index(Cells, Evaluate("Row(1:" & qtdLinhas & ")"), Split(novaOrdemColunas))
  'limpa planilha col_interesse
  Worksheets("Col_Interesse").Cells.Clear
  'Worksheets("Col_Interesse").Range("A1").CurrentRegion.Columns(1).NumberFormat = "@"
  Worksheets("Col_Interesse").Range("A1").CurrentRegion.Columns(1).NumberFormat = "[$-416]mmm-yy;@"
  
  Worksheets("Col_Interesse").Range("A1").CurrentRegion.Columns(10).NumberFormat = "@"
  'cola o array de colunas na planilha col_interesse
  Worksheets("Col_Interesse").Range(Cells(1, 1), Cells(qtdLinhas, UBound(Split(novaOrdemColunas)) + 1)).Value = arraycolunas

 'unificando as colunas pep e ordem interna
 Set colecaoLinhas = planilhaParaColecao(Worksheets("Col_Interesse"))
 'percorrendo todos os elementos da colecao de linhas
 For Line = 1 To colecaoLinhas.Count

    'array auxiliar criado porque  o objeto collections n�o permite alterar o valor dos itens, apenas remove-los ou adicion�-los
     arrayAuxiliar = colecaoLinhas(Line)

     If arrayAuxiliar(3) = "" Then
      arrayAuxiliar(3) = arrayAuxiliar(4)
      colecaoLinhas.Add arrayAuxiliar, after:=Line
      colecaoLinhas.Remove (Line)
     End If
     
 Next

 colecaoParaPlanilha colecaoLinhas, Worksheets("Col_Interesse")

 Worksheets("Col_Interesse").Columns(4).Delete
 Worksheets("Col_Interesse").Cells(1, 3).Value = "PEP/Ordem Interna"

 
 Worksheets("Col_Interesse").Range("A1").CurrentRegion.Columns(34).NumberFormat = "dd/mm/yyyy"
 Worksheets("Col_Interesse").Range("A1").CurrentRegion.Columns(35).NumberFormat = "dd/mm/yyyy"

 
     
    
    'DETERMINA O TAMANHO DAS COLUNAS"
    Worksheets("Col_Interesse").Columns.AutoFit
    
End Sub
Sub escrevePendencias()

If Not planilhaVazia(Worksheets("Pendencias")) Then

    qtdLinhas = Worksheets("Pendencias").Range("A1").CurrentRegion.Rows.Count
    
    With UserFormBmd.ListBox1
        .ColumnCount = 3
        .ColumnWidths = "150;220;10"
        .ColumnHeads = True
        .RowSource = "Pendencias!A2:C" & qtdLinhas
        
    End With
      
Else
    
    UserFormBmd.ListBox1.AddItem "NENHUMA PENDENCIA."
    
End If
     
End Sub
Function planilhaVazia(ByVal ws As Worksheet)
    planilhaVazia = False
    If ws.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Rows.Count = 1 Then
        planilhaVazia = True
    End If
End Function
Sub geraBoletim()

 Worksheets("Boletim").Activate

 Worksheets("Boletim").Rows(1).HorizontalAlignment = xlCenter
  
 Worksheets("Boletim").Columns(7).Cut
 Worksheets("Boletim").Columns(12).Insert

 'Worksheets("Boletim").Columns("J:AH").Hidden = True
 Worksheets("Boletim").Columns("K:AM").Delete
 
 Worksheets("Boletim").Columns(6).Cut
 Worksheets("Boletim").Columns(3).Insert
 
 'cria colunas num_bmd, sequencia, num_pedido e situacao_pep
  Worksheets("Boletim").Cells(1, 11).Value = "Num_Bmd"
  Worksheets("Boletim").Cells(1, 12).Value = "Sequ�ncia"
  Worksheets("Boletim").Cells(1, 13).Value = "Num_Pedido"
  Worksheets("Boletim").Cells(1, 14).Value = "Situacao_Pep"
  
  Worksheets("Boletim").Range("A1").CurrentRegion.Columns(13).NumberFormat = "0"
  
  classificaPlanilhaSimples "Boletim", "B1"
 
 
 Worksheets("Boletim").Columns.AutoFit

End Sub
Sub moveColunaFinal(ByVal ws As Worksheet, ByVal coluna As Integer)

 ' sempre mover de tras pra frente pois as colunas s�o movimentadas para
 ' o final modificando as posicoes das outras colunas tb.

 numero_colunas = ws.Range("A1").CurrentRegion.Columns.Count

'movendo as colunas que serao ocultas para o final pois currenteregion nao pega contiguamente os visiveis
 ws.Columns(coluna).Cut
 ws.Columns(numero_colunas + 1).Insert
 
 'ws.Columns(numero_colunas).Hidden = True
 
End Sub
Sub calculaItensBoletim()

    Set colecaoAtv = colecaoDeAtividades
    
    Worksheets("Itens_Boletim").Activate
 
    Columns(5).Insert
    Range("E1").Value = "Descri��o"
    
    Columns(9).Insert
    Range("I1").Value = "Total"
    Worksheets("Itens_Boletim").Range("A1").CurrentRegion.Columns(9).NumberFormat = "#,##0.00"
    Columns(9).Insert
    Range("I1").Value = "Valor_US"
    Worksheets("Itens_Boletim").Range("A1").CurrentRegion.Columns(9).NumberFormat = "#,##0.00"
    Columns(9).Insert
    Range("I1").Value = "Qtd_US"
    Worksheets("Itens_Boletim").Range("A1").CurrentRegion.Columns(9).NumberFormat = "#,##0.00"
    
   
    
    Set colecaoLinhas = planilhaParaColecao(Worksheets("Itens_Boletim"))
    
    contador = 1
    For Each elemento In colecaoLinhas
         
           'array auxiliar criado porque  o objeto collections n�o permite alterar o valor dos itens, apenas remove-los ou adicion�-los
            arrayAuxiliar = elemento
           
           ' valores relativos as US s�o substituidos no arrayauxiliar, cstr converte para string
           arrayAuxiliar(5) = colecaoAtv(CStr(elemento(4))).descricao
           arrayAuxiliar(9) = colecaoAtv(CStr(elemento(4))).qtd_us
           arrayAuxiliar(10) = colecaoAtv(CStr(elemento(4))).valor_us
           If colecaoAtv(CStr(elemento(4))).codigo = 1317 Then
            arrayAuxiliar(11) = arrayAuxiliar(8) * arrayAuxiliar(9) * arrayAuxiliar(10) * 0.022
           Else
            arrayAuxiliar(11) = arrayAuxiliar(8) * arrayAuxiliar(9) * arrayAuxiliar(10)
           End If
           
                                     
           colecaoLinhas.Add arrayAuxiliar, after:=contador
           colecaoLinhas.Remove (contador)
           contador = contador + 1

     Next
    
    colecaoParaPlanilha colecaoLinhas, Worksheets("Itens_Boletim")
    
  
End Sub
Function planilhaParaColecao(ByVal ws As Worksheet)

    'Transporta os valores de todas as celulas da planilha para um array, exceto colunas ocultas
     arrayPlanilhaInteira = ws.Range("A1").CurrentRegion
     
     'Foi criado uma cole��o de linhas para poder remover linhas depois
     Set planilhaParaColecao = New Collection
    
     'upper bound nos tras o ultimo elemento do array
     numeroColunas = UBound(arrayPlanilhaInteira, 2)
     numeroLinhas = UBound(arrayPlanilhaInteira, 1)
    
     'criando array contendo uma linha do arrayPlanilhaInteira e adicionando na cole��o de linhas
     For linha = 2 To numeroLinhas
    
        ReDim arrayLinha(1 To numeroColunas) As Variant
    
        For Col = 1 To numeroColunas
            'previne celulas com erro de formulas #NUM!
            If IsError(arrayPlanilhaInteira(linha, Col)) = False Then
             arrayLinha(Col) = arrayPlanilhaInteira(linha, Col)
            End If
    
        Next
    
        planilhaParaColecao.Add (arrayLinha)
    
     Next
     
End Function
Sub colecaoParaPlanilha(ByVal colecao As Collection, ByVal ws As Worksheet)

 ws.Activate
 
 'Apaga todas as celulas da planilha
 With ws
 .Range(Cells(2, 1), Cells(.Range("A1").CurrentRegion.Rows.Count + 1, .Range("A1").CurrentRegion.Columns.Count)).ClearContents
 End With

 'caso haja linhas na cole��o de linhas tranporta seus valores para um array bidimensional arrayResultado
 If colecao.Count > 0 Then

     ReDim arrayResultado(1 To colecao.Count, 1 To UBound(colecao(1))) As Variant

     For linha = 1 To UBound(arrayResultado, 1)

        For Col = 1 To UBound(arrayResultado, 2)
            arrayResultado(linha, Col) = colecao(linha)(Col)
        Next

     Next
     
     ws.Range(Cells(2, 1), Cells(colecao.Count + 1, UBound(colecao(1)))).Value = arrayResultado
     ws.Columns.AutoFit

 End If
End Sub
Function colecaoDeAtividades()
 

    Worksheets("Configura��es").Activate
    
    Worksheets("Configura��es").Unprotect
    
    numeroDeAtividades = Worksheets("Configura��es").Range("E13").CurrentRegion.Rows.Count - 6
    
     Worksheets("Configura��es").Protect
     
    'Transporta os valores de todas as celulas da planilha para um array
     arrayAtividades = Worksheets("Configura��es").Range(Cells(13, 5), Cells(numeroDeAtividades + 12, 8))
    
    'Foi criado uma cole��o com mesmo nome da function assim eh retornado a colecao
    Set colecaoDeAtividades = New Collection

    'upper bound nos tras o ultimo elemento do array
    numeroColunas = UBound(arrayAtividades, 2)
    numeroLinhas = UBound(arrayAtividades, 1)
       
     'criando array contendo uma linha do arrayPlanilhaInteira e adicionando na cole��o de linhas
     For linha = 1 To numeroLinhas
       
        Set atv = New atividade
        
        atv.codigo = arrayAtividades(linha, 1)
        atv.descricao = arrayAtividades(linha, 2)
        atv.qtd_us = arrayAtividades(linha, 3)
        atv.valor_us = arrayAtividades(linha, 4)
        
        colecaoDeAtividades.Add atv, atv.codigo
    
     Next

End Function
Function existeItemColecao(ByVal colecao As Collection, ByVal termo As String)

    On Error GoTo Erro
    
        existeItemColecao = True
        colecao (termo)
        Exit Function
Erro:
      existeItemColecao = False
      
End Function
Sub unificaDeslocamentoAtividade()

    'ABRE A PLANILHA "Deslocamento"
     Worksheets("Deslocamento").Activate
     'FAZ UMA C�PIA DA PLANILHA "Deslocamento" NA PLANILHA "Total_Atividades" na ultima celula
     If Worksheets("Deslocamento").Range("A1").CurrentRegion.Rows.Count > 1 Then
        With Worksheets("Deslocamento")
          .Range(Cells(2, 1), Cells(.Range("A1").CurrentRegion.Rows.Count, .Range("A1").CurrentRegion.Columns.Count)).Copy _
             Destination:=Worksheets("Total_Atividades").Range("A" & Worksheets("Total_Atividades").Range("A1").CurrentRegion.Rows.Count + 1)
        End With
     End If

End Sub
Sub unificaServicosOs()
    
        
    Worksheets("Unificado").Activate
    
    Columns(1).Insert
    Range("A1").Value = "M�s/Ano_Fechamento_Os"
    Columns(1).Insert
    Range("A1").Value = "Estado"
    Columns(1).Insert
    Range("A1").Value = "Deslocamento"
    Columns(35).Insert
    
    'Range("N1").Value = "M�s/Ano_Execucao_Os"
   qtdLinhas = Worksheets("Unificado").Range("A1").CurrentRegion.Rows.Count
    
  'para cada os na planilha servicos procura na planilha os a sua situa��o e cola na planilha servicos na coluna situa��o
  ReDim arrayResultado(qtdLinhas - 1)
  indice = 0
  For Line = 2 To qtdLinhas
     
     arrayResultado(indice) = Application.VLookup(Worksheets("Unificado").Cells(Line, 5).Value, Worksheets("Os_Cprog").Range("A:C"), 3, False)
     indice = indice + 1
     
  Next
 
  Worksheets("Unificado").Range(Cells(2, 2), Cells(qtdLinhas, 2)).Value = Application.Transpose(arrayResultado)
  
  'para cada os na planilha servicos procura na planilha os a qtd de km e cola na planilha servicos na coluna deslocamento
  ReDim arrayResultado(qtdLinhas - 1)
  indice = 0
  For Line = 2 To qtdLinhas
     
     arrayResultado(indice) = Application.VLookup(Worksheets("Unificado").Cells(Line, 5).Value, Worksheets("Os_Cprog").Range("A:T"), 20, False)
     indice = indice + 1
     
  Next
 
  Worksheets("Unificado").Range(Cells(2, 1), Cells(qtdLinhas, 1)).Value = Application.Transpose(arrayResultado)
  
  'para cada os na planilha servicos procura na planilha os a data de fechamento da os e cola na planilha servicos na coluna mes/ano de execu��o
  ReDim arrayResultado(qtdLinhas - 1)
  indice = 0
  For Line = 2 To qtdLinhas
     
     arrayResultado(indice) = Application.VLookup(Worksheets("Unificado").Cells(Line, 5).Value, Worksheets("Os_Cprog").Range("A:N"), 14, False)
     indice = indice + 1
     
  Next
  

  Worksheets("Unificado").Range("A1").CurrentRegion.Columns(14).NumberFormat = "0"
  Worksheets("Unificado").Range(Cells(2, 3), Cells(qtdLinhas, 3)).Value = Application.Transpose(arrayResultado)
  
  'faz copia da coluna C (mes/ano_fechamento) na coluna AI(data_fechamento)
  Worksheets("Unificado").Range("A1").CurrentRegion.Columns(35).NumberFormat = "dd/mm/yyyy"
  arraycolunas = Worksheets("Unificado").Range("C1:C" & qtdLinhas)
  Worksheets("Unificado").Range(Cells(1, 35), Cells(qtdLinhas, 35)).Value = Application.Transpose(Application.Transpose(arraycolunas))
  Range("AI1").Value = "Data_Fechamento_Os"
  
  'faz copia da coluna N (mes/ano_execucao) na coluna AJ(data_fechamento)
  Worksheets("Unificado").Range("A1").CurrentRegion.Columns(36).NumberFormat = "dd/mm/yyyy"
  arraycolunas = Worksheets("Unificado").Range("N1:N" & qtdLinhas)
  Worksheets("Unificado").Range(Cells(1, 36), Cells(qtdLinhas, 36)).Value = Application.Transpose(Application.Transpose(arraycolunas))
  Range("AJ1").Value = "Data_Execucao_Os"
  
  
  Worksheets("Unificado").Cells(1, 14).Value = "Mes/Ano_Execucao"

End Sub
Sub formataComoMes(ByVal ws As Worksheet, ByVal coluna As Integer)

  Dim UltimaLinha As Long
 
  
  ws.Range("A1").CurrentRegion.Columns(coluna).NumberFormat = "[$-416]mmm-yy;@"
    
  UltimaLinha = ws.Cells(Cells.Rows.Count, 1).End(xlUp).Row
    
  ws.Activate
    
  'COPIA COLUNA 1 PARA UM ARRAY -> MUDA NUMBERFORMAT PARA "@" (TEXTO) -> DEVOLVE O ARRAY PARA A COLUNA 1.FOI NECESSARIO ESTE ARTIFICIO POIS A COLUNA NAO ESTAVA SENDO CLASSIFICADA PELA NOVA FORMATACAO POR MES E SIM POR DATA
  qtdLinhas = ws.Range("A1").CurrentRegion.Rows.Count
  ReDim arrayResultado(qtdLinhas - 1)
  indice = 0
  For Line = 2 To qtdLinhas
     
     arrayResultado(indice) = Cells(Line, coluna).Text
     indice = indice + 1
     
  Next
  
  ws.Range("A1").CurrentRegion.Columns(coluna).NumberFormat = "@"
 
  ws.Range(Cells(2, coluna), Cells(qtdLinhas, coluna)).Value = Application.Transpose(arrayResultado)

  'Cells(1, coluna).Value = "M�s/Ano"

End Sub
Sub copiaColuna(ByVal cabecalho As String, _
                     ByVal planilhaOrigem As String, _
                     ByVal planilhaDestino As String, _
                     ByVal celulaDestino As String)

    
    For Each Cell In Worksheets("Unificado").Range("A1:AG1")
    
        If Cell.Value = cabecalho Then
                  
           Cell.EntireColumn.Copy Worksheets(planilhaDestino).Range(celulaDestino)
           Exit For
          
        End If
        
    Next
    
End Sub

Sub classificaPlanilha(ByVal planilha As String)
     With Worksheets(planilha).Sort
        
            .SortFields.Clear
            .SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range("B1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range("C1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range("D1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Worksheets(planilha).Range("A1").CurrentRegion
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            
        End With
End Sub

Sub classificaPlanilhaSimples(ByVal planilha As String, ByVal coluna As String)
     With Worksheets(planilha).Sort
        
            .SortFields.Clear
            .SortFields.Add Key:=Range(coluna), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Worksheets(planilha).Range("A1").CurrentRegion
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            
        End With
End Sub
Sub converteCidade(ByVal coluna As Integer)

 Set colecaoCarInconsistente = New Collection
 Set colecaoLinhas = planilhaParaColecao(Worksheets("Car"))

 Worksheets("Car").Activate
 qtdLinhas = Worksheets("Car").Range("A1").CurrentRegion.Rows.Count
 
 contColecaoLinhas = 1
 For contador = 1 To colecaoLinhas.Count
 
    arrayAuxiliar = colecaoLinhas(contColecaoLinhas)
  
    
    If arrayAuxiliar(2) Like "*OFC_EMEC*" Then
        arrayAuxiliar(2) = Worksheets("Configura��es").Range("C9").Value
    Else
    
       car = esplitaFrase(arrayAuxiliar(2))
       resultadoConsulta = consultaCidadeDicionario(car)
            
       arrayAuxiliar(2) = resultadoConsulta
     
    End If
    
    If arrayAuxiliar(2) = "" Then
                   
        If Not existeItemColecao(colecaoCarInconsistente, arrayAuxiliar(5)) Then
            colecaoCarInconsistente.Add CStr(arrayAuxiliar(5)), CStr(arrayAuxiliar(5))
            adicionaPendencia "Erro! OS N�o ser� paga!", "CAR n�o encontrado.", arrayAuxiliar(5)
        End If
                    
        colecaoLinhas.Remove (contColecaoLinhas)
        contColecaoLinhas = contColecaoLinhas - 1
           
        
    Else
    
        colecaoLinhas.Add arrayAuxiliar, after:=contColecaoLinhas
        colecaoLinhas.Remove (contColecaoLinhas)
    
    End If
    
    contColecaoLinhas = contColecaoLinhas + 1
 Next
 
 colecaoParaPlanilha colecaoLinhas, Worksheets("Car")
 
 Worksheets("Car").Columns.AutoFit
  
End Sub
Function esplitaFrase(ByVal frase As String)
 
 Dim resultado() As String
 
 resultado = Split(frase, "/")
 
 tamanho = tamanhoArray(resultado)
 
    If tamanho > 4 Then
     If IsNumeric(resultado(4)) Then
      esplitaFrase = resultado(4)
     Else
      esplitaFrase = "SEM CAR"
     End If
    Else
     esplitaFrase = "SEM CAR"
    End If
 

End Function
Sub extraiAtividade(ByVal coluna As Integer)

'Esta Sub rotina:

'   - remove as linhas duplicadas da planilha car considerando as colunas num os e sequencial de servico

'   - Isola, no campo coment�rio, os termos entre asteriscos se houver(alerta)
'   - remove espa�os em branco e verifica se o termo � um n�mero

'   - verifica se o numero corresponde � um c�digo de atividade v�lido conforme planilha configura��es(alerta)
'   - caso haja v�rias atividades no comentario s�o inseridas novas linhas para cada atividade encontrada.
'   - substitui o coment�rio pela atividade encontrada.
'   - caso atividade seja 1317 calcula numero de participantes e multiplica pelo km
'   - verifica se existe atividades repetidas na mesma os (erro)
'   - verifica se existe atividades do mesmo grupo(soda, estudo de protecao e atvDeslocamento. ex: 1500 e 1501) repetidas na mesma os (erro)
'   - verifica 1317 sem atividade que provoca deslocamento na mesma os (alerta) *
'   - verifica atividade que provoca deslocamento sem 1317 na mesma os (alerta, estacao de chaves(varios equipamentos sem deslocamento)) *
'   - verifica na mesma OS km acima do limite configurado na planilha configura��es (alerta) *
'   - 1317 com numero de participantes acima do limite configurado na planilha de configura��es (alerta) *

 Dim frase As String
 Dim esplitado() As String
 Dim numerico As Collection


 Set colecaoChavesOs = New Collection
 Set colecaoAtividades = New Collection
 Set colecaoRepetidos = New Collection
 Set colecaoRepetidosOs = New Collection
 Set colecaoChavesAtividade = New Collection
 
 Set colecaoAtv = colecaoDeAtividades
  
 Worksheets("Atividade").Activate
 Worksheets("Atividade").Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(5, 6), Header:=xlYes
 Worksheets("Atividade").UsedRange.Columns(1).NumberFormat = "@"
  
 Set colecaoLinhas = planilhaParaColecao(Worksheets("Atividade"))
 
 qtdLinhas = Worksheets("Car").Range("A1").CurrentRegion.Rows.Count
 
 limiteParticipantes = Worksheets("Configura��es").Range("C11").Value
 limiteKmPorOs = Worksheets("Configura��es").Range("C13").Value
 
 'array de Atividades que provocam deslocamento
 atividadesDeslocamento = Application.Transpose(Worksheets("Configura��es").Range("J12:J206"))
 'array de Atividades de soda
 atividadessoda = Application.Transpose(Worksheets("Configura��es").Range("L12:L206"))
 'array de Atividades de estudo de prote��o
 atividadesEstudoProtecao = Application.Transpose(Worksheets("Configura��es").Range("N12:N206"))
 
 'indexador paralelo criado para poder lidar com a remo��o de linhas que altera o tamanho do objeto que se esta iterando dinamicamente
 paralelo = 1
 'percorrendo todos os elementos da colecao de linhas
 For Line = 2 To colecaoLinhas.Count + 1
 
    'array auxiliar criado porque  o objeto collections n�o permite alterar o valor dos itens, apenas remove-los ou adicion�-los
    arrayAuxiliar = colecaoLinhas(paralelo)
    
    'calcula numero de participantes para cada sequencia de servico
    qtd_participantes = WorksheetFunction.CountIfs(Worksheets("Car").Range("E2:E" & qtdLinhas), arrayAuxiliar(5), _
                            Worksheets("Car").Range("F2:F" & qtdLinhas), arrayAuxiliar(6))
                            
    'lan�a alerta para limite de participantes configurado ultrapassado
    If qtd_participantes > limiteParticipantes Then adicionaPendencia "Alerta!", "N�m. participantes > LIMITE(" & arrayAuxiliar(6) & ")", arrayAuxiliar(5)
                            
    'armazena a qtd de km em uma variavel
    km = arrayAuxiliar(7)
    
    'lan�a alerta para limite de km configurado ultrapassado
    If km > limiteKmPorOs Then adicionaPendencia "Alerta!", "Qtd. KM > LIMITE(" & arrayAuxiliar(6) & ")", arrayAuxiliar(5)
    
    'cole��o numerico criada para verificar se termo entre * � um n�mero
    Set numerico = New Collection
    
    esplitado = Split(colecaoLinhas(paralelo)(coluna), "*")

    tamanhoEsplitado = tamanhoArray(esplitado)
    
    'se houver termos entre asteriscos
    If tamanhoEsplitado > 0 Then
        'verifica se � numerico
        contador = 0
        For Each elemento In esplitado
         'extremidades sao tratadas diferente pois nao estao entre asteriscos
         If contador = 0 Or contador = UBound(esplitado) Then
            If IsNumeric(Trim(elemento)) Then
               adicionaPendencia "Alerta!", "Num�rico sem asterisco(" & Trim(elemento) & ")", colecaoLinhas(paralelo)(5)
            End If
         Else
            If IsNumeric(Trim(elemento)) Then
              numerico.Add Trim(elemento)
            Else
             If Len(Trim(elemento)) < 7 And Len(Trim(elemento)) > 1 Then
                 adicionaPendencia "Alerta!", "C�digo de atividade inv�lido(" & Trim(elemento) & ")", colecaoLinhas(paralelo)(5)
             End If
            End If
          End If
            contador = contador + 1
        Next elemento

        'se houver termos entre astericos que sejam numericos
        If numerico.Count > 0 Then
            
            For CONT = 1 To numerico.Count
            
                If CONT = 1 Then
                   colecaoLinhas.Remove (paralelo)
                   paralelo = paralelo - 1
                End If
                    
                 
                ' se o numero corresponde a um codigo de atividade valido
                If existeItemColecao(colecaoAtv, numerico(CONT)) Then
                    'atribui o numero da atividade ao valor da celula correspondente no array
                    arrayAuxiliar(coluna) = numerico(CONT)
                    
                    'calcula numero de participantes e multiplica pelo km caso seja atividade 1317
                    If numerico(CONT) = "1317" Then
                        arrayAuxiliar(7) = km * qtd_participantes
                    Else
                        arrayAuxiliar(7) = 1
                    End If
                    
                    If paralelo = 0 Then
                     If colecaoLinhas.Count = 0 Then
                        colecaoLinhas.Add arrayAuxiliar
                     Else
                        colecaoLinhas.Add arrayAuxiliar, before:=1
                     End If
                    Else
                     colecaoLinhas.Add arrayAuxiliar, after:=paralelo
                    End If
                    paralelo = paralelo + 1
                    
                Else
                    adicionaPendencia "Alerta!", "C�digo de atividade inv�lido(" & numerico(CONT) & ")", arrayAuxiliar(5)
                End If
                
                Next
         Else
            adicionaPendencia "Alerta!", "Sequencial sem atividade(" & colecaoLinhas(paralelo)(6) & ")", colecaoLinhas(paralelo)(5)
            colecaoLinhas.Remove (paralelo)
            paralelo = paralelo - 1
         End If
    Else
       adicionaPendencia "Alerta!", "Sequencial sem atividade(" & colecaoLinhas(paralelo)(6) & ")", colecaoLinhas(paralelo)(5)
       'remove a linha caso n�o haja termos entre asteriscos
       colecaoLinhas.Remove (paralelo)
       paralelo = paralelo - 1
    End If
   paralelo = paralelo + 1
 Next
 
        
 
 ' TERCEIRA PARTE: VERIFICANDO EXISTENCIA NA MESMA OS DE ATIVIDADES REPETIDAS

         'determinando as OS chaves
         For Each elemento In colecaoLinhas
               If Not existeItemColecao(colecaoChavesOs, elemento(5)) Then
                colecaoChavesOs.Add CStr(elemento(5)), CStr(elemento(5))
               End If
         Next

        'para cada os gera uma colecao de atividades para aquela os
         For Each chave In colecaoChavesOs
               Set colecaoAtividades = New Collection
               'preenchendo colecao de atividades para uma os
               For Each linha In colecaoLinhas
                If CStr(linha(5)) = chave Then
                     colecaoAtividades.Add linha(4)
                End If
               Next
               
                
                 num1317 = 0
                 numAtvDeslocamento = 0
                 numAtvsoda = 0
                 numAtvEstudoProtecao = 0
                 
                 
                 For Each elemento In colecaoAtividades
                       
                       ' determina se existe atividade 1317
                       If elemento = "1317" Then num1317 = num1317 + 1
                       ' determina qtd de atividades que provocam deslocamento
                       If estaNoArray(elemento, atividadesDeslocamento) Then numAtvDeslocamento = numAtvDeslocamento + 1
                       ' determina qtd de atividades de soda
                       If estaNoArray(elemento, atividadessoda) Then numAtvsoda = numAtvsoda + 1
                       ' determina qtd de atividades de estudo de prote��o
                       If estaNoArray(elemento, atividadesEstudoProtecao) Then numAtvEstudoProtecao = numAtvEstudoProtecao + 1
                       

                       'determinando as atividades chaves
                       If Not existeItemColecao(colecaoChavesAtividade, elemento) Then
                        colecaoChavesAtividade.Add elemento, elemento
                       End If

                 Next
                 
                 If num1317 > 0 And numAtvDeslocamento = 0 Then adicionaPendencia "Alerta!", "1317 sem atv que provoca deslocamento.", chave
                 If num1317 = 0 And numAtvDeslocamento > 0 Then adicionaPendencia "Alerta!", "Atv. que provoca deslocamento sem 1317.", chave
                 
                 'marca os com atividades de mesmo grupo em duplicidade para serem removidas
                 If numAtvDeslocamento > 1 Or numAtvsoda > 1 Or numAtvEstudoProtecao > 1 Then
                  colecaoRepetidosOs.Add chave, chave
                 End If
                 
                 'cria alerta de erro para atividades de mesmo grupo em duplicidade na mesma os
                 If numAtvDeslocamento > 1 Then adicionaPendencia "Erro! OS N�o ser� paga!", "Atividade de deslocamento em duplicidade.", chave
                 If numAtvsoda > 1 Then adicionaPendencia "Erro! OS N�o ser� paga!", "Atividade de soda em duplicidade.", chave
                 If numAtvEstudoProtecao > 1 Then adicionaPendencia "Erro! OS N�o ser� paga!", "Atv. de estudo de prote��o em duplicidade.", chave
                 
               'verificando elementos repetidos na colecao de atividades para uma mesma os e guardando o num_os para posterior elimina��o
               For Each elemento In colecaoChavesAtividade
                    contador = 0
                    For Each e In colecaoAtividades
                        If elemento = e Then
                            If contador > 0 Then
                                ReDim arrayRepetidos(2) As String
                                arrayRepetidos(0) = chave
                                arrayRepetidos(1) = elemento
                                colecaoRepetidos.Add arrayRepetidos
                                If Not existeItemColecao(colecaoRepetidosOs, chave) Then
                                    colecaoRepetidosOs.Add chave, chave
                                End If
                            End If
                            contador = contador + 1
                        End If

                    Next

               Next

           Next

        
         'Eliminando as OS que cont�m atividades repetidas da colecao de linhas(estas OS n�o serao pagas)
         i = 1
         For Each linhaColecao In colecaoLinhas
         
             If existeItemColecao(colecaoRepetidosOs, linhaColecao(5)) Then
                colecaoLinhas.Remove i
                i = i - 1
             End If
                        
             i = i + 1

         Next
         
     'imprimindo na planilha pendencias as atividades repetidas e respectivos numeros de os
     For Each elemento In colecaoRepetidos
        adicionaPendencia "Erro! OS N�o ser� paga!", "Atividade em duplicidade(" & elemento(1) & ")", elemento(0)
     Next
 
     Worksheets("Pendencias").Columns.AutoFit
 
 
    'formata as novas linha como texto na coluna 1 "jan-22"
     Worksheets("Atividade").Range(Cells(2, 1), Cells(colecaoLinhas.Count + 1, 1)).NumberFormat = "@"
 
     colecaoParaPlanilha colecaoLinhas, Worksheets("Atividade")
     
     Worksheets("Atividade").Range("G1").Value = "Quantidade"
  
End Sub
Function estaNoArray(ByVal valorProcurado As String, ByVal vetor As Variant) As Boolean
  estaNoArray = (UBound(Filter(vetor, valorProcurado)) > -1)
End Function

Function tamanhoArray(vetor As Variant) As Long
   If IsEmpty(vetor) Then
      tamanhoArray = 0
   Else
      tamanhoArray = UBound(vetor) - LBound(vetor) + 1
   End If
End Function

