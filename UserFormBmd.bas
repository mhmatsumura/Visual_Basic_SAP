
Dim colecaoPendenciasPep As Collection
Dim session As Variant
Dim User As Variant
Dim contColecaoLinhas As Variant
Dim itemSODA As Variant
Dim itemComissionamento As Variant
Dim itemAutAtivacao As Variant
Dim itemEstprotecao As Variant
Dim contItensPedido As Variant
Dim atividadesDeslocamento As Variant
Dim atividadesSODA As Variant
Dim atividadesEstudoProtecao As Variant
Dim colecaoLinhas As Variant
   

Private Sub btLancarSap_Click()

    On Error GoTo Erro

    UserFormBmd.Hide

    Worksheets("Boletim").Activate
    
    If planilhaVazia(Worksheets("Boletim")) Then

        Err.Raise ERRO_PLANILHA_VAZIA, "Planilha vazia!" _
                            , "Nenhum ITEM na planilha BOLETIM."

    End If

    qtdLinhas = Worksheets("Boletim").Range("A1").CurrentRegion.Rows.Count

    'verifica celulas vazias
    For Each celula In Worksheets("Boletim").Range("D2:D" & qtdLinhas)
        If IsEmpty(celula) Then
            Err.Raise vbObjectError + 50, "DADO INV�LIDO!" _
                            , "Alguma c�lula est� vazia !! Verifique coluna: STATUS_PEP."
        End If
    Next
    
    'verifica celulas vazias ou n�o numericas
    For Each celula In Worksheets("Boletim").Range("K2:L" & qtdLinhas)
        If IsEmpty(celula) Or Not IsNumeric(celula) Then
            Err.Raise vbObjectError + 50, "DADO INV�LIDO!" _
                            , "Alguma c�lula est� vazia ou contendo valor n�o num�rico!! Verifique colunas: NUM_BMD, E SEQU�NCIA."
        End If
    Next
    
    'verifica celulas vazias ou n�o numericas
    For Each celula In Worksheets("Boletim").Range("C2:C" & qtdLinhas)
        If IsEmpty(celula) Or Not IsNumeric(celula) Then
            Err.Raise vbObjectError + 50, "DADO INV�LIDO!" _
                            , "Alguma c�lula est� vazia ou contendo valor n�o num�rico!! Verifique coluna: NUM_OS."
        End If
    Next

 
 Set session = abreSessao(nomeSistemaDesejado:="GCE", nomeConexaoDesejada:="GCE PUBLIC")
 
 session.findById("wnd[0]").Maximize
    
    Dim RE As Object
    Set RE = CreateObject("vbscript.regexp")
        With RE
            .MultiLine = False
            .Global = False
            .IgnoreCase = True
            .Pattern = "\d{8}"
        End With
        
    'retorna o foco para o excel
    AppActivate Title:=ThisWorkbook.Application.Caption
    
    Do
      dataRemessa = InputBox("Qual a DATA DA REMESSA?" & Chr(13) & Chr(13) & "Somente N�meros. (EX: 01042022) :", "DATA DA REMESSA.")
      If (Not IsNumeric(dataRemessa) Or Len(dataRemessa) <> 8) And StrPtr(dataRemessa) <> 0 Then
        MsgBox "Data inv�lida!"
      End If
    Loop Until (RE.test(dataRemessa) And Len(dataRemessa) = 8) Or StrPtr(dataRemessa) = 0 'STRPTR CAPTURA SE USUARIO CLICOU NO BOTAO CANCELAR
    
       
    If StrPtr(dataRemessa) = 0 Then
      Err.Raise vbObjectError + 50, "CANCELAMENTO." _
                             , "LAN�AMENTO DE PEDIDOS FOI CANCELADO.!!"
    End If
 
    session.findById("wnd[0]").Maximize
      
    'digita /nKK89S no campo transacao e da enter
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nKK89S"
    session.findById("wnd[0]").sendVKey 0
     
    Set User = session.findById("wnd[0]/usr")
    
    'retorna o foco para a tela criar pedido
    AppActivate Title:="Criar pedido"
    
     
    'vai para pedido normal e volta para pedido proc licit. para liberar coluna contrato
    session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").Key = "NB"
    session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").Key = "ZLIC"
    'clica no botao ocultar cabecalho
   ' session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB1:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4000/btnDYN_4000-BUTTON").press
     
    'extrai o numero do contrato sem /2021
    esplitado = Split(Worksheets("Configura��es").Range("C15").Value, "/")
    numeroContrato = esplitado(0)
    
    codigoServico = Worksheets("Configura��es").Range("C17").Value
    
    
    
    
    
    '=================================================
    
    'captura a tablecontrol e joga na variavel grade
    Set grade = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUSGC_6487")
    'extrai numero de linhas visiveis no scrollbar
    tamanhoPagina = grade.visibleRowCount
    
    'array de Atividades que provocam deslocamento
    atividadesDeslocamento = Application.Transpose(Worksheets("Configura��es").Range("J12:J206"))
    'array de Atividades de SODA
    atividadesSODA = Application.Transpose(Worksheets("Configura��es").Range("L12:L206"))
    'array de Atividades de estudo de prote��o
    atividadesEstudoProtecao = Application.Transpose(Worksheets("Configura��es").Range("N12:N206"))
    
    
    'inicializa os itens sap a partir da planilha configura��es
    itemEstprotecao = Worksheets("Configura��es").Range("C19").Value
    itemSODA = Worksheets("Configura��es").Range("C21").Value
    itemComissionamento = Worksheets("Configura��es").Range("C23").Value
    itemAutAtivacao = Worksheets("Configura��es").Range("C25").Value
    
    
    'classifica planilha boletim
    classificaPlanilha ("Boletim")
    
    
    Set colecaoLinhas = planilhaParaColecao(Worksheets("Boletim"))
    
    posicaoSincronismoScrollbar = 0
    
    'preenche as linhas visiveis do scrollbar e rola a barra de rolagem no tamanho da pagina e assim por diante at� o numero de elementos
    numeroLinhaGrade = 0
    contItensPedido = 1
    For contColecaoLinhas = 1 To colecaoLinhas.Count
    
        cidadeAtual = colecaoLinhas(contColecaoLinhas)(2)
     
        'se a cidade do primeiro item ainda n�o foi lan�ada ou se for segunda cidade executa
        If colecaoLinhas(1)(13) = "" Or contColecaoLinhas > 1 Then
              
            'se for celula visivel, dentro da pagina
            If numeroLinhaGrade < tamanhoPagina Then
                
                'digita o numero do contrato
                grade.findById("ctxtMEPO1211-KONNR[27," & numeroLinhaGrade & "]").Text = numeroContrato
                'digita o numero do item de acordo com a atividade
                grade.findById("txtMEPO1211-KTPNR[28," & numeroLinhaGrade & "]").Text = extraiItem(colecaoLinhas(contColecaoLinhas)(5))
                
                    'da enter
                    session.findById("wnd[0]").sendVKey 0
                    
                    'quando da enter, a tela muda de id, para isso chamamos a funcao tela(User) para capturar a nova tela ID
                    Set grade = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUSGC_6487")
                    'toda vez que da enter pode acontecer do sap alterar o tamanho do scrollbar, convem atualizar
                    tamanhoPagina = grade.visibleRowCount
                
                 'digita se � pep(p) ou ordem (f)
                tipoClassificacao = extraiPepOrdem(colecaoLinhas(contColecaoLinhas)(4))
                grade.findById("ctxtMEPO1211-KNTTP[2," & numeroLinhaGrade & "]").Text = tipoClassificacao
                
                   
                'digita quantidade de US de acordo com a atividade
                If colecaoLinhas(contColecaoLinhas)(5) = "1317" Then
                    qtdUS = colecaoLinhas(contColecaoLinhas)(7) * colecaoLinhas(contColecaoLinhas)(8) * 0.022
                Else
                    qtdUS = colecaoLinhas(contColecaoLinhas)(7) * colecaoLinhas(contColecaoLinhas)(8)
                End If
                grade.findById("txtMEPO1211-MENGE[5," & numeroLinhaGrade & "]").Text = qtdUS
        
                    
                'inclui numero da os na descricao
                'grade.findById("txtMEPO1211-TXZ01[4," & numeroLinhaGrade & "]").Text = grade.findById("txtMEPO1211-TXZ01[4," & numeroLinhaGrade & "]").Text & " - OS " & colecaoLinhas(contColecaoLinhas)(3)
                palavra = grade.findById("txtMEPO1211-TXZ01[4," & numeroLinhaGrade & "]").Text
                If Len(palavra) > 17 Then palavra = Left(palavra, 17)
                grade.findById("txtMEPO1211-TXZ01[4," & numeroLinhaGrade & "]").Text = palavra & " - OS " & colecaoLinhas(contColecaoLinhas)(3) & " - ATV " & colecaoLinhas(contColecaoLinhas)(5)
        
            
                'digita data remessa
                grade.findById("ctxtMEPO1211-EEIND[7," & numeroLinhaGrade & "]").Text = dataRemessa
    
                    'da enter
                    session.findById("wnd[0]").sendVKey 0
                    
                    'quando da enter, a tela muda de id, para isso chamamos a funcao tela(User) para capturar a nova tela ID
                    Set grade = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUSGC_6487")
                    'toda vez que da enter pode acontecer do sap alterar o tamanho do scrollbar, convem atualizar
                    tamanhoPagina = grade.visibleRowCount
                    
                'se n�o existir grupo de abas
                If session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsDSEM_SD", False) Is Nothing Then
                    'clica no botao ocultar detalhe de itens para exibir as abas
                    session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON").press
                End If
                    
                'utilizamos aqui artificio para garantir que o sap realmente abriu a aba(pegamos um componente da aba e verificamos se o componente existe)
                'abre a aba abaBrasil e digita codigoServico no edittext
                While abaBrasil.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1326/ctxtMEPO1326-J_1BNBM", False) Is Nothing
                  abaBrasil.Select
                Wend
                abaBrasil.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1326/ctxtMEPO1326-J_1BNBM").Text = codigoServico
              
                'abre a aba abaClassCont e digita colecaoLinhas(contColecaoLinhas)(4) no edittext
                If tipoClassificacao = "P" Then
                
                    While abaClassCont.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-PS_POSID", False) Is Nothing
                        abaClassCont.Select
                    Wend
                    abaClassCont.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-PS_POSID").Text = colecaoLinhas(contColecaoLinhas)(4)
                    
                    
                    'Campo conta do razao
                    abaClassCont.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/ctxtMEACCT1100-SAKTO").Text = Worksheets("Configura��es").Range("C27").Value
                     

                    While abaClassCont.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-PS_POSID", False) Is Nothing
                        abaClassCont.Select
                    Wend

                    'Campo fundos destins.
                    abaClassCont.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KBLNR").Text = Worksheets("Configura��es").Range("C30").Value
                    

                    'Campo digito fundos destins.
                    abaClassCont.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KBLPOS").Text = Worksheets("Configura��es").Range("C31").Value
                     
                    
                ElseIf tipoClassificacao = "F" Then
                
                    While abaClassCont.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-AUFNR", False) Is Nothing
                        abaClassCont.Select
                    Wend
                    abaClassCont.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-AUFNR").Text = colecaoLinhas(contColecaoLinhas)(4)
                    
                    'Campo conta do razao
                    abaClassCont.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/ctxtMEACCT1100-SAKTO").Text = Worksheets("Configura��es").Range("C28").Value

                    While abaClassCont.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-AUFNR", False) Is Nothing
                        abaClassCont.Select
                    Wend

                    'Campo fundos destins.
                    abaClassCont.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KBLNR").Text = Worksheets("Configura��es").Range("C33").Value

                    'Campo digito fundos destins.
                    abaClassCont.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KBLPOS").Text = Worksheets("Configura��es").Range("C34").Value

                    
                End If
                
                session.findById("wnd[0]").sendVKey 0
                verificaSbar
                
                    
                'se n�o existir grupo de abas
                If session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsDSEM_SD", False) Is Nothing Then
                    'clica no botao ocultar detalhe de itens para exibir as abas
                    session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON").press
                End If
             
                
            'se n�o for celula visivel rola a tela para baixo
            Else
                posicaoScrollbarDesejada = grade.verticalScrollbar.Position + tamanhoPagina - 1
                
                While grade.verticalScrollbar.Position <> posicaoScrollbarDesejada
                    'scrolldown tamanhoPagina
                    grade.verticalScrollbar.Position = grade.verticalScrollbar.Position + tamanhoPagina - 1
                    'quando da enter, a tela muda de id, para isso chamamos a funcao tela(User) para capturar a nova tela ID
                    Set grade = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUSGC_6487")
                Wend
                 
                'clica no botao ocultar cabecalho
                'session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB1:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4000/btnDYN_4000-BUTTON").press
    
                'quando scrolla, a tela muda de id, para isso chamamos a funcao tela(User) para capturar a nova tela ID
                Set grade = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUSGC_6487")
                'toda vez que da scrola pode acontecer do sap alterar o tamanho do scrollbar, convem atualizar
                tamanhoPagina = grade.visibleRowCount
                
                contColecaoLinhas = contColecaoLinhas - 1
                contItensPedido = contItensPedido - 1
                
                numeroLinhaGrade = 1
                posicaoSincronismoScrollbar = grade.verticalScrollbar.Position
                
                'botao ocultar detalhes de itens e cabecalho
                'session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB1:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4000/btnDYN_4000-BUTTON").press
                'session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON").press
                
            End If
            
                
                
                 If contColecaoLinhas < colecaoLinhas.Count Then
                    proximaCidade = colecaoLinhas(contColecaoLinhas + 1)(2)
                 Else
                    proximaCidade = ""
                 End If
                 
                 
                If proximaCidade <> cidadeAtual Then
                
                    If cidadeAtual <> "CURITIBA" Then
                    
                        'quando tem apenas um item no pedido n�o precisa replicar para todos os itens o nome da cidade
                        If contItensPedido = 1 Then
    
                            'abre a aba abaFatura
                            While abaFatura.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ", False) Is Nothing
                              abaFatura.Select
                            Wend
                            
                            'preenche campos
                            'abre a aba abaEndRemessa e digita "" no edittext da rua
                            While abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-STREET", False) Is Nothing
                              abaEndRemessa.Select
                            Wend
                            abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-STREET").Text = ""
                            
                            'abre a aba abaEndRemessa e digita "" no edittext do numero da casa
                            While abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-HOUSE_NUM1", False) Is Nothing
                              abaEndRemessa.Select
                            Wend
                            abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-HOUSE_NUM1").Text = ""
                            
                            If cidadeAtual = "SCHROEDER" Then
                                
                                'abre a aba abaEndRemessa e digita "SC" no edittext da regiao
                                While abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-REGION", False) Is Nothing
                                 abaEndRemessa.Select
                                Wend
                                abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-REGION").Text = "SC"
                                
                            End If
                            
                            'abre a aba abaEndRemessa e digita cidadeAtual no edittext da cidade
                            While abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-CITY1", False) Is Nothing
                             abaEndRemessa.Select
                            Wend
                            abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-CITY1").Text = cidadeAtual
                            
                            session.findById("wnd[0]").sendVKey 0
                            verificaSbar
                            
                        Else
                        
                            'abre a aba abaFatura
                            While abaFatura.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ", False) Is Nothing
                              abaFatura.Select
                            Wend
                           
                            'abre a aba abaEndRemessa e digita "" no edittext da rua
                            While abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-STREET", False) Is Nothing
                              abaEndRemessa.Select
                            Wend
                            abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-STREET").Text = ""
                            
                            session.findById("wnd[0]").sendVKey 0
                            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                            
                            'abre a aba abaEndRemessa e digita "" no edittext do numero da casa
                            While abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-HOUSE_NUM1", False) Is Nothing
                              abaEndRemessa.Select
                            Wend
                            abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-HOUSE_NUM1").Text = ""
                            
                            session.findById("wnd[0]").sendVKey 0
                            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                            
                            If cidadeAtual = "SCHROEDER" Then
                            
                                'abre a aba abaEndRemessa e digita "SC" no edittext da regiao
                                While abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-REGION", False) Is Nothing
                                 abaEndRemessa.Select
                                Wend
                                abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-REGION").Text = "SC"
                                
                            End If
                            
                            'abre a aba abaEndRemessa e digita cidadeAtual no edittext da cidade
                            While abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-CITY1", False) Is Nothing
                             abaEndRemessa.Select
                            Wend
                            abaEndRemessa.findById("ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1330/ssubADDRESS:SAPLMMDA:0200/ssubADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-CITY1").Text = cidadeAtual
                                                 
                            session.findById("wnd[0]").sendVKey 0
                            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                            
                        End If
                    
                    End If
                    
                    
                    'retorna o foco para o excel
                     AppActivate Title:=ThisWorkbook.Application.Caption
                     
                     linhaNumeroPedido = contColecaoLinhas
                    
                    If contColecaoLinhas = colecaoLinhas.Count Then
                    
                         'retorna o foco para o excel
                         AppActivate Title:=ThisWorkbook.Application.Caption
                         
                         If MsgBox("Foi conclu�do o lan�amento da cidade: " & cidadeAtual & Chr(13) & Chr(13) & "DESEJA GRAVAR O PEDIDO?", vbOKCancel, "ATEN��O.") = vbOK Then

                              'Bot�o disquete
                               session.findById("wnd[0]/tbar[0]/btn[11]").press
                              'ATEN��O! Grava o pedido
                               session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press


                              'captura numero do pedido gerado
                               capturaNumeroPedido session.findById("wnd[0]/sbar").Text, contColecaoLinhas
                              'capturaNumeroPedido "pedido n� 1234567", linhaNumeroPedido

                              MsgBox "Foi GRAVADO o pedido da �LTIMA CIDADE: " & cidadeAtual & "."

                         Else

                              MsgBox "A GRAVA��O do pedido da cidade: " & cidadeAtual & " foi CANCELADA!" & Chr(13) & Chr(13) & "Foram conclu�dos TODOS os lan�amentos." & Chr(13) & Chr(13) & "�LTIMA cidade: " & cidadeAtual & "."

                         End If
                             
                         ActiveWindow.ScrollRow = contColecaoLinhas
                         
                         MsgBox "Foram conclu�dos TODOS os lan�amentos!"
                         
                    Else
                    
                         If colecaoLinhas(contColecaoLinhas + 1)(13) <> "" Then
                         
                            Do While colecaoLinhas(contColecaoLinhas + 1)(13) <> ""
                                contColecaoLinhas = contColecaoLinhas + 1
                                If contColecaoLinhas = colecaoLinhas.Count Then
                                    Exit Do
                                End If
                            Loop
        
                            
                            If contColecaoLinhas < colecaoLinhas.Count Then
                                 proximaCidade = colecaoLinhas(contColecaoLinhas + 1)(2)
                            Else
                                 proximaCidade = "TODAS AS CIDADES J� FORAM LAN�ADAS"
                            End If
                         
                         End If

                         'retorna o foco para o excel
                         AppActivate Title:=ThisWorkbook.Application.Caption
                         
                         If MsgBox("Foi conclu�do o lan�amento da cidade: " & cidadeAtual & Chr(13) & Chr(13) & "DESEJA GRAVAR O PEDIDO?", vbOKCancel, "ATEN��O.") = vbOK Then

                              'Bot�o disquete
                               session.findById("wnd[0]/tbar[0]/btn[11]").press
                              'ATEN��O! Grava o pedido
                               session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press


                              'captura numero do pedido gerado
                              capturaNumeroPedido session.findById("wnd[0]/sbar").Text, contColecaoLinhas
                              'capturaNumeroPedido "pedido n� 1234567", linhaNumeroPedido

                              If MsgBox("Foi GRAVADO o pedido da cidade: " & cidadeAtual & Chr(13) & Chr(13) & "Clique em OK para iniciar o lan�amento da pr�xima cidade: " & Chr(13) & Chr(13) & proximaCidade & ".", vbOKCancel, "ATEN��O.") = vbCancel Then
                                 Err.Raise vbObjectError + 50, "CANCELAMENTO." _
                                     , "LAN�AMENTO DE PEDIDOS FOI CANCELADO.!!"
                              Else
                                If contColecaoLinhas = colecaoLinhas.Count Then
                                    MsgBox "Foram conclu�dos TODOS os lan�amentos!"
                                    Exit Sub
                                End If
                              End If

                         Else

                             If MsgBox("A GRAVA��O do pedido da cidade: " & cidadeAtual & " foi CANCELADA!" & Chr(13) & Chr(13) & "Clique em OK para iniciar o lan�amento da pr�xima cidade: " & Chr(13) & Chr(13) & proximaCidade & ".", vbOKCancel, "ATEN��O.") = vbCancel Then
                                Err.Raise vbObjectError + 50, "CANCELAMENTO." _
                                     , "LAN�AMENTO DE PEDIDOS FOI CANCELADO.!!"
                              Else
                                If contColecaoLinhas = colecaoLinhas.Count Then
                                    MsgBox "Foram conclu�dos TODOS os lan�amentos!"
                                    Exit Sub
                                End If
                              End If

                         End If
                        
                         ActiveWindow.ScrollRow = contColecaoLinhas
                         
                         'retorna o foco para a tela criar pedido
                         AppActivate Title:="Criar pedido"
                         
                         'digita /nKK89S no campo transacao e da enter
                         session.findById("wnd[0]/tbar[0]/okcd").Text = "/nKK89S"
                         session.findById("wnd[0]").sendVKey 0
                         
                         'vai para pedido normal e volta para pedido proc licit. para liberar coluna contrato
                         session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").Key = "FD"
                         session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").Key = "HSRE"
                         
                        
                         Set grade = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUSGC_6487")
                         tamanhoPagina = grade.visibleRowCount
                         posicaoSincronismoScrollbar = 0
                         numeroLinhaGrade = 0
                         contItensPedido = 0
                    
                    End If
                Else
                
                    'clica no botao ocultar detalhe de itens
                    session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDSD_6400-BUTTON").press
                    'captura a tablecontrol e joga na variavel grade, atualiza pq o ID da tela muda a todo instante
                    Set grade = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUSGC_6487")
                    'o sap tem a tela dinamica, conforme o tamanho da aba diminui, aumenta o tamanho do scrollbar e pode ocorrer de scrolar sozinho
                    tamanhoPagina = grade.visibleRowCount
                    
                    'caso o sap scrole automaticamente, o contador � zerado
                    If grade.verticalScrollbar.Position = posicaoSincronismoScrollbar Then
                        numeroLinhaGrade = numeroLinhaGrade + 1
                    Else
                        numeroLinhaGrade = contItensPedido - grade.verticalScrollbar.Position
                        posicaoSincronismoScrollbar = grade.verticalScrollbar.Position
                    End If
                     
                End If
                
               contItensPedido = contItensPedido + 1
        Else
        
            Do While colecaoLinhas(contColecaoLinhas + 1)(13) <> ""
                contColecaoLinhas = contColecaoLinhas + 1
                If contColecaoLinhas = colecaoLinhas.Count Then
                    Exit Do
                End If
            Loop
            
             If contColecaoLinhas < colecaoLinhas.Count Then
                 proximaCidade = colecaoLinhas(contColecaoLinhas + 1)(2)
             Else
                 proximaCidade = "TODAS AS CIDADES J� FORAM LAN�ADAS"
             End If
            
            
            'retorna o foco para o excel
            AppActivate Title:=ThisWorkbook.Application.Caption
            
            If MsgBox("O PEDIDO da cidade " & cidadeAtual & " j� foi lan�ado ANTERIORMENTE!" & Chr(13) & Chr(13) & "Clique em OK para INICIAR o lan�amento da pr�xima cidade: " & Chr(13) & Chr(13) & proximaCidade & ".", vbOKCancel, "ATEN��O.") = vbCancel Then
                Err.Raise vbObjectError + 50, "CANCELAMENTO." _
                     , "LAN�AMENTO DE PEDIDOS FOI CANCELADO.!!"
            Else
              If contColecaoLinhas = colecaoLinhas.Count Then
                  MsgBox "Foram conclu�dos TODOS os lan�amentos!"
                  Exit Sub
              End If
            End If
            
        End If
  
    Next

Exit Sub

Erro:
     MsgBox " Erro: " & Err.Description & Chr(13) & Chr(13) & "Local: M�dulo - UserFormBmd.LancarPedido"


        
End Sub
Function verificaSbar()

    If session.findById("wnd[0]/sbar").MessageType = "E" Then
    
        If MsgBox("OCORREU UM ERRO L� NO SAP DESEJA CONTINUAR?", vbOKCancel, "ATEN��O.") = vbCancel Then
          Err.Raise vbObjectError + 50, "CANCELAMENTO." _
               , "LAN�AMENTO DE PEDIDOS FOI CANCELADO.!!"
        End If
     
        'Err.Raise vbObjectError + 50, "ERRO NO SAP!" _
         '                  , "DEU ALGUM ERRO L� NO SAP. LAN�AMENTO CANCELADO."
                        
    End If


    While session.findById("wnd[0]/sbar").Text <> "" And session.findById("wnd[0]/sbar").MessageType <> "E"
    
        'da enter
        session.findById("wnd[0]").sendVKey 0
       
    
        If session.findById("wnd[0]/sbar").MessageType = "E" Then
        
            If MsgBox("OCORREU UM ERRO L� NO SAP DESEJA CONTINUAR?", vbOKCancel, "ATEN��O.") = vbCancel Then
              Err.Raise vbObjectError + 50, "CANCELAMENTO." _
                   , "LAN�AMENTO DE PEDIDOS FOI CANCELADO.!!"
            End If
         
'            Err.Raise vbObjectError + 50, "ERRO NO SAP!" _
'                               , "DEU ALGUM ERRO L� NO SAP. LAN�AMENTO CANCELADO."
                            
        End If
         
    Wend
   
End Function
Function abaFatura()

    'esta funcao atualiza a tela que o sap muda constantemente e devolve a abaFatura vigente
    
    'seleciona item na lista acima das abas
    Set listaItens = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbSDN_2300-LIST")
    listaItens.Key = formataKey(contItensPedido)
         
    'captura o grupo de abas
    Set grupoDeAbas = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsDSEM_SD")
    'aba Brasil codigo ncm
    Set abaFatura = grupoDeAbas.findById("tabpTABIDT7")
    
    While grupoDeAbas.findById("tabpTABIDT17", False) Is Nothing
        Set abaFatura = grupoDeAbas.findById("tabpTABIDT17")
    Wend
            
End Function
Function abaBrasil()

    'esta funcao atualiza a tela que o sap muda constantemente e devolve a abaBrasil vigente
    
    'seleciona item na lista acima das abas
    Set listaItens = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbSDN_2300-LIST")
    listaItens.Key = formataKey(contItensPedido)
         
    'captura o grupo de abas
    Set grupoDeAbas = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsDSEM_SD")
    'aba Brasil codigo ncm
    Set abaBrasil = grupoDeAbas.findById("tabpTABIDT12")
    While grupoDeAbas.findById("tabpTABIDT12", False) Is Nothing
        Set abaBrasil = grupoDeAbas.findById("tabpTABIDT12")
    Wend
End Function
Function abaClassCont()

    'esta funcao atualiza a tela que o sap muda constantemente e devolve a abaClassCont vigente
    
    'seleciona item na lista acima das abas
    Set listaItens = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbSDN_2300-LIST")
    
    listaItens.Key = formataKey(contItensPedido)
     
    'captura o grupo de abas
    Set grupoDeAbas = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsDSEM_SD")
    'aba classificacao contabil
    Set abaClassCont = grupoDeAbas.findById("tabpTASHDT13")
    While grupoDeAbas.findById("tabpTASHDT13", False) Is Nothing
        Set abaClassCont = grupoDeAbas.findById("tabpTASHDT13")
    Wend
            
End Function
Function abaEndRemessa()

    'esta funcao atualiza a tela que o sap muda constantemente e devolve a abaEndRemessa vigente
    
    'seleciona item na lista acima das abas
    Set listaItens = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbSDN_2300-LIST")
    listaItens.Key = formataKey(contItensPedido)
         
    'captura o grupo de abas
    Set grupoDeAbas = session.findById("wnd[0]/usr/sub" & tela(User) & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsDSEM_SD")
    'aba classificacao contabil
    Set abaEndRemessa = grupoDeAbas.findById("tabpTABIDT16")
    While grupoDeAbas.findById("tabpTABIDT16", False) Is Nothing
        Set abaEndRemessa = grupoDeAbas.findById("tabpTABIDT16")
    Wend
            
End Function
Function abreSessao(ByVal nomeSistemaDesejado As String, ByVal nomeConexaoDesejada As String)

    Set WSHShell = CreateObject("WScript.Shell")
    
    'verifica se o sap ja esta aberto
    If Not WSHShell.AppActivate("SAP Logon ") Then
        'abre o aplicativo sap logon
         Shell "C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe", vbNormalFocus
        'espera o aplicativo abrir para continuar
         Do Until WSHShell.AppActivate("SAP Logon ")
            Application.Wait Now + TimeValue("0:00:01")
         Loop
    End If
    
    
    'para desabilitar alertas do SAP -> SAP GUI -> alt+f12 -> Op��es -> accessibilidade e Scripting -> scripting -> desmarcar notifica��es
    If Not IsObject(SAPGuiApp) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set SAPGuiApp = SapGuiAuto.GetScriptingEngine
    End If
    
    If Not IsObject(Connection) Then

        Set Connections = SAPGuiApp.Connections()
        'verifica se existe alguma conex�o do sap aberta(ex: produ��o, qualidade)
      
        sessaoEncontrada = False
        
        If Connections.Count = 0 Then
        
            'caso n�o haja nenhuma conex�o aberta abre o qualidade
            Set Connection = SAPGuiApp.openConnection(nomeConexaoDesejada)
            Set session = Connection.Children(0)
            
        Else
        
            'caso haja conex�o aberta verifica se algum deles � qualidade e captura a sess�o para uso
            For Each Connection In Connections
              Set sessions = Connection.sessions()
                  For Each sess In sessions
                      If sess.Busy() = vbFalse Then
                      
                          If sess.Info().systemname() = nomeSistemaDesejado Then
                                  Set session = sess
                                  sessaoEncontrada = True
                                  Exit For
                          End If
    
                      End If
                  Next
             Next
    
             'caso nenhuma das sess�es aberta sejam o qualidade abre qualidade e captura sess�o
             If Not sessaoEncontrada Then
                
               Set Connection = SAPGuiApp.openConnection(nomeConexaoDesejada)
    
               If Not IsObject(session) Then
                 Set session = Connection.Children(0)
               End If
             End If
         End If
     End If
    
    
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject SAPGuiApp, "on"
    End If
    
   
    
   Set abreSessao = session
 
End Function
Sub capturaNumeroPedido(ByVal frase As String, ByVal linha As String)
    
    esplitado = Split(frase, "n�")
    If tamanhoArray(esplitado) > 1 Then
        If IsNumeric(Trim(esplitado(1))) Then
        
            qtdLinhas = Worksheets("Boletim").Range("A1").CurrentRegion.Rows.Count
            arraycolunas = Worksheets("Boletim").Range("M2:M" & qtdLinhas)
            
            For c = linha To linha - contItensPedido + 1 Step -1
                arraycolunas(c, 1) = Trim(esplitado(1))
            Next
     
            Worksheets("Boletim").Range("M2:M" & qtdLinhas).Value = Application.Transpose(Application.Transpose(arraycolunas))
      
        End If
    End If
    
    Application.Wait Now + TimeValue("0:00:01")
    
End Sub
Function estaNoArray(ByVal valorProcurado As String, ByVal vetor As Variant) As Boolean
  estaNoArray = (UBound(Filter(vetor, valorProcurado)) > -1)
End Function
Function extraiPepOrdem(ByVal elemento As Variant)

    If elemento Like "I*" Then
        extraiPepOrdem = "P"
    Else
        extraiPepOrdem = "F"
    End If

End Function
Function extraiItem(ByVal atividade As String)

    If atividade = "1315" Or atividade = "1316" Then
        extraiItem = itemComissionamento
    ElseIf estaNoArray(atividade, atividadesEstudoProtecao) Or atividade = "356" Or atividade = "357" Then
        extraiItem = itemEstprotecao
    ElseIf estaNoArray(atividade, atividadesSODA) Then
        extraiItem = itemSODA
    ElseIf estaNoArray(atividade, atividadesDeslocamento) Or atividade = "1317" Or atividade = "193" Then
        extraiItem = itemAutAtivacao
    Else
        extraiItem = 0
    End If
    

End Function
Function formataKey(ByVal contador As String)

    Select Case Len(contador)
        Case Is = 4: formataKey = contador
        Case Is = 3: formataKey = " " & contador
        Case Is = 2: formataKey = "  " & contador
        Case Is = 1: formataKey = "   " & contador
 
    Case Else
        formataKey = 0
    End Select
  
End Function
Function tela(ByVal usuario As Variant)

    'esta funcao busca a tela dentre os filhos da tela usuario que se parecem com SUB0:SAPLMEGUI:
    For i = 0 To usuario.Children.Count - 1
      nome = usuario.Children(CInt(i)).Name
      If Left(nome, 15) = "SUB0:SAPLMEGUI:" Then
        tela = nome
        Exit For
      End If
    Next
    
End Function


Private Sub btValidarPep_Click()

    On Error GoTo Erro

    If planilhaVazia(Worksheets("Boletim")) Then

    Err.Raise ERRO_PLANILHA_VAZIA, "Planilha vazia!" _
                        , "Nenhum ITEM na planilha BOLETIM."

    End If
    
    qtdLinhas = Worksheets("Boletim").Range("A1").CurrentRegion.Rows.Count

    'verifica celulas vazias
    For Each celula In Worksheets("Boletim").Range("D2:D" & qtdLinhas)
        If IsEmpty(celula) Then
            Err.Raise vbObjectError + 50, "DADO INV�LIDO!" _
                            , "Alguma c�lula est� vazia !! Verifique coluna: STATUS_PEP."
        End If
    Next

    qtdLinhas = Worksheets("Boletim").Range("A1").CurrentRegion.Rows.Count

   'transfere a coluna pep/ordem_interna da planilha boletim para um array
    arrayColunaSituacaoPep = Worksheets("Boletim").Range("D2:D" & qtdLinhas)
    
    contadorPeps = 0
    For Each elemento In arrayColunaSituacaoPep
     If elemento Like "I*" Then
        contadorPeps = contadorPeps + 1
     End If
    Next
    
    'se consultar apenas ordem no sap o sistema abre outra tela para informar que nada foi encontrado, por isso se houver apenas ordem n�o executa.
    If contadorPeps > 0 Then
    
        'nomeSistemaDesejado: para qualidade usar "EQ1", para produ��o usar "GCE"
        'nomeConexaoDesejada: para qualidade usar "EQ1 - Qualidade", para produ��o usar "GCE PUBLIC"
        Set session = abreSessao(nomeSistemaDesejado:="GCE", nomeConexaoDesejada:="GCE PUBLIC")

        session.findById("wnd[0]").Maximize
    
        'digita CL4A no campo transacao e da enter
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCL4A"
        session.findById("wnd[0]").sendVKey 0
            
        'verifica se � a primeira vez pelo tipo da tela ativa GuiMainWindow ou GuiModalWindow
        If session.ActiveWindow.Type = "GuiModalWindow" Then
            session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").Text = "000000000001" 'somente na primeira vez que abre o sap
            session.findById("wnd[1]/tbar[0]/btn[0]").press 'ok
        End If
        
        'retorna o foco para a tela "Sistema de informa��o de projetos: 1� tela elementos PEP"
        AppActivate Title:="Sistema de informa��o de projetos: 1� tela elementos PEP"
    
        'limpa os campos "at�"
        session.findById("wnd[0]/usr/ctxtCS_PSPNR-LOW").Text = ""
        session.findById("wnd[0]/usr/ctxtCS_PSPNR-HIGH").Text = ""
        session.findById("wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH").press 'botao abertura da janela de selecao pep

        'criando arquivo txt
        diretorio = Application.ActiveWorkbook.Path & "\"
        nomeArquivo = "temporarioPep.txt"
        caminhoArquivo = diretorio & nomeArquivo
        Set fso = New Scripting.FileSystemObject
        Set arquivoTxt = fso.CreateTextFile(caminhoArquivo, True)
    
        'escrevendo os pep�s no arquivo txt
        For Each elemento In arrayColunaSituacaoPep
            arquivoTxt.WriteLine elemento
        Next
        arquivoTxt.Close
    
        'carregando arquivo contendo os pep�s
        session.findById("wnd[1]/tbar[0]/btn[23]").press
        session.findById("wnd[2]/usr/ctxtDS_PATH").Text = diretorio
        session.findById("wnd[2]/usr/ctxtDS_FILENAME").Text = nomeArquivo
        session.findById("wnd[2]/tbar[0]/btn[0]").press
    
        'desagrupando
        session.findById("wnd[1]/tbar[0]/btn[8]").press 'botao fechamento da janela de selecao pep
        session.findById("wnd[0]/tbar[1]/btn[8]").press 'botao executar relatorio(relogio verde)
        session.findById("wnd[0]/mbar/menu[4]/menu[6]/menu[0]").Select 'menu configuracao->
        
        'extraindo numero de elementos "N�mero WBS element: 6"
        esplitado = Split(session.findById("wnd[0]/usr/sub/1[0,0]/sub/1/2[0,0]/lbl[0,0]").Text, ":")
        numeroElementos = Trim(esplitado(1))
        
        tamanhoPagina = session.findById("wnd[0]/usr").verticalScrollbar.pageSize
        
        Set colecaoPeps = New Collection
        
        contador = 5
        For contadorParalelo = 0 To numeroElementos - 1
            'se for celula visivel, dentro da pagina ou se n�o existir scrollbar (tamanhoPagina = 0)
            If contador < tamanhoPagina + 4 Or tamanhoPagina = 0 Then
            
                'retirando os tracos do pep do sap
                chavePep = Replace(session.findById("wnd[0]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/" & contador & "[0," & contador & "]/lbl[2," & contador & "]").Text, "-", "")
                Status = session.findById("wnd[0]/usr/sub/1[0,0]/sub/1/3[0,1]/sub/1/3/" & contador & "[0," & contador & "]/lbl[72," & contador & "]").Text
                colecaoPeps.Add Status, chavePep
                contador = contador + 1
                
            'se n�o for celula visivel rola a tela para baixo
            Else
                'scrolldown tamanhoPagina
                session.findById("wnd[0]/usr").verticalScrollbar.Position = session.findById("wnd[0]/usr").verticalScrollbar.Position + tamanhoPagina
                contadorParalelo = contadorParalelo - 1
                contador = 4
            End If
        Next
    
    End If
    
    Set colecaoLinhas = planilhaParaColecao(Worksheets("Boletim"))
    Set colecaoOsPepInconsistente = New Collection
    Set colecaoPendenciasPep = New Collection
    
        
    'preenche a coluna status_pep na planilha boletim
    CONT = 1
    For Each linha In colecaoLinhas

     arrayAuxiliar = linha
     numOS = CStr(arrayAuxiliar(3))
        
        'se a string iniciar com I
        If arrayAuxiliar(4) Like "I*" Then
            'se o sap encontrou o pep e seu status
            If existeItemColecao(colecaoPeps, arrayAuxiliar(4)) Then
                estado = extraiStatus(colecaoPeps(arrayAuxiliar(4)))
                'Debug.Print arrayAuxiliar(3) & " " & arrayAuxiliar(4) & " " & estado
                If estado = "LIBERADO" Then
                    arrayAuxiliar(14) = "PEP OK"
                    colecaoLinhas.Add arrayAuxiliar, after:=CONT
                    colecaoLinhas.Remove (CONT)
                Else
                
                    If Not existeItemColecao(colecaoOsPepInconsistente, numOS) Then
                        colecaoOsPepInconsistente.Add numOS, numOS
                        adicionaPendenciaPep "Erro! OS N�o ser� paga!", "PEP Impedido (" & arrayAuxiliar(4) & " - " & estado & " )", numOS
                    End If
                     
                     colecaoLinhas.Remove (CONT)
                     CONT = CONT - 1
                End If
            
            Else
            
                If Not existeItemColecao(colecaoOsPepInconsistente, arrayAuxiliar(3)) Then
                    colecaoOsPepInconsistente.Add numOS, numOS
                    adicionaPendenciaPep "Erro! OS N�o ser� paga!", "PEP Inexistente(" & arrayAuxiliar(4) & ")", numOS
                End If
                
                colecaoLinhas.Remove (CONT)
                CONT = CONT - 1
            End If
        Else
            arrayAuxiliar(14) = "ORDEM"
            colecaoLinhas.Add arrayAuxiliar, after:=CONT
            colecaoLinhas.Remove (CONT)
        End If

        

     CONT = CONT + 1
    Next

    'devolve a colecaolinhas para a planilha boletim, agora com os dados da coluna status_pep
    colecaoParaPlanilha colecaoLinhas, Worksheets("Boletim")
    
    'atualiza planilha pendencias
    colecaoParaPlanilhaPepPendencias colecaoPendenciasPep, Worksheets("Pendencias")
    
    classificaPlanilhaSimples "Pendencias", "B1"
    Worksheets("Pendencias").Columns.AutoFit
    
    Worksheets("Boletim").Activate
    
    'imprimi conteudo da planilha pendencias na textbox do userformbmd
    escrevePendencias
    
    
  
 
 'retorna o foco para o excel
 AppActivate Title:=ThisWorkbook.Application.Caption
 
 MsgBox ("Processo de valida��o de PEP�s foi concluido!")
 


'============================================================================================================================================================================
'FINAL SCRIPT
'============================================================================================================================================================================

Exit Sub

Erro:
     MsgBox " Erro: " & Err.Description & Chr(13) & Chr(13) & "Local: M�dulo - UserFormBmd.ValidarPep"

End Sub
Sub adicionaPendenciaPep(ByVal texto1 As String, ByVal texto2 As String, ByVal texto3 As String)
    ReDim arrayPendencia(1 To 3) As String
    arrayPendencia(1) = texto1
    arrayPendencia(2) = texto2
    arrayPendencia(3) = texto3
    colecaoPendenciasPep.Add arrayPendencia
End Sub
Function extraiStatus(ByVal termo As String)

    If InStr(termo, "ENCE") = 0 Then
        If InStr(termo, "ENCT") = 0 Then
            If InStr(termo, "ENTE") = 0 Then
                If InStr(termo, "CONC") = 0 Then
                    If InStr(termo, "CANC") = 0 Then
                        If InStr(termo, "BLOQ") = 0 Then
                            extraiStatus = "LIBERADO"
                        Else
                            extraiStatus = "BLOQ"
                        End If
                    Else
                        extraiStatus = "CANC"
                    End If
                Else
                    extraiStatus = "CONC"
                End If
            Else
                extraiStatus = "ENTE"
            End If
        Else
            extraiStatus = "ENCT"
        End If
    Else
     extraiStatus = "ENCE"
    End If
  
End Function
Private Sub btPreBmd_Click()

   On Error GoTo TE

    If planilhaVazia(Worksheets("Boletim")) Then

        Err.Raise ERRO_PLANILHA_VAZIA, "Planilha vazia!" _
                        , "Nenhum ITEM na planilha BOLETIM."

    End If
    
    qtdLinhas = Worksheets("Boletim").Range("A1").CurrentRegion.Rows.Count
    
    'verifica celulas vazias ou n�o numericas
    For Each celula In Worksheets("Boletim").Range("N2:N" & qtdLinhas)
        If IsEmpty(celula) Then
            Err.Raise vbObjectError + 50, "DADO INV�LIDO!" _
                            , "Alguma c�lula est� vazia !! Verifique coluna: SITUACAO_PEP."
        End If
    Next

     caminhoArquivo = Application.ActiveWorkbook.Path & "\registroLancamentosBmdsBD.csv"

    If Not arquivoExiste(caminhoArquivo) Then

        Err.Raise vbObjectError + 50, "ARQUIVO DE REGISTRO N�O ENCONTRADO!" _
                            , "Arquivo de registro de lan�amentos n�o foi encontrado!"

    End If

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .EnableAnimations = False
        .Calculation = xlCalculationManual
    End With

    'apaga todas as planilhas ap�s bmd
    Application.DisplayAlerts = False
    While Worksheets(Sheets.Count).Name <> "Bmd"
        Worksheets(Sheets.Count).Delete
    Wend
    Application.DisplayAlerts = True

    Application.PrintCommunication = False ' essencial pois pagesetup fica muito lento sem esta linha, porem deve ser reativado logo apos o page setup
    'configura margem do layout da planilha bmd para otimizar impressao em pdf
    With Worksheets("Bmd").PageSetup
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.1)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.2)
    End With
    Application.PrintCommunication = True ' essencial pois pagesetup fica muito lento sem esta linha, porem deve ser reativado logo apos o page setup
    'configura margem do layout da planilha bmd para otimizar impressao em pdf

    'classifica planilha boletim
    classificaPlanilha ("Boletim")


    Set colecaoLinhas = planilhaParaColecaoBmd(Worksheets("Boletim"))

    Set colecaoChavesBmds = New Collection
    Set colecaoBmds = New Collection
    Set colecaoTrinta = New Collection

    primeiroDiaMes = DateSerial(Year(colecaoLinhas(1)(1)), Month(colecaoLinhas(1)(1)), 1)
    ultimoDiaMes = DateSerial(Year(colecaoLinhas(1)(1)), Month(colecaoLinhas(1)(1)) + 1, 1) - 1
    
    jaIncrementouNumBmd = False
    
    'verifica se existe celula preenchida na coluna num_bmd
    For Each celula In Worksheets("Boletim").Range("K2:K" & qtdLinhas)
        If Not IsEmpty(celula) Then
            jaIncrementouNumBmd = True
            Exit For
        End If
    Next

    If jaIncrementouNumBmd Then
        numeroBmd = Worksheets("Boletim").Range("K2").Value
    Else
        numeroBmd = consultaUltimoBmd + 1
    End If
   
          'determinando as chaves dos bmds mes/ano + localizacao (jan-22TAMARANA)
          For Each elemento In colecaoLinhas
            If Not existeItemColecao(colecaoChavesBmds, elemento(1) & elemento(2)) Then
             colecaoChavesBmds.Add elemento(1) & elemento(2), elemento(1) & elemento(2)
            End If
          Next

        Worksheets("Bmd").Unprotect

        'para cada cidade gera um bmd
         contador = 1
         For Each chave In colecaoChavesBmds

               Set colecaoBmds = New Collection

               CONT = 1
               For Each linha In colecaoLinhas
                If linha(1) & linha(2) = chave Then

                     'adiciona as linhas da planilha boletim que pertencem a mesma cidade na colecaobmds
                     colecaoBmds.Add linha

                     'atribui ao campo sequencia o valor respectivo � cidade e preenche a coluna num_bmd
                     arrayAuxiliar = linha
                     arrayAuxiliar(11) = numeroBmd
                     arrayAuxiliar(12) = contador
                     colecaoLinhas.Add arrayAuxiliar, after:=CONT
                     colecaoLinhas.Remove (CONT)

                End If
                CONT = CONT + 1
               Next
               
               numeroContrato = Worksheets("Configura��es").Range("C15").Value
               
                'Se n�mero de itens do bmd for menor do que trinta:
                If colecaoBmds.Count <= 30 Then
                
                    'criando planilha "bmd + contador"
                    Worksheets("Bmd_Trinta").Copy after:=Worksheets(Sheets.Count)
                    ActiveSheet.Name = "Bmd" & contador
                    
                    With Worksheets("Bmd" & contador)
                    
                        .Activate
                        
                        Application.PrintCommunication = False ' essencial pois pagesetup fica muito lento sem esta linha, porem deve ser reativado logo apos o page setup
                        .PageSetup.FooterMargin = Application.InchesToPoints(0.1)
                        .PageSetup.RightFooter = "&I&9" & "Bmd N� " & numeroBmd & " (" & contador & " / " & colecaoChavesBmds.Count & ")"
                        Application.PrintCommunication = True ' essencial pois pagesetup fica muito lento sem esta linha, porem deve ser reativado logo apos o page setup
                        
                        'colando a colecaoTrinta na planilha criada
                        .Activate
                        colecaoParaPlanilhaBmd colecaoBmds, Worksheets("Bmd" & contador)
                    
                        .Unprotect
                        'cabecalho
                        .Cells(2, 7).Value = numeroBmd 'numero bmd
                        .Cells(5, 7).Value = numeroContrato 'numero contrato
                        .Cells(8, 7).Value = colecaoBmds(1)(2) 'municipio
                        .Cells(13, 3).Font.Color = RGB(255, 64, 0)
                        .Cells(13, 3).Value = "RASCUNHO" 'num_pedido
                        .Cells(11, 10).Value = contador & " / " & colecaoChavesBmds.Count 'sequencia
                        .Cells(10, 7).Value = colecaoBmds(1)(1) 'referencia
                         

                        'rodape
                        .Cells(49, 6).Value = primeiroDiaMes & " � " & ultimoDiaMes 'periodo
                        .Cells(47, 6).Value = colecaoBmds(1)(2) & " , " & Date 'local e data
                        .Protect
                        
                     End With
                     
                     
                Else
              
                
                        'criando planilha "bmd + contador"
                         Worksheets("Bmd").Copy after:=Worksheets(Sheets.Count)
                         ActiveSheet.Name = "Bmd" & contador
                        

                         'rodapezinho
                         With Worksheets("Bmd" & contador)
                            .Activate
                            Application.PrintCommunication = False ' essencial pois pagesetup fica muito lento sem esta linha, porem deve ser reativado logo apos o page setup
                                .PageSetup.FooterMargin = Application.InchesToPoints(0.1)
                                .PageSetup.RightFooter = "&I&9" & "Bmd N� " & numeroBmd & " (" & contador & " / " & colecaoChavesBmds.Count & ")"
                            Application.PrintCommunication = True ' essencial pois pagesetup fica muito lento sem esta linha, porem deve ser reativado logo apos o page setup
                         End With
                        
                         
                         
                        
                         
                         totalLinhas = colecaoBmds.Count + 14 + 14
                                                  
                         If totalLinhas < 58 Then
                          num_linhas_bmd = colecaoBmds.Count + 58 - totalLinhas
                         ElseIf totalLinhas = 58 Then
                          num_linhas_bmd = colecaoBmds.Count
                         Else
                         
                            resto = (totalLinhas) Mod 58
                            num_linhas_bmd = colecaoBmds.Count + 58 - resto
              

                         End If
                         
                             
                         
                         With Worksheets("Bmd" & contador)
                         
                             'cabecalho
                             .Cells(2, 7).Value = numeroBmd 'numero bmd
                             .Cells(5, 7).Value = numeroContrato 'numero contrato
                             .Cells(8, 7).Value = colecaoBmds(1)(2) 'municipio
                             .Cells(13, 3).Font.Color = RGB(255, 64, 0)
                             .Cells(13, 3).Value = "RASCUNHO" 'num_pedido
                             .Cells(11, 10).Value = contador & " / " & colecaoChavesBmds.Count 'sequencia
                             .Cells(10, 7).Value = colecaoBmds(1)(1) 'referencia
                         
                             quantidadePaginas = Application.WorksheetFunction.RoundUp((num_linhas_bmd + 14 + 14) / 58, 0)
                             quantidadeHeaders = quantidadePaginas - 1
                             
                             num_linhas_bmd = num_linhas_bmd - (quantidadeHeaders * 2)
                             
                            
                             'criando linhas conforme numero de linhas do bmd
                             .Range("15:15").Copy
                             .Range("15:15").Resize(num_linhas_bmd - 1).Insert
                             
                             'numerando as linhas sequencialmente
                             .Range("A15:A" & num_linhas_bmd + 14).DataSeries , xlDataSeriesLinear
                             
                             'colando a colecaoBmds na planilha criada
                             colecaoParaPlanilhaBmd colecaoBmds, Worksheets("Bmd" & contador)
    
                             .Unprotect
                                                          
                              'criando header para cada pagina adicional
                             CONT = 60
                             For pagina = 1 To quantidadePaginas - 1
                                .Range("13:14").Copy
                                .Rows(CONT).Insert
                                CONT = CONT + 58
                                num_linhas_bmd = num_linhas_bmd + 2
                             Next
                             
                             'somando valor total
                             .Range("I" & num_linhas_bmd + 15).Value = _
                                WorksheetFunction.Sum(.Range("I15:I" & num_linhas_bmd + 14))
                              
                             'rodape
                             .Cells(num_linhas_bmd + 17, 6).Value = primeiroDiaMes & " � " & ultimoDiaMes 'periodo
                             .Cells(num_linhas_bmd + 19, 6).Value = colecaoBmds(1)(2) & " , " & Date 'local e data
                             
                             'definindo quebra de pagina para otimizar area de impressao para impressoras diferentes, por definicao cada pagina tem 58 linhas
                             quantidadePaginas = Application.WorksheetFunction.RoundUp((num_linhas_bmd + 14 + 14) / 58, 0)
                             
                             num_linha = 60
                             For CONT = 1 To quantidadePaginas - 1
                                .Rows(num_linha).PageBreak = xlPageBreakManual
                                num_linha = num_linha + 58
                             Next
                             
                             .Protect
                         
                         End With
                         
                  End If

            contador = contador + 1

         Next

         With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .EnableAnimations = True
            .Calculation = xlCalculationAutomatic
         End With

         'devolve a colecaolinhas para a planilha boletim, agora com os dados da coluna num_bmd e sequencia
         colecaoParaPlanilha colecaoLinhas, Worksheets("Boletim")

         Worksheets("Bmd").Protect
         
         Worksheets("Itens_Boletim").Activate
         Worksheets("Boletim").Activate
         
         'MsgBox ("Processo GERAR BMD foi concluido!")
         
         gerarPdf ("PRE_BMD")
         
Exit Sub
TE:     'Tratamento de Erros

    MsgBox " Erro: " & Err.Description & Chr(13) & Chr(13) & "Local: UserForm UserFormBmd.btPreBmd_Click"

End Sub
Sub btGerarBmd_Click()
   
   On Error GoTo TE

   If planilhaVazia(Worksheets("Boletim")) Then

    Err.Raise ERRO_PLANILHA_VAZIA, "Planilha vazia!" _
                        , "Nenhum ITEM na planilha BOLETIM."

    End If
    
    qtdLinhas = Worksheets("Boletim").Range("A1").CurrentRegion.Rows.Count
    
    'verifica celulas vazias ou n�o numericas
    For Each celula In Worksheets("Boletim").Range("K2:M" & qtdLinhas)
        If IsEmpty(celula) Or Not IsNumeric(celula) Then
            Err.Raise vbObjectError + 50, "DADO INV�LIDO!" _
                            , "Alguma c�lula est� vazia ou contendo valor n�o num�rico!! Verifique colunas: NUM_BMD, SEQU�NCIA E NUM_PEDIDO."
        End If
    Next

    caminhoArquivo = Application.ActiveWorkbook.Path & "\registroLancamentosBmdsBD.csv"

    If Not arquivoExiste(caminhoArquivo) Then

        Err.Raise vbObjectError + 50, "ARQUIVO DE REGISTRO N�O ENCONTRADO!" _
                            , "Arquivo de registro de lan�amentos n�o foi encontrado!"

    End If

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .EnableAnimations = False
        .Calculation = xlCalculationManual
    End With

    'apaga todas as planilhas ap�s bmd
    Application.DisplayAlerts = False
    While Worksheets(Sheets.Count).Name <> "Bmd"
        Worksheets(Sheets.Count).Delete
    Wend
    Application.DisplayAlerts = True

    Application.PrintCommunication = False ' essencial pois pagesetup fica muito lento sem esta linha, porem deve ser reativado logo apos o page setup
    'configura margem do layout da planilha bmd para otimizar impressao em pdf
    With Worksheets("Bmd").PageSetup
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.1)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.2)
    End With
    Application.PrintCommunication = True ' essencial pois pagesetup fica muito lento sem esta linha, porem deve ser reativado logo apos o page setup
    'configura margem do layout da planilha bmd para otimizar impressao em pdf

    'classifica planilha boletim
    classificaPlanilha ("Boletim")


    Set colecaoLinhas = planilhaParaColecaoBmd(Worksheets("Boletim"))

    Set colecaoChavesBmds = New Collection
    Set colecaoBmds = New Collection
    Set colecaoTrinta = New Collection

    primeiroDiaMes = DateSerial(Year(colecaoLinhas(1)(1)), Month(colecaoLinhas(1)(1)), 1)
    ultimoDiaMes = DateSerial(Year(colecaoLinhas(1)(1)), Month(colecaoLinhas(1)(1)) + 1, 1) - 1

    numeroBmd = Worksheets("Boletim").Range("K2").Value

          'determinando as chaves dos bmds mes/ano + localizacao (jan-22TAMARANA)
          For Each elemento In colecaoLinhas
            If Not existeItemColecao(colecaoChavesBmds, elemento(1) & elemento(2)) Then
             colecaoChavesBmds.Add elemento(1) & elemento(2), elemento(1) & elemento(2)
            End If
          Next

        Worksheets("Bmd").Unprotect

        'para cada cidade gera um bmd
         contador = 1
         For Each chave In colecaoChavesBmds

               Set colecaoBmds = New Collection

               CONT = 1
               For Each linha In colecaoLinhas
                If linha(1) & linha(2) = chave Then

                     'adiciona as linhas da planilha boletim que pertencem a mesma cidade na colecaobmds
                     colecaoBmds.Add linha

                     'atribui ao campo sequencia o valor respectivo � cidade e preenche a coluna num_bmd
                     arrayAuxiliar = linha
                     arrayAuxiliar(11) = numeroBmd
                     arrayAuxiliar(12) = contador
                     colecaoLinhas.Add arrayAuxiliar, after:=CONT
                     colecaoLinhas.Remove (CONT)

                End If
                CONT = CONT + 1
               Next
               
               numeroContrato = Worksheets("Configura��es").Range("C15").Value
               
                'Se n�mero de itens do bmd for menor do que trinta:
                If colecaoBmds.Count <= 30 Then
                
                    'criando planilha "bmd + contador"
                    Worksheets("Bmd_Trinta").Copy after:=Worksheets(Sheets.Count)
                    ActiveSheet.Name = "Bmd" & contador
                    
                    With Worksheets("Bmd" & contador)
                    
                        .Activate
                        
                        Application.PrintCommunication = False ' essencial pois pagesetup fica muito lento sem esta linha, porem deve ser reativado logo apos o page setup
                        .PageSetup.FooterMargin = Application.InchesToPoints(0.1)
                        .PageSetup.RightFooter = "&I&9" & "Bmd N� " & numeroBmd & " (" & contador & " / " & colecaoChavesBmds.Count & ")"
                        Application.PrintCommunication = True ' essencial pois pagesetup fica muito lento sem esta linha, porem deve ser reativado logo apos o page setup
                        
                        'colando a colecaoTrinta na planilha criada
                        .Activate
                        colecaoParaPlanilhaBmd colecaoBmds, Worksheets("Bmd" & contador)
                    
                        .Unprotect
                        'cabecalho
                        .Cells(2, 7).Value = numeroBmd 'numero bmd
                        .Cells(5, 7).Value = numeroContrato 'numero contrato
                        .Cells(8, 7).Value = colecaoBmds(1)(2) 'municipio
                        .Cells(13, 3).Value = colecaoBmds(1)(13) 'num_pedido
                        .Cells(11, 10).Value = contador & " / " & colecaoChavesBmds.Count 'sequencia
                        .Cells(10, 7).Value = colecaoBmds(1)(1) 'referencia
                         

                        'rodape
                        .Cells(49, 6).Value = primeiroDiaMes & " � " & ultimoDiaMes 'periodo
                        .Cells(47, 6).Value = colecaoBmds(1)(2) & " , " & Date 'local e data
                        .Protect
                        
                     End With
                     
                     
                Else
              
                
                        'criando planilha "bmd + contador"
                         Worksheets("Bmd").Copy after:=Worksheets(Sheets.Count)
                         ActiveSheet.Name = "Bmd" & contador
                        

                         'rodapezinho
                         With Worksheets("Bmd" & contador)
                            .Activate
                            Application.PrintCommunication = False ' essencial pois pagesetup fica muito lento sem esta linha, porem deve ser reativado logo apos o page setup
                                .PageSetup.FooterMargin = Application.InchesToPoints(0.1)
                                .PageSetup.RightFooter = "&I&9" & "Bmd N� " & numeroBmd & " (" & contador & " / " & colecaoChavesBmds.Count & ")"
                            Application.PrintCommunication = True ' essencial pois pagesetup fica muito lento sem esta linha, porem deve ser reativado logo apos o page setup
                         End With
                        
                         
                         
                        
                         
                         totalLinhas = colecaoBmds.Count + 14 + 14
                                                  
                         If totalLinhas < 58 Then
                          num_linhas_bmd = colecaoBmds.Count + 58 - totalLinhas
                         ElseIf totalLinhas = 58 Then
                          num_linhas_bmd = colecaoBmds.Count
                         Else
                         
                            resto = (totalLinhas) Mod 58
                            num_linhas_bmd = colecaoBmds.Count + 58 - resto
              

                         End If
                         
                             
                         
                         With Worksheets("Bmd" & contador)
                         
                             'cabecalho
                             .Cells(2, 7).Value = numeroBmd 'numero bmd
                             .Cells(5, 7).Value = numeroContrato 'numero contrato
                             .Cells(8, 7).Value = colecaoBmds(1)(2) 'municipio
                             .Cells(13, 3).Value = colecaoBmds(1)(13) 'num_pedido
                             .Cells(11, 10).Value = contador & " / " & colecaoChavesBmds.Count 'sequencia
                             .Cells(10, 7).Value = colecaoBmds(1)(1) 'referencia
                         
                             quantidadePaginas = Application.WorksheetFunction.RoundUp((num_linhas_bmd + 14 + 14) / 58, 0)
                             quantidadeHeaders = quantidadePaginas - 1
                             
                             num_linhas_bmd = num_linhas_bmd - (quantidadeHeaders * 2)
                             
                            
                             'criando linhas conforme numero de linhas do bmd
                             .Range("15:15").Copy
                             .Range("15:15").Resize(num_linhas_bmd - 1).Insert
                             
                             'numerando as linhas sequencialmente
                             .Range("A15:A" & num_linhas_bmd + 14).DataSeries , xlDataSeriesLinear
                             
                             'colando a colecaoBmds na planilha criada
                             colecaoParaPlanilhaBmd colecaoBmds, Worksheets("Bmd" & contador)
    
                             .Unprotect
                                                          
                              'criando header para cada pagina adicional
                             CONT = 60
                             For pagina = 1 To quantidadePaginas - 1
                                .Range("13:14").Copy
                                .Rows(CONT).Insert
                                CONT = CONT + 58
                                num_linhas_bmd = num_linhas_bmd + 2
                             Next
                             
                             'somando valor total
                             .Range("I" & num_linhas_bmd + 15).Value = _
                                WorksheetFunction.Sum(.Range("I15:I" & num_linhas_bmd + 14))
                              
                             'rodape
                             .Cells(num_linhas_bmd + 17, 6).Value = primeiroDiaMes & " � " & ultimoDiaMes 'periodo
                             .Cells(num_linhas_bmd + 19, 6).Value = colecaoBmds(1)(2) & " , " & Date 'local e data
                             
                             'definindo quebra de pagina para otimizar area de impressao para impressoras diferentes, por definicao cada pagina tem 58 linhas
                             quantidadePaginas = Application.WorksheetFunction.RoundUp((num_linhas_bmd + 14 + 14) / 58, 0)
                             
                             num_linha = 60
                             For CONT = 1 To quantidadePaginas - 1
                                .Rows(num_linha).PageBreak = xlPageBreakManual
                                num_linha = num_linha + 58
                             Next
                             
                             .Protect
                         
                         End With
                         
                  End If

            contador = contador + 1

         Next

         With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .EnableAnimations = True
            .Calculation = xlCalculationAutomatic
         End With

         'devolve a colecaolinhas para a planilha boletim, agora com os dados da coluna num_bmd e sequencia
         colecaoParaPlanilha colecaoLinhas, Worksheets("Boletim")

         Worksheets("Bmd").Protect
         
         Worksheets("Itens_Boletim").Activate
         Worksheets("Boletim").Activate
         
         'MsgBox ("Processo GERAR BMD foi concluido!")
         
         gerarPdf ("BOLETIM")
      

Exit Sub
TE:     'Tratamento de Erros

    MsgBox " Erro: " & Err.Description & Chr(13) & Chr(13) & "Local: UserForm UserFromBmd.btGerarBmd_Click"

End Sub

Sub gerarPdf(ByVal termo As String)

    
    Dim caminhoArquivo As String
    
    caminhoArquivo = Application.ActiveWorkbook.Path & "\" & termo & _
                  Format(Now, "-dd-mmm-yyyy-h-mm-ss") & ".pdf"
                  
                  contadorDecrescente = Sheets.Count
                  contadorCrescente = 0
                  Dim arrayPlanilhas() As String
                  While Worksheets(contadorDecrescente).Name <> "Bmd"
                  
                    ReDim Preserve arrayPlanilhas(contadorCrescente)
                    arrayPlanilhas(contadorCrescente) = Worksheets(contadorDecrescente).Name
                    contadorDecrescente = contadorDecrescente - 1
                    contadorCrescente = contadorCrescente + 1
                    
                  Wend
                  
                  If contadorCrescente = 0 Then
                    MsgBox ("Nenhum BMD para imprimir!")
                  Else
                 
                    Sheets(arrayPlanilhas).Select
                    
                    
                    
                    'exporta as planilha selecionadas para pdf no caminhoarquivo e abre o aquivo
                    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=caminhoArquivo, _
                        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
                  

                  End If
    
        Worksheets("Boletim").Activate

End Sub
Sub colecaoParaPlanilhaBmd(ByVal colecao As Collection, ByVal ws As Worksheet)

 
 'caso haja linhas na cole��o de linhas tranporta seus valores para um array bidimensional arrayResultado
 If colecao.Count > 0 Then

     ReDim arrayResultado(1 To colecao.Count, 1 To UBound(colecao(1)) - 1) As Variant

     For linha = 1 To UBound(arrayResultado, 1)

        For Col = 3 To UBound(arrayResultado, 2) + 1
            arrayResultado(linha, Col - 2) = colecao(linha)(Col)
        Next

     Next
  
     ws.Unprotect
     ws.Range("B15:I" & colecao.Count + 14).Value = arrayResultado
     ws.Protect



 End If
End Sub
Sub colecaoParaPlanilhaPepPendencias(ByVal colecao As Collection, ByVal ws As Worksheet)

 ws.Activate
  
 'caso haja linhas na cole��o de linhas transporta seus valores para um array bidimensional arrayResultado
 If colecao.Count > 0 Then

     ReDim arrayResultado(1 To colecao.Count, 1 To UBound(colecao(1))) As Variant

     For linha = 1 To UBound(arrayResultado, 1)

        For Col = 1 To UBound(arrayResultado, 2)
            arrayResultado(linha, Col) = colecao(linha)(Col)
        Next

     Next
 
     numeroLinhasPlanilhaPendencia = Worksheets("Pendencias").Range("A1").CurrentRegion.Rows.Count
     
     ws.Range(Cells(numeroLinhasPlanilhaPendencia + 1, 1), Cells(colecao.Count + numeroLinhasPlanilhaPendencia, 3)).Value = arrayResultado
     ws.Columns.AutoFit

 End If
 
End Sub
Function planilhaParaColecaoBmd(ByVal ws As Worksheet)


     'Transporta os valores de todas as celulas da planilha para um array, exceto colunas ocultas
      arrayPlanilhaInteira = ws.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)

     'Foi criado uma cole��o de linhas para poder remover linhas depois
     Set planilhaParaColecaoBmd = New Collection

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

        planilhaParaColecaoBmd.Add (arrayLinha)

     Next
     
     
End Function


Private Sub btSalvarLancamentos_Click()

    On Error GoTo TE
    
    qtdLinhas = Worksheets("Boletim").Range("A1").CurrentRegion.Rows.Count
    
    'verifica celulas vazias ou n�o numericas
    For Each celula In Worksheets("Boletim").Range("K2:M" & qtdLinhas)
        If IsEmpty(celula) Or Not IsNumeric(celula) Then
            Err.Raise vbObjectError + 50, "DADO INV�LIDO!" _
                            , "Alguma c�lula est� vazia ou contendo valor n�o num�rico!! Verifique colunas: NUM_BMD, SEQU�NCIA E NUM_PEDIDO."
        End If
    Next
    
    'verifica celulas vazias ou n�o numericas
    For Each celula In Worksheets("Boletim").Range("C2:C" & qtdLinhas)
        If IsEmpty(celula) Or Not IsNumeric(celula) Then
            Err.Raise vbObjectError + 50, "DADO INV�LIDO!" _
                            , "Alguma c�lula est� vazia ou contendo valor n�o num�rico!! Verifique coluna: NUM_OS."
        End If
    Next
    
    'verifica se todos os num_bmd s�o iguais ao ultimo numero bmd do arquivo de registro + 1
    If Not WorksheetFunction.CountIf(Worksheets("Boletim").Range("K2:K" & qtdLinhas), consultaUltimoBmd + 1) = qtdLinhas - 1 Then
     Err.Raise vbObjectError + 50, "N�MERO BMD INV�LIDO!" _
                            , "Algum n�mero de BMD n�o segue a sequ�ncia do arquivo de registros de lan�amentos!"
    End If
   
    
    caminhoArquivo = Application.ActiveWorkbook.Path & "\registroLancamentosBmdsBD.csv"
    
    'verifica se existe os ja pagas na planilha boletim antes de salvar no arquivo de registros
    arraycolunaNumOs = Worksheets("Boletim").Range("C2:C" & qtdLinhas)
    For Each elemento In arraycolunaNumOs
        If termoExisteNoArquivo(caminhoArquivo, 3, elemento) Then
           Err.Raise vbObjectError + 50, "OS J� LAN�ADA!" _
                            , "OS " & elemento & " consta como j� lan�ada no arquivo de registros!! O processo de salvamento foi cancelado!"
        End If
    Next
    
    'verifica se existe num pedidos ja pagos na planilha boletim antes de salvar no arquivo de registros
    arraycolunaNumPedido = Worksheets("Boletim").Range("M2:M" & qtdLinhas)
    For Each elemento In arraycolunaNumPedido
        If termoExisteNoArquivo(caminhoArquivo, 4, elemento) Then
           Err.Raise vbObjectError + 50, "PEDIDO J� LAN�ADO!" _
                            , "PEDIDO " & elemento & " consta como j� lan�ado no arquivo de registros!!  O processo de salvamento foi cancelado!"
        End If
    Next
    

    Worksheets("Registros_Bmds").Activate
    Worksheets("Registros_Bmds").Cells.Clear
    
    
    'fazendo copia da coluna num_os da planilha boletim para registros_bmds
    arraycolunas = Worksheets("Boletim").Range("C1:C" & qtdLinhas)
    Worksheets("Registros_Bmds").Range("C1:C" & qtdLinhas).Value = arraycolunas
    
    'fazendo copia da coluna num_bmd da planilha boletim para registros_bmds
    arraycolunas = Worksheets("Boletim").Range("K1:K" & qtdLinhas)
    Worksheets("Registros_Bmds").Range("A1:A" & qtdLinhas).Value = arraycolunas
    
    'fazendo copia da coluna sequencia da planilha boletim para registros_bmds
    arraycolunas = Worksheets("Boletim").Range("L1:L" & qtdLinhas)
    Worksheets("Registros_Bmds").Range("B1:B" & qtdLinhas).Value = arraycolunas
    
    'fazendo copia da coluna num_pedido da planilha boletim para registros_bmds
    arraycolunas = Worksheets("Boletim").Range("M1:M" & qtdLinhas)
    Worksheets("Registros_Bmds").Range("D1:D" & qtdLinhas).Value = arraycolunas
    
    'remove os duplicadas
    Worksheets("Registros_Bmds").Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(3), Header:=xlYes
      
    Worksheets("Registros_Bmds").Columns.AutoFit
    
    
    'SALVA CONTEUDO DA PLANILHA REGISTROS_BMD NO ARQUIVO CSV DE REGISTRO DE LANCAMENTOS
    
    caminhoArquivo = Application.ActiveWorkbook.Path & "\registroLancamentosBmdsBD.csv"
    
    Set colecaoLinhas = planilhaParaColecao(Worksheets("Registros_Bmds"))

    If Not arquivoExiste(caminhoArquivo) Then
        MsgBox "ARQUIVO DE REGISTROS N�O ENCONTRADO!! REGISTRO DE LAN�AMENTOS N�O FORAM SALVOS!"
    Else
     If arquivoRegistroEhFormatoValido(caminhoArquivo) Then
         salvaColecaoDeLinhasNoCsvRegistro caminhoArquivo, colecaoLinhas
         MsgBox "LAN�AMENTOS SALVOS COM SUCESSO!!"
        Else
         MsgBox "CABE�ALHO DO ARQUIVO DE REGISTROS � INV�LIDO!! REGISTRO DE LAN�AMENTOS N�O FORAM SALVOS!"
        End If
    End If
    
 
Exit Sub
TE:     'Tratamento de Erros

    MsgBox " Erro: " & Err.Description & Chr(13) & Chr(13) & "Local: M�dulo sanearCsv.unificaServicosOs"

End Sub
Function existeOsPagaLancamento()

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
    
Exit Function
TE:     'Tratamento de Erros

    MsgBox " Erro: " & Err.Description & Chr(13) & Chr(13) & "Local: M�dulo sanearCsv.filtraOsPaga"

End Function
Sub nenhumaCelulaVaziaENaoNumerica()
    For Each celula In myRange '
        c = c + 1
        If IsEmpty(myCell) Then
            myCell.Interior.Color = RGB(255, 87, 87)
            i = i + 1
        End If
    Next myCell
End Sub
Private Sub UserForm_Initialize()

      'imprimi conteudo da planilha pendencias na textbox do userformbmd
       escrevePendencias
      
      'posiciona o userform no centro da tela do excel, util qdo usa dois monitores
      Me.StartUpPosition = 0
      Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
      Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
      
     ListBox1.SetFocus
      
  
End Sub

