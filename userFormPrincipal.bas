
Private Sub botaApagaRecria_Click()

    recriaPlanilhas.recriaPlanilhas
 
    'apaga todas as planilhas ap�s bmd
    Application.DisplayAlerts = False
    While Worksheets(Sheets.Count).Name <> "Bmd"
        Worksheets(Sheets.Count).Delete
    Wend
    Application.DisplayAlerts = True
    
End Sub

Private Sub botaoCsvOs_Click()
    carregarArquivoOs.carregar_Os ("Os")
End Sub

Private Sub botaoCsvServicos_Click()

 carregarArquivoServicos.carregar_Servicos ("Servicos")
 
End Sub

Private Sub botaoGerarBoletim_Click()

  sanearCsv.mesAno = userFormPrincipal.ComboBoxMes.Value & "-" & userFormPrincipal.TextBoxAno.Value
  
  Unload userFormPrincipal
  
  resetaPlanilhas.reseta_Planilhas
  
  'carregar_Registros_Bmds
  
  processoPrincipal.processo_Principal
 
End Sub

Private Sub UserForm_Initialize()

 
    
    inicializaDicionarioCidades.inicializa_Dicionario_Cidades
    
    ComboBoxMes.AddItem "JAN"
    ComboBoxMes.AddItem "FEV"
    ComboBoxMes.AddItem "MAR"
    ComboBoxMes.AddItem "ABR"
    ComboBoxMes.AddItem "MAI"
    ComboBoxMes.AddItem "JUN"
    ComboBoxMes.AddItem "JUL"
    ComboBoxMes.AddItem "AGO"
    ComboBoxMes.AddItem "SET"
    ComboBoxMes.AddItem "OUT"
    ComboBoxMes.AddItem "NOV"
    ComboBoxMes.AddItem "DEZ"
    
    'seta como padrao mes anterior
    'ComboBoxMes.ListIndex = Format(Date, "m") - 2
    ComboBoxMes.ListIndex = 0
    
    TextBoxAno.MaxLength = 2
    TextBoxAno.Text = Format(Date, "yy")
    

'Start Userform Centered inside Excel Screen (for dual monitors)
  Me.StartUpPosition = 0
  Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
  Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub
'aceita apenas numeros
Private Sub TextBoxAno_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
         KeyAscii = 0
    End If
End Sub

