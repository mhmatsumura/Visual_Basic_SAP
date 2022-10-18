Attribute VB_Name = "resetaPlanilhas"
Sub reseta_Planilhas()

    Worksheets("Registros_Bmds").Cells.Clear
    Worksheets("Unificado").Cells.Clear
    Worksheets("Pendencias").Cells.Clear
    Worksheets("Col_Interesse").Cells.Clear
    Worksheets("Car").Cells.Clear
    Worksheets("Atividade").Cells.Clear
    Worksheets("Itens_Boletim").Cells.Clear
    Worksheets("Boletim").Cells.Clear
    
    'apaga todas as planilhas ap�s bmd
    Application.DisplayAlerts = False
    While Worksheets(Sheets.Count).Name <> "Bmd"
        Worksheets(Sheets.Count).Delete
    Wend
    Application.DisplayAlerts = True
    
End Sub
