Attribute VB_Name = "recriaPlanilhas"
Sub recriaPlanilhas()

 If MsgBox("TODOS OS DADOS SER�O APAGADOS!! DESEJA PROSSEGUIR?", vbYesNo, "ATEN��O!!") = vbYes Then
    
    Application.DisplayAlerts = False

    
    Worksheets("Os").Delete
    Worksheets("Servicos").Delete
    
    
    Application.DisplayAlerts = True
    
   
    
    Sheets.Add
    Sheets(1).Name = "Os"
     
    Sheets.Add after:=Sheets(1)
    Sheets(2).Name = "Servicos"
    
                
    userFormPrincipal.textboxOs.Text = ""
    userFormPrincipal.textboxServicos.Text = ""
    

 End If
 
End Sub
