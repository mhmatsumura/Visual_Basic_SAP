Attribute VB_Name = "inicializaDicionarioCidades"
Public dicionarioCidades As Scripting.Dictionary
'INSERE DADOS DO CAR E DOS NOMES DAS CIDADES NO DICIONARIOCIDADES A PARTIR DA PLANILHA CIDADES
Sub inicializa_Dicionario_Cidades()
    
     Set dicionarioCidades = New Scripting.Dictionary
    
    'ABRE A PLANILHA "Col_Interesse"
     Worksheets("Cidades").Activate

     UltimaLinha = Worksheets("Cidades").Cells(Cells.Rows.Count, 1).End(xlUp).Row
 
     For Line = 1 To UltimaLinha
        dicionarioCidades(Cells(Line, 1).Text) = Cells(Line, 2).Text
     Next
    
End Sub
'CONSULTA CIDADE NO DICIONARIO , PROCESSO MAIS RAPIDO DO QUE CONSULTAR NA PLANILHA
Function consultaCidadeDicionario(ByVal car)

    If dicionarioCidades.Exists(car) Then
     consultaCidadeDicionario = dicionarioCidades(car)
    Else
     consultaCidadeDicionario = ""
    End If
    
End Function
