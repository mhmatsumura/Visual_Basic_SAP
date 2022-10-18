Attribute VB_Name = "processoPrincipal"
'A O CLICAR NO BOTAO "CARREGAR ARQUIVO CSV" ESTE PROCESSO EH CHAMADO

Sub processo_Principal()

    sanearCsv.sanear_Csv
    
End Sub

Sub carregaUserFormBmd()

    UserFormBmd.Show
 
End Sub

Sub carregaUserFormPrincipal()

    userFormPrincipal.Show
 
End Sub

