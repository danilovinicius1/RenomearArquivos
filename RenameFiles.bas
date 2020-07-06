Sub Renomear_Arquivo()
Dim nAntigo As String
Dim nNovo As String
Dim varNovoNome As String
Dim varAntigoNome As String
            
    For i = 1 To Range("A:A").End(xlDown).Row
    
    varAntigoNome = ActiveSheet.Cells(i, 2)
    
    varNovoNome = ActiveSheet.Cells(i, 3)
    
    nAntigo = "D:\TEMP\" & varAntigoNome
    
    nNovo = "D:\TEMP\" & varNovoNome

    Name nAntigo As nNovo
           
    Next
End Sub
