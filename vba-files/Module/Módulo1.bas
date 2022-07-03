Attribute VB_Name = "Módulo1"
Sub teste()
Attribute teste.VB_ProcData.VB_Invoke_Func = " \n14"

    Worksheets("base_de_ticks_1").Shapes("Picture 45").Copy
    Worksheets(1).Paste Range("A1")
    
    
End Sub
